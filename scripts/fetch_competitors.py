#!/usr/bin/env python3
"""Build a geocoded list of home-visit medical care providers (競合) in Kanto.

Source: 関東信越厚生局 「施設基準の届出状況（全体）（届出受理医療機関名簿）」 医科 Excel ZIP
  page : https://kouseikyoku.mhlw.go.jp/kantoshinetsu/chousa/kijyun.html
  zip  : https://kouseikyoku.mhlw.go.jp/kantoshinetsu/shisetsu_ika_rMMYY.zip
  legend: https://kouseikyoku.mhlw.go.jp/kantoshinetsu/ryakushoMMYY.pdf

Codes (受理記号, confirmed from the legend PDF, R8 revision):
  支援診１ / 支援病１            -> kyoka_tandoku (機能強化型 単独型)
  支援診２ア / 支援診２イ / 支援病２ -> kyoka_renkei  (機能強化型 連携型)
  支援診３ / 支援病３            -> zaishi        (従来型 在支診/在支病)
  在医総管１ (and 在医総管２/３)   -> general       (在宅時医学総合管理料 only)

Steps:
  python3 scripts/fetch_competitors.py            # download + extract + geocode
  python3 scripts/fetch_competitors.py --no-geocode
Outputs: data/work/competitors_raw.csv, data/work/competitors.csv,
         data/work/geocode_cache.json (resumable).
"""
import argparse
import json
import re
import subprocess
import sys
import time
import unicodedata
import urllib.parse
import zipfile
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw" / "kouseikyoku"
WORK = ROOT / "data" / "work"
BASE = "https://kouseikyoku.mhlw.go.jp"
PAGE = BASE + "/kantoshinetsu/chousa/kijyun.html"

PREFS = {"08": "茨城県", "09": "栃木県", "10": "群馬県", "11": "埼玉県",
         "12": "千葉県", "13": "東京都", "14": "神奈川県"}

TYPE_RANK = {"kyoka_tandoku": 4, "kyoka_renkei": 3, "zaishi": 2, "general": 1}


def code_type(code: str):
    c = unicodedata.normalize("NFKC", str(code or "")).strip()
    if c in ("支援診1", "支援病1"):
        return "kyoka_tandoku"
    if c.startswith("支援診2") or c == "支援病2":
        return "kyoka_renkei"
    if c in ("支援診3", "支援病3"):
        return "zaishi"
    if c.startswith("在医総管"):
        return "general"
    return None


def curl(url, out=None):
    cmd = ["curl", "-sS", "-f", "--retry", "3", "--max-time", "300", url]
    if out:
        cmd += ["-o", str(out)]
        subprocess.run(cmd, check=True)
        return None
    return subprocess.run(cmd, check=True, capture_output=True).stdout


def download():
    RAW.mkdir(parents=True, exist_ok=True)
    html = curl(PAGE).decode("utf-8", "ignore")
    (RAW / "kijyun.html").write_text(html, encoding="utf-8")
    m = re.search(r'href="(/kantoshinetsu/shisetsu_ika_r\d+\.zip)"', html)
    if not m:
        sys.exit("医科 ZIP link not found on " + PAGE)
    zurl = BASE + m.group(1)
    zpath = RAW / Path(m.group(1)).name
    lg = re.search(r'href="(/kantoshinetsu/ryakusho\d+\.pdf)"', html)
    asof = re.search(r"(令和\s*\d+年\s*\d+月\s*\d+日)現在", html)
    if not zpath.exists():
        print("downloading", zurl)
        curl(zurl, zpath)
    if lg and not (RAW / Path(lg.group(1)).name).exists():
        curl(BASE + lg.group(1), RAW / Path(lg.group(1)).name)
    print("source:", zurl, "| as of:", asof.group(1) if asof else "?")
    return zpath


def extract_files(zpath):
    """Return {pref_code: xlsx path}. Zip names are Shift-JIS mangled, so key on the 2-digit prefix."""
    out_dir = RAW / "ika"
    files = {}
    with zipfile.ZipFile(zpath) as z:
        for info in z.infolist():
            name = Path(info.filename).name
            pc = name[:2]
            if pc in PREFS and name.lower().endswith((".xlsx", ".xls")):
                dest = out_dir / f"{pc}_ika{Path(name).suffix}"
                if not dest.exists():
                    dest.parent.mkdir(parents=True, exist_ok=True)
                    dest.write_bytes(z.read(info))
                files[pc] = dest
    missing = set(PREFS) - set(files)
    if missing:
        sys.exit(f"missing prefectures in zip: {missing}")
    return files


def parse_beds(s):
    s = unicodedata.normalize("NFKC", str(s or ""))
    nums = [int(x) for x in re.findall(r"\d+", s)]
    return sum(nums) if nums else 0


def build_raw(files):
    rows = []
    for pc, path in sorted(files.items()):
        df = pd.read_excel(path, header=3, dtype=str).fillna("")
        df = df[df["医療機関番号"].str.strip() != ""]
        df["_type"] = df["受理記号"].map(code_type)
        hits = df[df["_type"].notna()]
        n_pref = 0
        for num, g in df[df["医療機関番号"].isin(hits["医療機関番号"])].groupby("医療機関番号", sort=False):
            h = g[g["_type"].notna()]
            best = max(h["_type"], key=lambda t: TYPE_RANK[t])
            codes = sorted(set(unicodedata.normalize("NFKC", c) for c in h["受理記号"]))
            first = g.iloc[0]
            beds = parse_beds(first["病床数"])
            if any(c.startswith("支援病") for c in codes):
                is_hosp = "病院"
            elif any(c.startswith("支援診") for c in codes):
                is_hosp = "診療所"
            else:
                is_hosp = "病院" if beds >= 20 else "診療所"
            rows.append({
                "pref": PREFS[pc],
                "inst_code": f"{pc}1{num.strip()}",   # 10-digit 医療機関コード (県2+点数表1+番号7)
                "inst_no": num.strip(),
                "name": first["医療機関名称"].strip(),
                "address": first["医療機関所在地（住所）"].strip(),
                "postal": first["医療機関所在地（郵便番号）"].strip(),
                "is_hospital": is_hosp,
                "type": best,
                "codes": "|".join(codes),
                "beds": beds if beds else "",
                "beds_detail": re.sub(r"\s+", " ", unicodedata.normalize("NFKC", first["病床数"])).strip(),
                "tel": first["電話番号"].strip(),
            })
            n_pref += 1
        print(PREFS[pc], n_pref)
    out = pd.DataFrame(rows)
    WORK.mkdir(parents=True, exist_ok=True)
    out.to_csv(WORK / "competitors_raw.csv", index=False, encoding="utf-8-sig")
    return out


# ---------------- geocoding ----------------
GSI = "https://msearch.gsi.go.jp/address-search/AddressSearch?q="
NUM_TAIL = r"\d+(?:(?:丁目|番地の|番地|番|号|の|-)\d+)*(?:丁目|番地|番|号)?"


def clean_address(pref, addr):
    a = unicodedata.normalize("NFKC", addr)
    a = re.sub(r"〒?\s*\d{3}-?\d{4}", "", a)
    a = re.sub(r"[‐‑‒–—―−ｰー－](?=\d)", "-", a)        # dash variants between digits
    a = re.sub(r"(?<=\d)[‐‑‒–—―−ｰー－]", "-", a)
    a = re.sub(r"\s+", " ", a).strip()  # keep spaces as a boundary so "3-6-1 1F" is not merged
    a = re.sub(r"[（(].*?[）)]", "", a)
    # keep up to the first house-number block; drops building names / floors / 号室
    m = re.search(NUM_TAIL, a)
    if m:
        a = a[:m.end()]
    a = a.replace(" ", "")
    a = re.sub(r"[-の]$", "", a)
    if not a.startswith(pref):
        a = pref + a
    return a


def truncate_address(a):
    """Cut to 丁目 / first number level (e.g. 宇都宮市小幡2-5-15 -> 宇都宮市小幡2)."""
    m = re.search(r"^(.*?\d+丁目)", a)
    if m:
        return m.group(1)
    m = re.search(r"^(.*?[^\d-]\d+)(?=[-番の])", a)
    if m:
        return m.group(1)
    m = re.search(r"^(.*?)\d", a)
    if m and len(m.group(1)) > 4:
        return m.group(1)
    return None


class Geocoder:
    def __init__(self, path, rate=5.0):
        self.path = path
        self.cache = json.loads(path.read_text()) if path.exists() else {}
        self.min_dt = 1.0 / rate
        self.last = 0.0
        self.dirty = 0

    def query(self, q):
        if q in self.cache:
            return self.cache[q]
        dt = time.time() - self.last
        if dt < self.min_dt:
            time.sleep(self.min_dt - dt)
        self.last = time.time()
        res = None
        for attempt in range(4):
            try:
                body = curl(GSI + urllib.parse.quote(q))
                data = json.loads(body or b"[]")
                if data:
                    f0 = data[0]
                    lon, lat = f0["geometry"]["coordinates"][:2]
                    res = {"lon": lon, "lat": lat, "title": f0["properties"].get("title", "")}
                break
            except Exception as e:  # noqa
                time.sleep(2 * (attempt + 1))
        else:
            return None  # network failure: do not cache
        self.cache[q] = res
        self.dirty += 1
        if self.dirty >= 50:
            self.save()
        return res

    def save(self):
        tmp = self.path.with_suffix(".tmp")
        tmp.write_text(json.dumps(self.cache, ensure_ascii=False))
        tmp.replace(self.path)
        self.dirty = 0


def geocode(df):
    gc = Geocoder(WORK / "geocode_cache.json")
    lats, lons, quals, queries = [], [], [], []
    for i, r in enumerate(df.itertuples()):
        q = clean_address(r.pref, r.address)
        res, qual = gc.query(q), "exact"
        if not res:
            t = truncate_address(q)
            if t and t != q:
                res, qual = gc.query(t), "truncated"
        if res:
            lats.append(res["lat"]); lons.append(res["lon"])
        else:
            lats.append(None); lons.append(None); qual = "failed"
        quals.append(qual); queries.append(q)
        if i % 500 == 0:
            print(f"geocode {i}/{len(df)}", flush=True)
    gc.save()
    df = df.copy()
    df["geocode_query"] = queries
    df["lat"], df["lon"], df["geocode_quality"] = lats, lons, quals
    df.to_csv(WORK / "competitors.csv", index=False, encoding="utf-8-sig")
    print(df["geocode_quality"].value_counts().to_string())
    return df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-geocode", action="store_true")
    args = ap.parse_args()
    zpath = download()
    files = extract_files(zpath)
    df = build_raw(files)
    print("total", len(df))
    print(df.groupby(["pref", "type"]).size().unstack(fill_value=0).to_string())
    if not args.no_geocode:
        geocode(df)


if __name__ == "__main__":
    main()
