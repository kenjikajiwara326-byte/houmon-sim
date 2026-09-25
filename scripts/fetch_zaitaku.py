#!/usr/bin/env python3
"""Fetch current home-visit medical care (訪問診療) supply statistics for Kanto.

1) 厚生労働省「在宅医療にかかる地域別データ集」(1,741 市区町村別)
   Page : https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000061944.html
   File : the xlsx linked as 「在宅医療にかかる地域別データ集」 (as of 2026-09: /content/10800000/001681907.xlsx,
          "令和６年度分データを更新")
   Sheets used:
     「R５年」 : 医療施設調査（静態, 令和5年10月1日）特別集計 — 訪問診療/往診/看取り の実施施設数と
                 実施件数（令和5年9月の1か月間）, 医療機関所在地ベース.  This is the latest static survey
                 (the R6年 sheet has no 医療施設調査 items).
     「R6年」  : 在支診/在支病 届出数 (厚生局, 2024-03-31), 訪問看護ST (2024-10-01), 人口 (2024-01-01).
   Output: data/work/zaitaku_muni.csv  (Kanto 08-14, 316 rows; 政令市は市単位, 東京23区は区単位)

2) 社会医療診療行為別統計 (NDB-based, e-Stat toukei=00450048) 第22表 医科診療（総数－１総数）
   実施件数・回数 by 診療行為(細分類) — national monthly totals of 在宅患者訪問診療料 / 在医総管 / 施設総管.
   Latest edition (令和7年8月審査分) and 令和5年6月審査分 (for comparison with 医療施設調査 R5.9).
   Output: data/work/ndb_national.txt

Usage: python3 scripts/fetch_zaitaku.py [--force]
"""
import html
import re
import subprocess
import sys
import time
import unicodedata
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data/raw/zaitaku"
WORK = ROOT / "data/work"
PREFS = {"08", "09", "10", "11", "12", "13", "14"}
FORCE = "--force" in sys.argv
VISITS_PER_PATIENT = 1.89  # NDB 訪問診療料 延べ回数 / レセプト件数 (national)

MHLW_PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/0000061944.html"
ESTAT = "https://www.e-stat.go.jp"
ESTAT_TOP = ESTAT + "/stat-search/files?page=1&toukei=00450048&tstat=000001029602"


def fetch(url: str, dest: Path, force: bool = FORCE) -> Path:
    if dest.exists() and dest.stat().st_size > 0 and not force:
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["curl", "-sSfL", "-o", str(dest), url], check=True)
    time.sleep(0.5)
    return dest


def text(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


# ---------------------------------------------------------------- 1. 地域別データ集
def fetch_chiikibetsu() -> Path:
    page = fetch(MHLW_PAGE, RAW / "mhlw_0000061944.html", force=True)
    m = re.search(r'href="([^"]+\.xlsx)"[^>]*>\s*在宅医療にかかる地域別データ集', text(page))
    if not m:
        raise SystemExit("xlsx link for 在宅医療にかかる地域別データ集 not found")
    url = m.group(1)
    if url.startswith("/"):
        url = "https://www.mhlw.go.jp" + url
    print("地域別データ集:", url)
    return fetch(url, RAW / Path(url).name)


def read_sheet(xlsx: Path, sheet: str) -> pd.DataFrame:
    """Return DataFrame indexed by 5-digit code, columns keyed by 項目番号 (int)."""
    raw = pd.read_excel(xlsx, sheet_name=sheet, header=None)
    # layout: row 0 = 項目番号 (cols 7..), row 1 = データ時点, rows 2-5 = labels, row 6 = 全国計,
    # row 7 = units / column labels, rows 8.. = municipalities.
    # cols: 1 都道府県コード, 2 二次医療圏コード, 3 市区町村コード (no check digit), 4 県, 5 市区町村, 6 区分
    assert "市区町村コード" in str(raw.iat[7, 3]), "unexpected layout"
    body = raw.iloc[8:].copy()
    body = body[body[3].notna()]
    out = pd.DataFrame({
        "code": body[3].astype(int).astype(str).str.zfill(5),
        "pref": body[4].astype(str).str.strip(),
        "name": body[5].astype(str).str.strip(),
        "kind": body[6].astype(str).str.strip(),
        "iryoken": body[2].astype(str).str.strip().str.replace(r"\.0$", "", regex=True).str.zfill(4).where(lambda s: s.str.isdigit(), ""),
    })
    for j in range(7, raw.shape[1]):
        n = raw.iat[0, j]
        if isinstance(n, (int, float)) and not pd.isna(n):
            out[int(n)] = pd.to_numeric(body[j], errors="coerce")
    return out.set_index("code")


def build_muni(xlsx: Path) -> pd.DataFrame:
    sheets = pd.ExcelFile(xlsx).sheet_names
    s5 = next(s for s in sheets if s.startswith("R") and ("５" in s or "5" in s) and "年" in s and "(" not in s)
    s6 = next(s for s in sheets if s.startswith("R") and ("６" in s or "6" in s) and "年" in s and "(" not in s)
    r5 = read_sheet(xlsx, s5)
    r6 = read_sheet(xlsx, s6)
    k = r5[r5.index.str[:2].isin(PREFS)]
    k6 = r6.reindex(k.index)
    df = pd.DataFrame(index=k.index)
    df["pref"] = k["pref"]
    df["name"] = k["name"]
    df["kind"] = k["kind"]
    df["iryoken"] = k["iryoken"]
    df["pop_2024"] = k6[1]
    df["pop65_2024"] = k6[2]
    # --- 医療施設調査 R5 (施設数 as of 2023-10-01; 件数 = 2023年9月中の実施件数; 医療機関所在地ベース)
    df["visit_cases"] = k[14] + k[20]            # 訪問診療 実施件数 病院+診療所 (件/月, R5.9)
    # rough monthly patients: NDB 令和5年6月審査分 訪問診療料 回数/件数 = 1.89 (see ndb_national.txt)
    df["visit_patients_est"] = (df["visit_cases"] / VISITS_PER_PATIENT).round(1)
    df["visit_cases_hosp"] = k[14]
    df["visit_cases_clinic"] = k[20]
    df["visit_cases_clinic_zaishi"] = k[22]      # うち在支診
    df["visit_hospitals"] = k[13]                # 訪問診療を実施する病院数
    df["visit_hospitals_zaishi"] = k[15]
    df["visit_clinics"] = k[19]                  # 訪問診療を実施する一般診療所数
    df["visit_clinics_zaishi"] = k[21]
    df["oushin_cases"] = k[26] + k[32]           # 往診 実施件数 (件/月)
    df["mitori"] = k[38] + k[44]                 # 看取り 実施件数 (件/月, R5.9)
    df["mitori_hospitals"] = k[37]
    df["mitori_clinics"] = k[43]
    # --- 届出 (R6 sheet: 2024-03-31)
    df["zaishi_clinics"] = k6[9]
    df["zaishi_clinics_kyoka"] = k6[10] + k6[11]
    df["zaishi_hospitals"] = k6[5]
    df["zaishi_hospitals_kyoka"] = k6[6] + k6[7]
    df["zaishi_clinics_2023"] = k[9]             # 2023-03-31, for trend
    df["zaishi_hospitals_2023"] = k[5]
    df["houkan_st"] = k6[58]                     # 訪問看護ステーション (2024-10-01)
    df["basis"] = "医療機関所在地"
    raw = pd.read_excel(xlsx, sheet_name=s5, header=None)
    col = {int(raw.iat[0, j]): raw.iat[6, j] for j in range(7, raw.shape[1])
           if isinstance(raw.iat[0, j], (int, float)) and not pd.isna(raw.iat[0, j])}
    out = df.reset_index()
    out.attrs["national_visit"] = (col[14], col[20])
    return out


# ---------------------------------------------------------------- 2. 社会医療診療行為別統計
def estat_list(url: str, dest: Path) -> str:
    fetch(url, dest, force=True)
    return text(dest)


def ika_list_urls(top_html: str) -> list:
    """Return, newest first, (year_label, url) of the 医科診療 datalist for each edition."""
    s = re.sub(r"<script.*?</script>", "", top_html, flags=re.S)
    out = []
    for m in re.finditer(r"■([^<]{0,40}?年)社会医療診療行為別統計", s):
        seg = s[m.end(): m.end() + 20000]
        seg = seg[: seg.find("■")] if "■" in seg else seg
        mm = re.search(r"医科診療.{0,4000}?href=\"([^\"]*tclass3=[^\"]*)\"", seg, re.S) or \
            re.search(r"href=\"([^\"]*tclass3=[^\"]*)\"[^>]*>(?:(?!</a>).)*医科", seg, re.S)
        if mm:
            out.append((m.group(1), ESTAT + html.unescape(mm.group(1))))
    return out


def find_t22(list_url: str, tag: str) -> str:
    for p in range(1, 5):
        u = re.sub(r"page=\d+", f"page={p}", list_url)
        h = estat_list(u, RAW / f"estat_{tag}_p{p}.html")
        s = re.sub(r"<[^>]+>", " ", re.sub(r'<a[^>]*statInfId=(\d+)[^>]*fileKind=1[^>]*>', r" [ID \1] ", h))
        s = re.sub(r"\s+", " ", html.unescape(s))
        m = re.search(r"第２２表 医科診療（総数－１総数）.*?\[ID (\d+)\]", s)
        if m:
            return m.group(1)
    raise SystemExit(f"第22表 not found for {tag}")


def parse_t22(csv: Path) -> pd.DataFrame:
    raw = pd.read_csv(csv, encoding="cp932", header=None, dtype=str, skiprows=0)
    title = raw.iat[0, 0]
    hdr = raw[raw[5].astype(str).str.strip() == "総数"].index[0]
    d = raw.iloc[hdr + 1:].copy()
    d = d[d[0].astype(str).str.startswith("在宅医療")]

    def num(c):
        return pd.to_numeric(d[c].astype(str).str.strip().replace({"-": "0"}), errors="coerce")
    # col 5/6/7 = 総数 実施件数/回数/点数 ; 8/9 = 病院総数 ; 26/27 = 診療所総数
    out = pd.DataFrame({
        "item": d[2].astype(str).str.strip(), "saikei": d[4].astype(str).str.strip() == "*",
        "cases": num(5), "times": num(6), "hosp_cases": num(8), "clinic_cases": num(26),
    })
    out.attrs["title"] = title
    return out


def summarize(t: pd.DataFrame) -> dict:
    main = t[~t.saikei]
    r = {}
    def g(pat):
        x = main[main["item"].str.contains(pat, regex=True)]
        return x[["cases", "times", "hosp_cases", "clinic_cases"]].sum()
    r["I1_other"] = g(r"^在宅患者訪問診療料（Ⅰ）１　同一建物居住者以外")
    r["I1_same"] = g(r"^在宅患者訪問診療料（Ⅰ）１　同一建物居住者(?!以外)")
    r["I2"] = g(r"^在宅患者訪問診療料（Ⅰ）２")
    r["II"] = g(r"^在宅患者訪問診療料（Ⅱ）(?!.*加算)")
    r["zaiikan"] = g(r"^在宅時医学総合管理料(?!.*加算)")
    r["shisetsukan"] = g(r"^施設入居時等医学総合管理料(?!.*加算)")
    return r


def ndb_notes(zaitaku: pd.DataFrame) -> str:
    top = estat_list(ESTAT_TOP, RAW / "estat_shinryo_top.html")
    eds = ika_list_urls(top)
    lines = []
    results = {}
    for label, url in [eds[0]] + [e for e in eds if "令和５" in e[0]][:1]:
        n = unicodedata.normalize("NFKC", label)
        m = re.search(r"\((\d{4})\)", n) or re.search(r"令和(\d+)", n)
        tag = m.group(1) if len(m.group(1)) == 4 else str(2018 + int(m.group(1)))
        sid = find_t22(url, tag)
        csv = fetch(f"{ESTAT}/stat-search/file-download?statInfId={sid}&fileKind=1",
                    RAW / f"shinryo_koui_{tag}_t22_{sid}.csv")
        t = parse_t22(csv)
        results[label] = (sid, t.attrs["title"], summarize(t))

    f = lambda v: f"{int(v):,}"
    lines.append("National totals for sanity checks — 社会医療診療行為別統計 (NDB全数集計)")
    lines.append("Table: 診療行為の状況 医科診療 第22表 医科診療（総数－１総数） 実施件数・回数・点数，"
                 "一般医療－後期医療、診療行為（細分類）、病院（種類別）－診療所（有床－無床）別")
    lines.append("  実施件数 = 当該行為が算定されたレセプト(明細書)数 ≈ 1か月の患者数（医療機関×患者単位）")
    lines.append("  回数     = 延べ算定回数 ≈ 1か月の訪問回数")
    lines.append("  Download: https://www.e-stat.go.jp/stat-search/file-download?statInfId=<ID>&fileKind=1")
    lines.append("")
    for label, (sid, title, r) in results.items():
        lines.append(f"== {label}  statInfId={sid}")
        lines.append(f"   {title.strip(' ,')}")
        rows = [("在宅患者訪問診療料(I)1 同一建物居住者以外", r["I1_other"]),
                ("在宅患者訪問診療料(I)1 同一建物居住者", r["I1_same"]),
                ("在宅患者訪問診療料(I)2 (他院依頼)", r["I2"]),
                ("在宅患者訪問診療料(II) (併設施設)", r["II"])]
        tot = sum(x for _, x in rows)
        rows.append(("訪問診療料 合計", tot))
        rows.append(("在宅時医学総合管理料 (在医総管) 計", r["zaiikan"]))
        rows.append(("施設入居時等医学総合管理料 (施設総管) 計", r["shisetsukan"]))
        lines.append(f"   {'item':<42}{'実施件数(レセ/月)':>16}{'回数(/月)':>12}{'回/件':>7}{'病院件数':>10}{'診療所件数':>11}")
        for n, x in rows:
            lines.append(f"   {n:<42}{f(x.cases):>16}{f(x.times):>12}{x.times / x.cases:>7.2f}"
                         f"{f(x.hosp_cases):>10}{f(x.clinic_cases):>11}")
        lines.append(f"   在医総管+施設総管 実施件数 = {f(r['zaiikan'].cases + r['shisetsukan'].cases)} "
                     "(≈ 定期訪問診療を受ける患者数/月; 訪問診療料(I)1 件数と比較)")
        lines.append("")
    # compare with 医療施設調査 (national total in 地域別データ集 R5 sheet)
    lines.append("== Cross-check vs 地域別データ集 (医療施設調査 R5.9月間 訪問診療 実施件数)")
    lines.append(f"   Kanto sum visit_cases = {f(zaitaku['visit_cases'].sum())} 件/月 (R5.9)")
    h, c = zaitaku.attrs["national_visit"]
    lines.append(f"   National (全国計 in xlsx): 病院 {f(h)} + 診療所 {f(c)} = {f(h + c)} 件/月 "
                 "(vs NDB 令和5年6月審査分 回数 above)")
    lines.append("   -> this is close to NDB 回数 (visits), not 実施件数 (patients); treat 医療施設調査の実施件数 as "
                 "延べ訪問回数/月 and divide by ~回/件 above to get monthly patients.")
    return "\n".join(lines) + "\n"


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    WORK.mkdir(parents=True, exist_ok=True)
    xlsx = fetch_chiikibetsu()
    df = build_muni(xlsx)
    df.to_csv(WORK / "zaitaku_muni.csv", index=False)
    print("zaitaku_muni.csv rows:", len(df))
    notes = ndb_notes(df)
    (WORK / "ndb_national.txt").write_text(notes, encoding="utf-8")
    print(notes)


if __name__ == "__main__":
    main()
