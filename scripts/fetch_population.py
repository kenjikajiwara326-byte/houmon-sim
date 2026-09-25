#!/usr/bin/env python3
"""Download and clean Japanese public population statistics for the Kanto
home-visit clinic (訪問診療) market model.

Outputs (UTF-8 CSV, under data/work/):
  centroids.csv     令和2年国勢調査 市区町村別人口重心 (統計局 統計トピックスNo.135)
  mesh1km_age.csv   令和2年国勢調査 3次メッシュ(1km) 人口 総数/65+/75+/85+ (e-Stat 統計GIS T001140)
  mesh_muni.csv     市区町村別メッシュ・コード一覧 (統計局, 令和2年10月1日境域)
  muni_pop.csv      令和2年国勢調査 市区町村別 年齢別人口 (第2-5表) + 面積 (第1-1表)

Raw downloads are cached under data/raw/pop/ (re-run uses the cache; pass
--refresh to force re-download).

Usage:  python3 scripts/fetch_population.py [--refresh]
"""
import io
import sys
import time
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
import requests

ROOT = Path(__file__).resolve().parent.parent
RAW = ROOT / "data" / "raw" / "pop"
WORK = ROOT / "data" / "work"
REFRESH = "--refresh" in sys.argv

KANTO = ["08", "09", "10", "11", "12", "13", "14"]
NEIGHBORS = ["07", "15", "19", "20", "22"]  # 福島 新潟 山梨 長野 静岡
FIRST_MESH = [5238, 5239, 5240, 5241, 5338, 5339, 5340, 5341,
              5438, 5439, 5440, 5441, 5538, 5539, 5540, 5541]
BUFFER_KM = 16.0

# --- sources -----------------------------------------------------------------
CENTROID_URL = "https://www.stat.go.jp/data/kokusei/topics/zuhyou/topics135_10.xlsx"
MESH_URL = ("https://www.e-stat.go.jp/gis/statmap-search/data"
            "?statsId=T001140&code={code}&downloadType=2")
MESHMUNI_URL = "https://www.stat.go.jp/data/mesh/csv/{pref}.csv"
ESTAT_FILE = "https://www.e-stat.go.jp/stat-search/file-download?statInfId={id}&fileKind=0"
# 令和2年国勢調査 人口等基本集計 第2-5表 (男女，年齢（各歳），… 市区町村), one file per prefecture:
# statInfId 000032148536 (01 北海道) … 000032148582 (47 沖縄県), sequential.
AGE_STATINF_BASE = 32148536
# 第1-1表 (男女別人口 … 面積（参考）及び人口密度 － 全国，都道府県，市区町村), all Japan
AREA_STATINF = "000032142402"

session = requests.Session()
session.headers["User-Agent"] = "houmon-sim data fetch (research use)"


def fetch(url: str, dest: Path) -> Path:
    """Download url to dest unless cached. Sleeps 0.5s after each request."""
    if dest.exists() and dest.stat().st_size > 0 and not REFRESH:
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    r = session.get(url, timeout=120)
    time.sleep(0.5)
    if r.status_code == 404:
        raise FileNotFoundError(url)
    r.raise_for_status()
    if "text/html" in r.headers.get("content-type", ""):
        raise FileNotFoundError(f"HTML instead of data: {url}")
    dest.write_bytes(r.content)
    print(f"  downloaded {dest.relative_to(ROOT)} ({len(r.content):,} B)")
    return dest


# --- 1. centroids --------------------------------------------------------------
def build_centroids() -> pd.DataFrame:
    print("[1] 人口重心")
    f = fetch(CENTROID_URL, RAW / "topics135_10.xlsx")
    sheets = pd.read_excel(f, sheet_name=None, header=None, dtype=str)
    rows = []
    for sheet, df in sheets.items():
        if sheet == "全国":
            continue
        df = df.iloc[2:, :4].dropna(subset=[0])
        df.columns = ["code", "name", "lon", "lat"]
        df = df[df["code"].str.fullmatch(r"\d{5}")]
        df = df[df["code"].str[:2].isin(KANTO)]
        df["pref"] = sheet
        rows.append(df)
    out = pd.concat(rows)
    out = out[~out["code"].str.endswith("000")]  # drop prefecture rows
    out["lat"] = out["lat"].astype(float).round(6)
    out["lon"] = out["lon"].astype(float).round(6)
    out = out[["code", "pref", "name", "lat", "lon"]].sort_values("code")
    out.to_csv(WORK / "centroids.csv", index=False)
    print(f"  centroids.csv: {len(out)} rows")
    return out


# --- 2. 1km mesh -----------------------------------------------------------------
def build_mesh() -> pd.DataFrame:
    print("[2] 3次メッシュ 年齢別人口 (T001140)")
    frames = []
    for code in FIRST_MESH:
        dest = RAW / "mesh" / f"tblT001140S{code}.zip"
        try:
            fetch(MESH_URL.format(code=code), dest)
        except FileNotFoundError:
            print(f"  1次メッシュ {code}: no data (sea)")
            continue
        with zipfile.ZipFile(dest) as z:
            name = [n for n in z.namelist() if n.lower().endswith(".txt")][0]
            text = z.read(name).decode("cp932")
        df = pd.read_csv(io.StringIO(text), dtype=str, skiprows=[1])  # row 2 = Japanese labels
        frames.append(df)
    raw = pd.concat(frames, ignore_index=True)
    col = {"T001140001": "pop_total", "T001140019": "pop65",
           "T001140022": "pop75", "T001140025": "pop85"}
    df = raw[["KEY_CODE", "HTKSYORI", "HTKSAKI", "GASSAN", *col]].rename(
        columns={"KEY_CODE": "mesh", **col})
    for c in col.values():
        df[c] = pd.to_numeric(df[c].replace({"*": np.nan, "X": np.nan, "-": 0}), errors="coerce")
    df["htk"] = df["HTKSYORI"].astype(int)
    # 秘匿処理: HTKSYORI=2 -> this cell's age detail is merged ("*") into cell HTKSAKI;
    # HTKSYORI=1 -> this cell's age detail includes the cells listed in GASSAN.
    # pop_total is always the cell's own value. We re-distribute the merged age
    # counts across the group in proportion to pop_total (float, not rounded).
    df["group"] = np.where(df["htk"] == 2, df["HTKSAKI"], df["mesh"])
    grp_pop = df.groupby("group")["pop_total"].transform("sum")
    host = df[df["htk"] != 2].set_index("mesh")
    share = np.where(grp_pop > 0, df["pop_total"] / grp_pop.replace(0, np.nan), 0)
    for c in ["pop65", "pop75", "pop85"]:
        gval = df["group"].map(host[c]).fillna(0)
        alloc = np.where(df["htk"] == 0, df[c], gval * np.nan_to_num(share))
        df[c] = np.round(alloc.astype(float), 2)
    missing_host = (~df["group"].isin(host.index)).sum()
    if missing_host:
        print(f"  WARNING {missing_host} merged cells whose host is outside download set -> 0")
    df["suppressed"] = (df["htk"] == 2).astype(int)
    out = df[["mesh", "pop_total", "pop65", "pop75", "pop85", "suppressed"]].copy()
    out["pop_total"] = out["pop_total"].fillna(0).astype(int)
    out = out.sort_values("mesh")
    out.to_csv(WORK / "mesh1km_age.csv", index=False)
    print(f"  mesh1km_age.csv: {len(out)} rows, suppressed(age merged) cells: "
          f"{int(out['suppressed'].sum())}, host cells: {(df['htk'] == 1).sum()}")
    return out


# --- 3. mesh <-> municipality ------------------------------------------------------
def mesh_rc(m: pd.Series):
    s = m.astype(str)
    row = s.str[0:2].astype(int) * 80 + s.str[4].astype(int) * 10 + s.str[6].astype(int)
    col = s.str[2:4].astype(int) * 80 + s.str[5].astype(int) * 10 + s.str[7].astype(int)
    return row.to_numpy(), col.to_numpy()


def build_mesh_muni() -> pd.DataFrame:
    print("[3] 市区町村別メッシュコード一覧")
    frames = []
    for p in KANTO + NEIGHBORS:
        f = fetch(MESHMUNI_URL.format(pref=p), RAW / "mesh_muni" / f"{p}.csv")
        df = pd.read_csv(f, encoding="cp932", dtype=str)
        df.columns = ["code", "name", "mesh"]
        frames.append(df)
    df = pd.concat(frames, ignore_index=True)
    df["code"] = df["code"].str.zfill(5)
    kanto = df[df["code"].str[:2].isin(KANTO)]
    nb = df[~df["code"].str[:2].isin(KANTO)]
    # Keep neighbor-prefecture rows only for meshes within BUFFER_KM of a Kanto
    # mesh (mainland 1次メッシュ window only). Grid dilation with elliptical kernel.
    r0, c0 = 52 * 80, 38 * 80
    grid = np.zeros((4 * 80, 4 * 80), dtype=bool)
    kr, kc = mesh_rc(kanto["mesh"])
    ok = (kr >= r0) & (kr < r0 + 320) & (kc >= c0) & (kc < c0 + 320)
    grid[kr[ok] - r0, kc[ok] - c0] = True
    dy, dx = 0.925, 1.125  # km per 30" lat / 45" lon at ~36N
    ny, nx = int(BUFFER_KM / dy) + 1, int(BUFFER_KM / dx) + 1
    buf = np.zeros_like(grid)
    ys, xs = np.nonzero(grid)
    for oy in range(-ny, ny + 1):
        for ox in range(-nx, nx + 1):
            if (oy * dy) ** 2 + (ox * dx) ** 2 > BUFFER_KM ** 2:
                continue
            y, x = ys + oy, xs + ox
            m = (y >= 0) & (y < 320) & (x >= 0) & (x < 320)
            buf[y[m], x[m]] = True
    nr, nc = mesh_rc(nb["mesh"])
    inwin = (nr >= r0) & (nr < r0 + 320) & (nc >= c0) & (nc < c0 + 320)
    keep = np.zeros(len(nb), dtype=bool)
    keep[inwin] = buf[nr[inwin] - r0, nc[inwin] - c0]
    nb = nb[keep]
    out = pd.concat([kanto, nb])[["code", "name", "mesh"]].drop_duplicates()
    out = out.sort_values(["code", "mesh"])
    out.to_csv(WORK / "mesh_muni.csv", index=False)
    print(f"  mesh_muni.csv: {len(out)} rows (Kanto {len(kanto)}, neighbors within "
          f"{BUFFER_KM:g}km {len(nb)}; {out['mesh'].nunique()} unique meshes)")
    return out


# --- 4. municipality population by age ---------------------------------------------
def build_muni_pop() -> pd.DataFrame:
    print("[4] 市区町村別 年齢別人口 (第2-5表) + 面積 (第1-1表)")
    rows = []
    for p in KANTO:
        sid = f"{AGE_STATINF_BASE + int(p) - 1:012d}"
        f = fetch(ESTAT_FILE.format(id=sid), RAW / "muni_age" / f"b02_05-{p}.xlsx")
        x = pd.read_excel(f, header=None, dtype=str)
        hdr = x.iloc[9].tolist()
        idx = {h.split("_", 1)[1] if isinstance(h, str) and "_" in h else h: i
               for i, h in enumerate(hdr)}
        d = x.iloc[12:]
        d = d[d[0].str.startswith("0_") & d[1].str.startswith("0_")  # 国籍総数, 男女総数
              & (d[2] != "9")]  # exclude 2000年旧市町村 rows
        assert d[7].str[:2].eq(p).all(), f"statInfId {sid} is not pref {p}"
        rows.append(pd.DataFrame({
            "code": d[7].str.zfill(5),
            "kind": d[2],
            "name": d[8].str.split("_", n=1).str[1],
            "pop_total": d[idx["総数"]],
            "pop65": d[idx["（再掲）65歳以上"]],
            "pop75": d[idx["（再掲）75歳以上"]],
            "pop85": d[idx["（再掲）85歳以上"]],
            "age_unknown": d[idx["年齢「不詳」"]],
        }))
    age = pd.concat(rows, ignore_index=True)
    for c in ["pop_total", "pop65", "pop75", "pop85", "age_unknown"]:
        # "-" in e-Stat tables means zero
        age[c] = pd.to_numeric(age[c].str.replace(",", "").replace("-", "0"),
                               errors="coerce").astype("Int64")

    f = fetch(ESTAT_FILE.format(id=AREA_STATINF), RAW / "b01_01.xlsx")
    a = pd.read_excel(f, header=None, dtype=str).iloc[15:]
    a = a[a[0] != "9"]  # drop 2000年(平成12年)旧市町村 rows
    area = pd.DataFrame({"code": a[5], "area_km2": pd.to_numeric(a[14], errors="coerce")})
    area = area.drop_duplicates("code")
    out = age.merge(area, on="code", how="left")
    kind_map = {"a": "pref", "1": "designated_city_or_tokubetsukubu", "0": "ward",
                "2": "city", "3": "town_village"}
    out["kind"] = out["kind"].map(kind_map).fillna(out["kind"])
    out = out[["code", "name", "kind", "pop_total", "pop65", "pop75", "pop85",
               "age_unknown", "area_km2"]].sort_values("code")
    out.to_csv(WORK / "muni_pop.csv", index=False)
    print(f"  muni_pop.csv: {len(out)} rows, missing area: {out['area_km2'].isna().sum()}")
    return out


if __name__ == "__main__":
    WORK.mkdir(parents=True, exist_ok=True)
    build_centroids()
    build_mesh()
    build_mesh_muni()
    build_muni_pop()
