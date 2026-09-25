#!/usr/bin/env python3
"""Fetch IPSS 日本の地域別将来推計人口（令和5(2023)年推計） and extract Kanto municipalities.

Source page: https://www.ipss.go.jp/pp-shicyoson/j/shicyoson23/t-page.asp
Data file  : https://www.ipss.go.jp/pp-shicyoson/j/shicyoson23/3kekka/suikei_kekka.xlsx
  (全国・都道府県・市区町村 男女5歳階級別, 2020(国勢調査実績)〜2050年, 5年毎)

Output: data/work/projection.csv
  code (5-digit 市区町村コード, 政令市の区・政令市本体を含む), name, kind, year,
  pop_total, pop65, pop75, pop85
kind: 0=政令市の区（東京23区含む）, 1=政令市, 2=その他の市, 3=町村
Note: 政令市 rows and their 区 rows both appear; avoid double counting when summing.
"""
import time
import subprocess
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RAW = ROOT / "data/raw/ipss"
OUT = ROOT / "data/work/projection.csv"
URL = "https://www.ipss.go.jp/pp-shicyoson/j/shicyoson23/3kekka/suikei_kekka.xlsx"
PREFS = {"08", "09", "10", "11", "12", "13", "14"}


def fetch(url: str, dest: Path, force: bool = False) -> Path:
    if dest.exists() and dest.stat().st_size > 0 and not force:
        return dest
    dest.parent.mkdir(parents=True, exist_ok=True)
    subprocess.run(["curl", "-sSfL", "-o", str(dest), url], check=True)
    time.sleep(0.5)
    return dest


def main() -> None:
    xlsx = fetch(URL, RAW / "suikei_kekka.xlsx")
    # header rows: row 4 (category) and row 5 (age bands); data from row 6
    df = pd.read_excel(xlsx, header=None, skiprows=5)
    # columns 0-4: code, kind, pref, muni, year; 5: total 計; 6..25: 20 age bands (総数)
    bands = ["0-4", "5-9", "10-14", "15-19", "20-24", "25-29", "30-34", "35-39",
             "40-44", "45-49", "50-54", "55-59", "60-64", "65-69", "70-74",
             "75-79", "80-84", "85-89", "90-94", "95+"]
    df = df.iloc[:, :26]
    df.columns = ["code", "kind", "pref", "muni", "year", "total"] + bands
    df = df.dropna(subset=["code", "year"])
    df["code"] = df["code"].astype(int).astype(str).str.zfill(5)
    df = df[df["code"].str[:2].isin(PREFS) & (df["kind"].astype(str) != "a")]
    df["year"] = df["year"].astype(str).str.extract(r"(\d{4})")[0].astype(int)
    for c in ["total"] + bands:
        df[c] = pd.to_numeric(df[c])
    out = pd.DataFrame({
        "code": df["code"],
        "name": df["muni"],
        "kind": df["kind"].astype(int),
        "year": df["year"],
        "pop_total": df["total"],
        "pop65": df[bands[13:]].sum(axis=1),
        "pop75": df[bands[15:]].sum(axis=1),
        "pop85": df[bands[17:]].sum(axis=1),
    })
    # sanity: bands sum to total
    diff = (df[bands].sum(axis=1) - df["total"]).abs().max()
    assert diff < 1, f"age bands do not sum to total (max diff {diff})"
    out = out.sort_values(["code", "year"]).reset_index(drop=True)
    for c in ["pop_total", "pop65", "pop75", "pop85"]:
        out[c] = out[c].round().astype(int)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    out.to_csv(OUT, index=False)
    print(f"wrote {OUT} rows={len(out)} municipalities={out['code'].nunique()}")


if __name__ == "__main__":
    main()
