"""関東の市区町村ごとに「中心から半径16kmの診療圏」の需要・供給・競合をまとめ、data/area_kanto.json を作る。

入力（scripts/fetch_*.py と scripts/calc_util.py の出力）:
  data/work/centroids.csv       市区町村の人口重心
  data/work/mesh1km_age.csv     1kmメッシュの年齢別人口
  data/work/mesh_muni.csv       メッシュと市区町村の対応
  data/work/util_muni.csv       保険者別の在宅認定者数・医師の居宅療養管理指導受給者数
  data/work/zaitaku_muni.csv    市区町村別の訪問診療件数（医療機関所在地）
  data/work/projection.csv      将来推計人口
  data/work/competitors.csv     競合（届出名簿＋緯度経度）
  data/util_bench.csv           要介護度別の基準利用率

方法は docs/area-model.md を参照。
"""
import json, math
import numpy as np, pandas as pd

W = 'data/work/'
R_KM = 16.0
L = ['shien1', 'shien2', 'kaigo1', 'kaigo2', 'kaigo3', 'kaigo4', 'kaigo5']
KYORYO_COVERAGE = 0.855
VISITS_PER_PATIENT = 1.89
SEIREI = {11100: range(11101, 11111), 12100: range(12101, 12107), 14100: range(14101, 14119),
          14130: range(14131, 14138), 14150: range(14151, 14154)}
OSATO = {'11202', '11218', '11408'}  # 大里広域市町村圏組合（熊谷市・深谷市・寄居町）
ISLANDS = {str(c) for c in range(13360, 13422)}  # 伊豆諸島・小笠原（メッシュ人口なし）
TYPE_WEIGHT = {'kyoka_tandoku': 8, 'kyoka_renkei': 3, 'zaishi': 1, 'general': 0.25}


def insurer(code):
    if code in OSATO:
        return 'OSATO'
    c = int(code)
    for city, wards in SEIREI.items():
        if c in wards:
            return f'{city:05d}'
    return code


def loc_unit(code):  # 訪問診療件数（医療施設調査）の単位。政令市は市全体
    return insurer(code) if code not in OSATO else code


def mesh_center(m):
    m = str(m)
    lat = int(m[0:2]) / 1.5 + int(m[4]) * 5 / 60 + int(m[6]) * 30 / 3600 + 15 / 3600
    lon = int(m[2:4]) + 100 + int(m[5]) * 7.5 / 60 + int(m[7]) * 45 / 3600 + 22.5 / 3600
    return lat, lon


def dist_km(lat1, lon1, lat2, lon2):
    p = np.pi / 180
    a = np.sin((lat2 - lat1) * p / 2) ** 2 + np.cos(lat1 * p) * np.cos(lat2 * p) * np.sin((lon2 - lon1) * p / 2) ** 2
    return 12742 * np.arcsin(np.sqrt(a))


# ---------- 保険者ごとの在宅認定者数と現在の患者数 ----------
u = pd.read_csv(W + 'util_muni.csv', dtype={'code': str, 'pref_code': str})
u.loc[u.is_kouiki == True, 'code'] = 'OSATO'
ins = u.set_index('code')
zt_cols = [f'zaitaku_{l}' for l in L]
ph_cols = [f'phys_{l}' for l in L]
ins['S'] = ins[ph_cols].sum(axis=1) / KYORYO_COVERAGE

# ---------- メッシュ → 保険者（複数にまたがるメッシュは等分） ----------
mesh = pd.read_csv(W + 'mesh1km_age.csv', dtype={'mesh': str})
mm = pd.read_csv(W + 'mesh_muni.csv', dtype={'code': str, 'mesh': str})
mm = mm[~mm.code.str.endswith('99')]
mm['ins'] = mm.code.map(insurer)
mm['share'] = 1 / mm.groupby('mesh').code.transform('count')
mm = mm.merge(mesh[['mesh', 'pop65', 'pop75', 'pop85']], on='mesh', how='inner')
for c in ['pop65', 'pop75', 'pop85']:
    mm[c + '_s'] = mm[c] * mm.share
ins_p75 = mm.groupby('ins').pop75_s.sum()

# 保険者ごとの「75歳以上1人あたり」の在宅認定者数・患者数。関東外（福島・新潟・山梨・長野・静岡）は関東の中央値を使う
dens = pd.DataFrame(index=ins.index)
for c in zt_cols + ['S']:
    dens[c] = ins[c] / ins_p75.reindex(ins.index)
dens = dens.replace([np.inf, -np.inf], np.nan)
fallback = dens.median()
mm = mm.join(dens, on='ins')
mm['kanto'] = mm.ins.isin(dens.index)
for c in zt_cols + ['S']:
    mm[c] = mm[c].fillna(fallback[c]) * mm.pop75_s

# メッシュ単位に集約
agg = mm.groupby('mesh')[zt_cols + ['S', 'pop65_s', 'pop75_s', 'pop85_s']].sum()
agg_ins = mm.groupby(['mesh', 'ins']).pop75_s.sum().reset_index()
latlon = np.array([mesh_center(m) for m in agg.index])
agg['lat'], agg['lon'] = latlon[:, 0], latlon[:, 1]

# ---------- 将来推計（75歳以上と85歳以上の伸びの平均） ----------
pj = pd.read_csv(W + 'projection.csv', dtype={'code': str})
pj = pj[(pj.kind != 0) | pj.code.str.startswith('131')]  # 政令市は区ではなく市全体の行を使う（東京23区はそのまま）
pj['ins'] = pj.code.map(insurer)
pg = pj.groupby(['ins', 'year'])[['pop75', 'pop85']].sum().reset_index()
base = pg[pg.year == 2020].set_index('ins')
YEARS = [2025, 2030, 2035, 2040, 2045, 2050]
grow = {}
for y in YEARS:
    cur = pg[pg.year == y].set_index('ins')
    grow[y] = 0.5 * (cur.pop75 / base.pop75 + cur.pop85 / base.pop85)
grow = pd.DataFrame(grow)

# ---------- 競合 ----------
import os
comp = pd.read_csv(W + 'competitors.csv', dtype={'inst_code': str}) if os.path.exists(W + 'competitors.csv') else pd.DataFrame(columns=['name', 'type', 'lat', 'lon', 'is_hospital'])
comp = comp.dropna(subset=['lat', 'lon']).copy()
cm_lat = np.array([mesh_center(m)[0] for m in agg.index])
cm_lon = np.array([mesh_center(m)[1] for m in agg.index])
mesh_codes = mm.groupby('mesh').code.first()


def nearest_mesh(lat, lon):
    d = (cm_lat - lat) ** 2 + ((cm_lon - lon) * math.cos(lat * math.pi / 180)) ** 2
    return agg.index[int(np.argmin(d))]


comp['mesh'] = [nearest_mesh(a, b) for a, b in zip(comp.lat, comp.lon)]
comp['muni'] = comp.mesh.map(mesh_codes)
comp['unit'] = comp.muni.map(loc_unit)
zm = pd.read_csv(W + 'zaitaku_muni.csv', dtype={'code': str}).set_index('code')
scale = ins.S.sum() / zm.visit_patients_est.sum()  # 医療施設調査(2023) → 居宅療養管理指導ベース(2024)の水準へ
comp['hosp'] = comp.is_hospital.eq('病院')
comp['w'] = comp.type.map(TYPE_WEIGHT)
pool_clinic = zm.visit_cases_clinic / VISITS_PER_PATIENT * scale
pool_hosp = zm.visit_cases_hosp / VISITS_PER_PATIENT * scale
wsum = comp.groupby(['unit', 'hosp']).w.transform('sum')
pool = np.where(comp.hosp, comp.unit.map(pool_hosp), comp.unit.map(pool_clinic))
comp['n'] = np.nan_to_num(pool) * comp.w / wsum

# ---------- 診療圏（中心から16km） ----------
ce = pd.read_csv(W + 'centroids.csv', dtype={'code': str})
ce = ce[~ce.code.isin(ISLANDS) & (ce.code != '13100') & ~ce.code.astype(int).isin(SEIREI.keys())]
bench = pd.read_csv('data/util_bench.csv').set_index('level').u_bench

centers = []
for r in ce.itertuples():
    d = dist_km(r.lat, r.lon, agg.lat.values, agg.lon.values)
    inside = d <= R_KM
    a = agg[inside]
    g = {}
    wi = agg_ins[agg_ins.mesh.isin(a.index)].groupby('ins').pop75_s.sum()
    wd = wi * 0  # 需要の重み
    for j in wi.index:
        if j in dens.index:
            wd[j] = wi[j] * dens.loc[j, zt_cols].values @ bench.values
    for y in YEARS:
        gy = grow[y].reindex(wd.index).fillna(1.0)
        g[y] = round(float((wd * gy).sum() / wd.sum()) if wd.sum() else 1.0, 4)
    own = insurer(r.code)
    centers.append(dict(
        code=r.code, pref=r.pref, name=r.name, lat=round(r.lat, 5), lon=round(r.lon, 5),
        pop65=round(a.pop65_s.sum()), pop75=round(a.pop75_s.sum()), pop85=round(a.pop85_s.sum()),
        zt=[round(a[c].sum()) for c in zt_cols],
        S=round(a.S.sum()),
        growth=g,
        ins=own, insName=str(ins.loc[own, 'name']) if own in ins.index else '',
    ))

comps_out = [[str(x.name), x.type, round(x.lat, 5), round(x.lon, 5), round(float(x.n), 1), int(x.hosp)] for x in comp.itertuples()]
b1 = pd.read_csv(W + 'phys_ratio_national.csv').set_index('level').share_b1
out = dict(
    meta=dict(
        radius_km=R_KM, kyoryo_coverage=KYORYO_COVERAGE,
        sources=['国勢調査2020（1kmメッシュ・人口重心）', '介護保険事業状況報告 年報 令和6年度', '介護給付費等実態統計 令和6年度',
                 '在宅医療にかかる地域別データ集（医療施設調査2023）', '関東信越厚生局 届出受理医療機関名簿', '社人研 地域別将来推計人口（令和5年推計）'],
    ),
    levels=L, bench=[round(float(bench[l]), 4) for l in L], homeShare=[round(float(b1[l]), 3) for l in L],
    centers=centers, competitors=comps_out,
)
with open('data/area_kanto.json', 'w') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
print(f'centers={len(centers)} competitors={len(comps_out)} scale={scale:.3f}')


# ================= メッシュの需要・供給と、集患しやすさスコア =================
# メッシュごとの D_m（基準利用率100%のとき）と S_m をページに同梱し、選んだ市区町村の計算はブラウザで行う。
# 関東全体での順位（スコア）は、標準的なクリニックの定常患者数 N*_std をここで計算して付ける。
MU, LAM, THETA, P_SW, ALPHA, N0 = 0.035, 0.03, 1.5, 0.10, 0.5, 20
Q = {'kyoka_tandoku': 1.3, 'kyoka_renkei': 1.15, 'zaishi': 1.0, 'general': 0.6}
pop_tot = mesh.set_index('mesh').pop_total
agg['pop'] = pop_tot.reindex(agg.index).fillna(0)
agg['D'] = agg[zt_cols].values @ bench.values / KYORYO_COVERAGE
agg['d0'] = np.where(agg['pop'] >= 4000, 4.0, np.where(agg['pop'] >= 1000, 6.0, 9.0))
fac_share = 1 - float((b1 * bench).sum() / bench.sum())  # 施設（単一建物2人以上）の割合

# 競合による「紹介の取り合い」の強さ C_m = Σ_i A_i f_m(d_im)
C = np.zeros(len(agg))
if len(comp):
    A_i = comp.type.map(Q).values * np.power(comp.n.values, ALPHA)
    clat, clon = comp.lat.values, comp.lon.values
    for s0 in range(0, len(agg), 2000):
        sl = slice(s0, s0 + 2000)
        d = dist_km(agg.lat.values[sl, None], agg.lon.values[sl, None], clat[None, :], clon[None, :])
        f = np.exp(-d / agg.d0.values[sl, None]) * (d <= R_KM)
        C[sl] = f @ A_i
agg['C'] = C

g5 = {c['code']: c['growth'][2030] / c['growth'][2025] for c in centers}
for c in centers:
    d = dist_km(c['lat'], c['lon'], agg.lat.values, agg.lon.values)
    m = d <= R_KM
    a = agg[m]
    f0 = np.exp(-d[m] / a.d0.values)
    G = np.maximum(0, a.D.values - a.S.values)
    dD = a.D.values * (g5[c['code']] ** (1 / 60) - 1)
    F_ref = MU * a.S.values + np.maximum(0, dD) + a.S.values * fac_share * P_SW / 12
    N = 100.0
    for _ in range(40):
        A0 = (N + N0) ** ALPHA
        pi = A0 * f0 / (A0 * f0 + a.C.values + 1e-9)
        I = (pi * F_ref + np.minimum(1, THETA * pi) * LAM * G).sum()
        N = 0.5 * N + 0.5 * I / MU
    Dt, St = a.D.sum(), a.S.sum()
    c.update(D=round(Dt), S=round(St), Nstd=round(N, 1), Istd=round(I, 2),
             comp_idx=round(float((a.C * a.D).sum() / Dt), 2) if Dt else 0)


def pct(vals):
    v = np.array(vals, dtype=float)
    return [round(float((v < x).mean() * 100)) for x in v]


for key, vals in [('score', [c['Nstd'] for c in centers]),
                  ('p_gap', [1 - c['S'] / c['D'] if c['D'] else 0 for c in centers]),
                  ('p_size', [c['D'] for c in centers]),
                  ('p_comp', [-c['comp_idx'] for c in centers]),
                  ('p_growth', [c['growth'][2030] / c['growth'][2025] for c in centers])]:
    for c, p in zip(centers, pct(vals)):
        c[key] = p

meshes_out = [[k, round(float(r['pop'])), round(float(r['D']), 2), round(float(r['S']), 2)]
              for k, r in agg[['pop', 'D', 'S']].iterrows() if r['D'] > 0 or r['S'] > 0]
out.update(centers=centers, meshes=meshes_out, facShare=round(fac_share, 3),
           params=dict(mu=MU, lam=LAM, theta=THETA, p_sw=P_SW, alpha=ALPHA, n0=N0, q=Q))
with open('data/area_kanto.json', 'w') as f:
    json.dump(out, f, ensure_ascii=False, separators=(',', ':'))
print(f'meshes={len(meshes_out)} facShare={fac_share:.3f}')
top = sorted(centers, key=lambda c: -c['score'])
for c in top[:5] + top[-3:]:
    print(c['name'], c['score'], c['Nstd'], c['Istd'], c['D'], c['S'], c['comp_idx'])
