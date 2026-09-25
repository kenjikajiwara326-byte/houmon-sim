"""要介護度別の訪問診療利用率を市区町村（保険者）単位で計算し、関東の上位水準（在宅認定者で重みづけした75パーセンタイル）を潜在需要の基準にする。

入力: data/work/kaigo_hokensha.csv, data/work/phys_ratio_national.csv（scripts/fetch_kaigo.py の出力）
出力: data/util_bench.csv, data/work/util_muni.csv
"""
import numpy as np, pandas as pd

L = ['shien1', 'shien2', 'kaigo1', 'kaigo2', 'kaigo3', 'kaigo4', 'kaigo5']
KYORYO_COVERAGE = 0.855  # 訪問診療患者のうち医師の居宅療養管理指導が算定される割合（全国: 居宅療養管理指導 医師 約89万 / 在医総管+施設総管 約104万）

k = pd.read_csv('data/work/kaigo_hokensha.csv', dtype={'code': str, 'pref_code': str})
r = pd.read_csv('data/work/phys_ratio_national.csv').set_index('level')['r_phys_per_recipient']
for l in L:
    fac = k[f'shisetsu_{l}'] if f'shisetsu_{l}' in k else 0
    k[f'zaitaku_{l}'] = k[f'nintei_{l}'] - fac
    k[f'phys_{l}'] = k[f'kyoryo_all_{l}'] * r[l]
    k[f'u_{l}'] = k[f'phys_{l}'] / k[f'zaitaku_{l}']


def wq(x, w, q):
    i = np.argsort(x.values)
    c = np.cumsum(w.values[i]) / w.values.sum()
    return x.values[i][np.searchsorted(c, q)]


rows = []
for l in L:
    u, w = k[f'u_{l}'], k[f'zaitaku_{l}']
    rows.append(dict(level=l, u_median=u.median(), u_wmedian=wq(u, w, .5), u_bench=wq(u, w, .75), u_w90=wq(u, w, .9)))
b = pd.DataFrame(rows)
b.to_csv('data/util_bench.csv', index=False, float_format='%.4f')
k.to_csv('data/work/util_muni.csv', index=False)
print((b.set_index('level') * 100).round(1).to_string())
