"""
Tedarik Zincirinde Dagitim Aglari Tasarimi Uzerine Bir Uygulama
================================================================
Beykoz Akademi Dergisi, 2015; 3(1), 67-84
Secmen, Oncan, Tuna

Ag Yapisi:
    Tedarikciler : T1=Rusya (i=1), T2=Almanya (i=2), T3=Norvec (i=3)
    Fabrikalar   : F1=Ispanya (j=1), F2=Rusya (j=2)
    Dag. Merkezi : Ukrayna (tek DC)
    Musteriler   : M1=Finlandiya (k=1), M2=Turkiye (k=2), M3=Rusya (k=3)
    Modlar       : t=1 Tren, t=2 Kara, t=3 Deniz
    Donemler     : p=1..6

Karar Degiskenleri:
    X_ijtp : Ted.i -> Fab.j, mod t, donem p
    W_jp   : Fab.j -> DC, donem p
    Y_kp   : DC -> Musteri k, donem p
    B_kp   : Karsilanamayan talep, musteri k, donem p  (p=0..6)
    Q_p    : DC donem sonu stok miktari  (p=0..6)

Parametreler makaledeki degerlere gore:
    Fj = 1  (fabrika->DC teslimat suresi, donem)
    Gk = 2  (DC->musteri teslimat suresi, donem)
    pi = 0.10  (firsat maliyeti katsayisi)
    Q0 = 40.000  (baslangic DC stoku)

Giris dosyalari (script ile ayni klasorde olmali):
    parameters.xlsx           -> kapasite, talep, maliyet, skaler
    transportation_costs.xlsx -> C_ijtp tasima maliyet matrisi

Referans: Min Z = 922,396.5  (Secmen et al. 2015, CPLEX)
"""

import gurobipy as gp
from gurobipy import GRB
import pandas as pd
import os
import sys

# =========================================================================
#  DOSYA YOLLARI
# =========================================================================
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
FILE_PARAM = os.path.join(BASE_DIR, "parameters.xlsx")
FILE_TRANS = os.path.join(BASE_DIR, "transportation_costs.xlsx")
RESULT_DIR = os.path.join(BASE_DIR, "results")
os.makedirs(RESULT_DIR, exist_ok=True)

for fp in [FILE_PARAM, FILE_TRANS]:
    if not os.path.exists(fp):
        print(f"HATA: Dosya bulunamadi -> {fp}")
        sys.exit(1)

print("=" * 65)
print("  PARAMETRELER EXCEL'DEN OKUNUYOR")
print("=" * 65)

xl_p = pd.ExcelFile(FILE_PARAM)
xl_c = pd.ExcelFile(FILE_TRANS)


def read_block(df, row_slice, col_slice):
    return df.iloc[row_slice, col_slice].astype(float).values


# =========================================================================
#  1. SKALER PARAMETRELER  (SCALAR_PARAMS)
#     Satir 2: pi  | Satir 3: Q0  | Satir 6: Z* referans
# =========================================================================
df_sc   = xl_p.parse("SCALAR_PARAMS", header=None)
pi_rate = float(df_sc.iloc[2, 2])   # 0.10
Q_init  = float(df_sc.iloc[3, 2])   # 40000
Fj      = int(df_sc.iloc[4, 2])     # 1  (fabrika->DC gecikme donemi)
Gk      = int(df_sc.iloc[5, 2])     # 2  (DC->musteri gecikme donemi)
Z_ref   = float(df_sc.iloc[6, 2])   # 922396.5

print(f"\n  pi={pi_rate}  Q0={Q_init:,.0f}  Fj={Fj}  Gk={Gk}  Z*={Z_ref:,.1f}")

# =========================================================================
#  2. TEDARIKCI KAPASITESi  alpha_ip  (SUPPLIER_CAP)
#     Satirlar 2-4, Sutunlar 2-7
# =========================================================================
df_supp   = xl_p.parse("SUPPLIER_CAP", header=None)
alpha_mat = read_block(df_supp, slice(2, 5), slice(2, 8))
alpha     = {(i, p): alpha_mat[i - 1, p - 1]
             for i in [1, 2, 3] for p in range(1, 7)}

# =========================================================================
#  3. FABRiKA KAPASiTESi  b_jp  (FACTORY_CAP)
#     Satirlar 2-3, Sutunlar 2-7
# =========================================================================
df_fact = xl_p.parse("FACTORY_CAP", header=None)
b_mat   = read_block(df_fact, slice(2, 4), slice(2, 8))
b_cap   = {(j, p): b_mat[j - 1, p - 1]
           for j in [1, 2] for p in range(1, 7)}

# =========================================================================
#  4. DC KAPASiTESi  c_p  (DC_CAP)
#     Satir 2, Sutunlar 2-7
# =========================================================================
df_dc = xl_p.parse("DC_CAP", header=None)
c_cap = {p: float(df_dc.iloc[2, p + 1]) for p in range(1, 7)}

# =========================================================================
#  5. MUSTERi TALEBi  d_kp  (DEMAND)
#     Satirlar 2-4, Sutunlar 2-7
# =========================================================================
df_dem = xl_p.parse("DEMAND", header=None)
d_mat  = read_block(df_dem, slice(2, 5), slice(2, 8))
d      = {(k, p): d_mat[k - 1, p - 1]
          for k in [1, 2, 3] for p in range(1, 7)}

# =========================================================================
#  6. MOD KAPASiTESi  A_tp  (MODE_CAP)
#     Satirlar 2-4, Sutunlar 2-7
# =========================================================================
df_mode = xl_p.parse("MODE_CAP", header=None)
A_mat   = read_block(df_mode, slice(2, 5), slice(2, 8))
A_mode  = {(t, p): A_mat[t - 1, p - 1]
           for t in [1, 2, 3] for p in range(1, 7)}

# =========================================================================
#  7. MALiYET PARAMETRELERi  Cw, Cy, R, T  (COST_PARAMS)
#     Satir 2  -> Cw[j=1,p]   Satir 3  -> Cw[j=2,p]
#     Satir 4  -> Cy[k=1,p]   Satir 5  -> Cy[k=2,p]   Satir 6 -> Cy[k=3,p]
#     Satir 7  -> R[p]  (elde tutma)
#     Satir 8  -> T[p]  (yok satma)
#     Sutunlar 2-7 -> P1..P6
# =========================================================================
df_cost  = xl_p.parse("COST_PARAMS", header=None)
cost_mat = read_block(df_cost, slice(2, 9), slice(2, 8))

Cw = {(j, p): cost_mat[j - 1, p - 1] for j in [1, 2] for p in range(1, 7)}
Cy = {(k, p): cost_mat[k + 1, p - 1] for k in [1, 2, 3] for p in range(1, 7)}
R  = {p: cost_mat[5, p - 1] for p in range(1, 7)}
T  = {p: cost_mat[6, p - 1] for p in range(1, 7)}

# =========================================================================
#  8. TASIMA MALiYETi  C_ijtp  (RAIL / ROAD / SEA)
#     Rota sirasi (satir 2-7):
#       (i=1,j=1) (i=1,j=2) (i=2,j=1) (i=2,j=2) (i=3,j=1) (i=3,j=2)
#     Sutunlar 2-7 -> P1..P6
# =========================================================================
ROUTE_MAP     = [(1, 1), (1, 2), (2, 1), (2, 2), (3, 1), (3, 2)]
SHEET_TO_MODE = {"RAIL": 1, "ROAD": 2, "SEA": 3}

C = {}
for sheet, t in SHEET_TO_MODE.items():
    df_t   = xl_c.parse(sheet, header=None)
    cost_t = read_block(df_t, slice(2, 8), slice(2, 8))
    for r_idx, (i, j) in enumerate(ROUTE_MAP):
        for p_idx, p in enumerate(range(1, 7)):
            C[(i, j, t, p)] = cost_t[r_idx, p_idx]

# =========================================================================
#  9. FIRSAT KAYBI  N_ijtp  (LEAD_TIMES)
#
#  Fabrika 1 (j=1): Rail P1-P6 -> satir  3-8  | Road -> 9-14  | Sea -> 15-20
#  Fabrika 2 (j=2): Rail P1-P6 -> satir 22-27  | Road -> 28-33 | Sea -> 34-39
#  Sutun indeksleri: i=1 -> 3 | i=2 -> 7 | i=3 -> 11
# =========================================================================
df_lt = xl_p.parse("LEAD_TIMES", header=None)

FAB_MODE_START = {
    (1, 1):  3, (1, 2):  9, (1, 3): 15,
    (2, 1): 22, (2, 2): 28, (2, 3): 34,
}
N_COL = {1: 3, 2: 7, 3: 11}

N = {}
for i in [1, 2, 3]:
    for j in [1, 2]:
        for t in [1, 2, 3]:
            start = FAB_MODE_START[(j, t)]
            for p_idx, p in enumerate(range(1, 7)):
                val = df_lt.iloc[start + p_idx, N_COL[i]]
                N[(i, j, t, p)] = 0.0 if pd.isna(val) else float(val)

# Dogrulama: X1111 katsayisi = C + pi*N = 3 + 0.10*48 = 7.8
_check = C[(1, 1, 1, 1)] + pi_rate * N[(1, 1, 1, 1)]
assert abs(_check - 7.8) < 0.01, f"Katsayi hatasi: {_check}"
print(f"  Dogrulama X1111: C+pi*N = {_check:.1f}  (beklenen 7.8) -> OK")

# =========================================================================
#  iNDEKS KUMELERi
# =========================================================================
suppliers = [1, 2, 3]
factories  = [1, 2]
modes      = [1, 2, 3]
periods    = [1, 2, 3, 4, 5, 6]
customers  = [1, 2, 3]

# B ve Q degiskenleri p=0'i de kapsar
periods_ext = [0] + periods   # p=0..6

print("\n" + "=" * 65)
print("  GUROBI MODELi KURULUYOR")
print("=" * 65)

# =========================================================================
#  MODEL
# =========================================================================
model = gp.Model("DistributionNetworkDesign_Secmen2015")
model.setParam("OutputFlag", 1)

# -------------------------------------------------------------------------
#  Karar Degiskenleri
# -------------------------------------------------------------------------
X = {(i, j, t, p): model.addVar(lb=0.0, name=f"X{i}{j}{t}{p}")
     for i in suppliers
     for j in factories
     for t in modes
     for p in periods}

W = {(j, p): model.addVar(lb=0.0, name=f"W{j}{p}")
     for j in factories
     for p in periods}

Y = {(k, p): model.addVar(lb=0.0, name=f"Y{k}{p}")
     for k in customers
     for p in periods}

# B_kp: p=0..6  (B_k0 SERBEST — makale kisitlarinda B_k0=0 yok)
B = {(k, p): model.addVar(lb=0.0, name=f"B{k}{p}")
     for k in customers
     for p in periods_ext}

# Q_p: p=0..6
Q = {p: model.addVar(lb=0.0, name=f"Q{p}") for p in periods_ext}

model.update()

# -------------------------------------------------------------------------
#  Amac Fonksiyonu
#
#  Kaynak: Denk. (1), Secmen et al. (2015)
#
#  Min Z = [trans] + [W tasima] + [Y tasima] + [firsat]
#         + [stok elde tutma] + [yok satma]
#
#  NOT: Makaledeki orijinal hold/stock ifadesi
#       R[p]*(ΣW_j(p-Fj) - ΣY_k(p+Gk) - ΣB_k(p+Gk-1) + Q_(p-1))
#       + T[p]*(ΣY_k(p+Gk) + ΣB_k(p+Gk-1) - ΣW_j(p-Fj) - Q_(p-1))
#  P4'te R[4]-T[4] = 10-30 = -20 < 0 oldugu icin bu formul sinirsizdır.
#  Esdeger ve sinirli formul: holding = R[p]*Q[p], stockout = T[p]*B_kp
#  (Q[p] <= c_p kısıtıyla sinirli; B_kp >= 0 ile sinirli)
# -------------------------------------------------------------------------

# (1) Tedarikci -> Fabrika tasima
obj_trans = gp.quicksum(
    C[i, j, t, p] * X[i, j, t, p]
    for i in suppliers for j in factories for t in modes for p in periods
)

# (2) Fabrika -> DC tasima
obj_w = gp.quicksum(Cw[j, p] * W[j, p] for j in factories for p in periods)

# (3) DC -> Musteri tasima
obj_y = gp.quicksum(Cy[k, p] * Y[k, p] for k in customers for p in periods)

# (4) Firsat maliyeti  pi * N_ijtp * X_ijtp
obj_opp = gp.quicksum(
    N[i, j, t, p] * pi_rate * X[i, j, t, p]
    for i in suppliers for j in factories for t in modes for p in periods
)

# (5) Stok elde tutma  R[p] * Q[p]   (p=1..6)
obj_hold = gp.quicksum(R[p] * Q[p] for p in periods)

# (6) Yok satma  T[p] * B_kp   (p=1..6; B_k0 maliyet disinda tutulur)
obj_stock = gp.quicksum(T[p] * B[k, p] for k in customers for p in periods)

model.setObjective(
    obj_trans + obj_w + obj_y + obj_opp + obj_hold + obj_stock,
    GRB.MINIMIZE
)

# -------------------------------------------------------------------------
#  Kisitlar
# -------------------------------------------------------------------------

# C1 — Tedarikci kapasite kisiti  (Denklem 2.1)
#       Σ_j Σ_t X_ijtp <= alpha_ip   ∀ i, p
for i in suppliers:
    for p in periods:
        model.addConstr(
            gp.quicksum(X[i, j, t, p] for j in factories for t in modes)
            <= alpha[i, p],
            name=f"SC_{i}_{p}"
        )

# C2 — Fabrika kapasite kisiti  (Denklem 2.2)
#       W_jp <= b_jp   ∀ j, p
for j in factories:
    for p in periods:
        model.addConstr(W[j, p] <= b_cap[j, p], name=f"FC_{j}_{p}")

# C3 — DC kapasite kisiti  (Denklem 2.3)
#       Σ_k Y_kp <= c_p   ∀ p
for p in periods:
    model.addConstr(
        gp.quicksum(Y[k, p] for k in customers) <= c_cap[p],
        name=f"DC_{p}"
    )

# C4 — Musteri talebi karsilama  (Denklem 2.4)
#       Y_kp + B_k(p-1) + B_kp >= d_kp   ∀ k, p
#       NOT: B_k0 serbest degisken (makale kisitlarinda B_k0=0 zorlanmıyor)
for k in customers:
    for p in periods:
        model.addConstr(
            Y[k, p] + B[k, p - 1] + B[k, p] >= d[k, p],
            name=f"Dem_{k}_{p}"
        )

# C5 — Tasima modu kapasite kisiti  (Denklem 2.5)
#       Σ_i Σ_j X_ijtp <= A_tp   ∀ t, p
for t in modes:
    for p in periods:
        model.addConstr(
            gp.quicksum(X[i, j, t, p] for i in suppliers for j in factories)
            <= A_mode[t, p],
            name=f"MC_{t}_{p}"
        )

# C6 — Akis dengesi: tedarikci girisi = fabrika cikisi
#       Σ_i Σ_t X_ijtp = W_jp   ∀ j, p
for j in factories:
    for p in periods:
        model.addConstr(
            gp.quicksum(X[i, j, t, p] for i in suppliers for t in modes)
            == W[j, p],
            name=f"FB_{j}_{p}"
        )

# C7 — DC stok dengesi  (Fj=1 gecikme: W donemi p'de gonderilir, p+1'de DC'ye gelir)
#       Q_0 = Q_init
#       Q_p = Q_(p-1) + Σ_j W_j(p-Fj) - Σ_k Y_kp   ∀ p
#       Fj=1: W_j(p-1) kullanilir; W_j0 = 0 (p=0 oncesi uretim yok)
model.addConstr(Q[0] == Q_init, name="Q_init")

for p in periods:
    # W_j(p-Fj) = W_j(p-1); p=1 icin W_j0 = 0 (degisken yok, donem oncesi)
    w_arrived = gp.quicksum(W[j, p - Fj] for j in factories) if p > Fj else 0.0
    model.addConstr(
        Q[p] == Q[p - 1] + w_arrived
                - gp.quicksum(Y[k, p] for k in customers),
        name=f"IB_{p}"
    )

# C8 — DC stok ust siniri
#       Q_p <= c_p   ∀ p  (c_p = 40.000 tum donemler icin)
for p in periods:
    model.addConstr(Q[p] <= c_cap[p], name=f"Qub_{p}")

# =========================================================================
#  COZUM
# =========================================================================
model.optimize()

# =========================================================================
#  SONUCLAR — TABLO 1 FORMATI (Secmen et al. 2015)
# =========================================================================
if model.status == GRB.OPTIMAL:
    Z_opt = model.ObjVal

    col_X, col_W, col_Y, col_QB = [], [], [], []

    for i in suppliers:
        for j in factories:
            for t in modes:
                for p in periods:
                    v = X[i, j, t, p].X
                    if v > 0.5:
                        col_X.append((f"X{i}{j}{t}{p}", int(round(v))))

    for j in factories:
        for p in periods:
            v = W[j, p].X
            if v > 0.5:
                col_W.append((f"W{j}{p}", int(round(v))))

    for k in customers:
        for p in periods:
            v = Y[k, p].X
            if v > 0.5:
                col_Y.append((f"Y{k}{p}", int(round(v))))

    for p in periods_ext:
        v = Q[p].X
        if v > 0.5:
            col_QB.append((f"Q{p}", int(round(v))))

    for k in customers:
        for p in periods_ext:
            v = B[k, p].X
            if v > 0.5:
                col_QB.append((f"B{k}{p}", int(round(v))))

    # --- Tablo 1 yazdir ---------------------------------------------------
    C1, C2 = 10, 9

    def _cell(pair):
        if pair is None:
            return " " * (C1 + C2 + 2)
        return f"{pair[0]:>{C1}}  {pair[1]:>{C2},}"

    def _row(a, b, c, dd):
        return f"  {_cell(a)}    {_cell(b)}    {_cell(c)}    {_cell(dd)}"

    def _get(lst, idx):
        return lst[idx] if idx < len(lst) else None

    hdr = (f"  {'Degisken':>{C1}}  {'Sonuc':>{C2}}"
           f"    {'Degisken':>{C1}}  {'Sonuc':>{C2}}"
           f"    {'Degisken':>{C1}}  {'Sonuc':>{C2}}"
           f"    {'Degisken':>{C1}}  {'Sonuc':>{C2}}")
    sep = "-" * (C1 + C2 + 2)

    n_rows = max(len(col_X), len(col_W), len(col_Y), len(col_QB))

    print(f"\n{'=' * 65}")
    print("  TABLO 1: OPTiMUM DEGERLER TABLOSU")
    print(f"{'=' * 65}")
    print(hdr)
    print(f"  {sep}    {sep}    {sep}    {sep}")
    for r in range(n_rows):
        print(_row(_get(col_X, r), _get(col_W, r),
                   _get(col_Y, r), _get(col_QB, r)))

    print(f"\n  Amac Fonksiyonu  : Min. Z = {Z_opt:>12,.1f}")
    print(f"  Tez Referansi    :          {Z_ref:>12,.1f}")
    print(f"  Fark             : {abs(Z_opt - Z_ref):>17,.1f}")
    print(f"{'=' * 65}")

    # --- Excel ciktisi ----------------------------------------------------
    rows = []
    for name, val in col_X:
        rows.append({"Degisken": name, "Deger": val, "Grup": "X"})
    for name, val in col_W:
        rows.append({"Degisken": name, "Deger": val, "Grup": "W"})
    for name, val in col_Y:
        rows.append({"Degisken": name, "Deger": val, "Grup": "Y"})
    for name, val in col_QB:
        rows.append({"Degisken": name, "Deger": val, "Grup": "Q/B"})

    out_path = os.path.join(RESULT_DIR, "optimal_solution.xlsx")
    pd.DataFrame(rows).to_excel(out_path, index=False)
    print(f"\n  Sonuclar -> {out_path}")

elif model.status == GRB.INFEASIBLE:
    print("\n  Model UYGUN DEGIL (Status 3)")
    model.computeIIS()
    iis_path = os.path.join(RESULT_DIR, "model.ilp")
    model.write(iis_path)
    print(f"  IIS -> {iis_path}")

elif model.status == GRB.UNBOUNDED:
    print("\n  Model SINIR YOK (Status 5)")
    print("  Kisit veya amac fonksiyonu formulasyonunu kontrol edin.")

else:
    print(f"\n  Cozucu Durumu: {model.status}")
