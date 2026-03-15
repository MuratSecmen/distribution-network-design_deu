"""
Cok Urunlu, Cok DC'li Dagiticim Agi Tasarimi — MILP Genisletmesi
=================================================================
Temel model: Secmen, Oncan, Tuna (Beykoz Akademi Dergisi, 2015)

Orijinal modele eklenenler:
  1. Cok urun  (n in N = {1,2,3})
  2. Cok DC    (l in L = {1,2,3}: Ukrayna, Polonya, Romanya)
  3. Dinamik DC acma/kapama karari  z_lp in {0,1}
  4. DC sabit acilis maliyeti  f_lp
  5. DC gecis maliyeti (switching)  sc_l * |z_lp - z_l(p-1)|

Hesapsal zorluk ozellikleri:
  - Binary degiskenler: z_lp (3 DC x 6 donem = 18 binary)
  - Big-M bagli kapasite kisitlari
  - Gecis maliyeti lineerlestirilmis (linearization)
  - Valid inequality: en az 1 DC her donem acik olmali
  - Symmetry breaking: DC'ler kapasiteye gore siralanmis

Ag:
    Tedarikciler : T1=Rusya(i=1), T2=Almanya(i=2), T3=Norvec(i=3)
    Fabrikalar   : F1=Ispanya(j=1), F2=Rusya(j=2)
    DC'ler       : L1=Ukrayna(l=1), L2=Polonya(l=2), L3=Romanya(l=3)
    Musteriler   : M1=Finlandiya(k=1), M2=Turkiye(k=2), M3=Rusya(k=3)
    Urunler      : UrunA(n=1), UrunB(n=2), UrunC(n=3)
    Modlar       : Tren(t=1), Kara(t=2), Deniz(t=3)
    Donemler     : p=1..6

Karar Degiskenleri:
    X_ijtnp : Tedarikci i -> Fab j, mod t, urun n, donem p
    W_jlnp  : Fab j -> DC l, urun n, donem p
    Y_klnp  : DC l -> Musteri k, urun n, donem p
    B_knp   : Karsilanamayan talep, musteri k, urun n, donem p  (p=0..6)
    Q_lnp   : DC l stoku, urun n, donem p  (p=0..6)
    z_lp    : DC l acik mi donem p'de?  (binary)
    delta_lp: Gecis degiskeni (lineerizasyon icin)  (binary)
"""

import gurobipy as gp
from gurobipy import GRB
import pandas as pd
import numpy as np
import os
import sys
import time

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

print("=" * 70)
print("  COK URUNLU - COK DC - DINAMIK ACMA/KAPAMA  (MILP)")
print("=" * 70)

xl_p = pd.ExcelFile(FILE_PARAM)
xl_c = pd.ExcelFile(FILE_TRANS)


def read_block(df, row_slice, col_slice):
    return df.iloc[row_slice, col_slice].astype(float).values


# =========================================================================
#  INDEKS KUMELERI
# =========================================================================
suppliers = [1, 2, 3]
factories  = [1, 2]
dcs        = [1, 2, 3]          # L1=Ukrayna, L2=Polonya, L3=Romanya
customers  = [1, 2, 3]
modes      = [1, 2, 3]
products   = [1, 2, 3]          # UrunA, UrunB, UrunC
periods    = [1, 2, 3, 4, 5, 6]
periods_ext = [0] + periods

# =========================================================================
#  TEMEL PARAMETRELER  (Excel'den)
# =========================================================================
df_sc   = xl_p.parse("SCALAR_PARAMS", header=None)
pi_rate = float(df_sc.iloc[2, 2])   # 0.10
Q_init  = float(df_sc.iloc[3, 2])   # 40000  (DC basina)
Fj      = int(df_sc.iloc[4, 2])     # 1

print(f"\n  pi={pi_rate}  Q0(DC basina)={Q_init:,.0f}  Fj={Fj}")

# Tedarikci kapasitesi (urun bazli: toplam kap / urun sayisi x agirlik)
_alpha_mat = read_block(xl_p.parse("SUPPLIER_CAP", header=None), slice(2, 5), slice(2, 8))
# Urun bazli kapasite: A paylari %40, %35, %25
_prod_share = {1: 0.40, 2: 0.35, 3: 0.25}
alpha = {(i, n, p): _alpha_mat[i - 1, p - 1] * _prod_share[n]
         for i in suppliers for n in products for p in periods}

# Fabrika kapasitesi (toplam, urunler arasi paylasimli)
_b_mat = read_block(xl_p.parse("FACTORY_CAP", header=None), slice(2, 4), slice(2, 8))
b_cap  = {(j, p): _b_mat[j - 1, p - 1] for j in factories for p in periods}

# Musteri talebi (urun bazli: toplam talep / urun agirlikli)
_d_mat = read_block(xl_p.parse("DEMAND", header=None), slice(2, 5), slice(2, 8))
demand = {(k, n, p): _d_mat[k - 1, p - 1] * _prod_share[n]
          for k in customers for n in products for p in periods}

# Mod kapasitesi (urunler arasi paylasimli toplam)
_A_mat = read_block(xl_p.parse("MODE_CAP", header=None), slice(2, 5), slice(2, 8))
A_mode = {(t, p): _A_mat[t - 1, p - 1] for t in modes for p in periods}

# Maliyet parametreleri
_cm  = read_block(xl_p.parse("COST_PARAMS", header=None), slice(2, 9), slice(2, 8))
R    = {p: _cm[5, p - 1] for p in periods}
T    = {p: _cm[6, p - 1] for p in periods}

# Taşıma maliyeti C_ijtp (urun bazli agirliklandirilmis)
_prod_cost = {1: 1.00, 2: 1.30, 3: 0.80}   # B daha degerli, C daha ucuz
_route_map = [(1,1),(1,2),(2,1),(2,2),(3,1),(3,2)]
C_base = {}
for sheet, t in [("RAIL", 1), ("ROAD", 2), ("SEA", 3)]:
    ct = read_block(xl_c.parse(sheet, header=None), slice(2, 8), slice(2, 8))
    for r, (i, j) in enumerate(_route_map):
        for pp, p in enumerate(periods):
            C_base[(i, j, t, p)] = ct[r, pp]

C = {(i, j, t, n, p): C_base[(i, j, t, p)] * _prod_cost[n]
     for i in suppliers for j in factories
     for t in modes for n in products for p in periods}

# Firsat kaybi N (urun bagimsiz — orijinal degerler)
df_lt = xl_p.parse("LEAD_TIMES", header=None)
_fms  = {(1,1):3,(1,2):9,(1,3):15,(2,1):22,(2,2):28,(2,3):34}
_nc   = {1:3, 2:7, 3:11}
N_opp = {}
for i in suppliers:
    for j in factories:
        for t in modes:
            s = _fms[(j, t)]
            for pp, p in enumerate(periods):
                val = df_lt.iloc[s + pp, _nc[i]]
                N_opp[(i, j, t, p)] = 0.0 if pd.isna(val) else float(val)

N = {(i, j, t, n, p): N_opp[(i, j, t, p)]
     for i in suppliers for j in factories
     for t in modes for n in products for p in periods}

# =========================================================================
#  GENISLETME PARAMETRELERi — COKLU DC
# =========================================================================

# DC kapasiteleri (toplam, urunler paylasir)
dc_cap = {
    (1, p): 40000 for p in periods  # Ukrayna: buyuk
}
dc_cap.update({(2, p): 35000 for p in periods})   # Polonya: orta
dc_cap.update({(3, p): 30000 for p in periods})   # Romanya: kucuk

# DC baslangic stoku (Q0) — urun bazli
Q0 = {(l, n): Q_init * _prod_share[n] if l == 1 else 0.0
      for l in dcs for n in products}
# Ukrayna mevcut DC, diger ikisi bos baslar

# DC sabit acilis maliyeti ($/donem, acik oldugunda odenecek)
f_dc = {(1, p): 5_000 for p in periods}   # Ukrayna ucuz (kurulu)
f_dc.update({(2, p): 8_000 for p in periods})   # Polonya
f_dc.update({(3, p): 6_500 for p in periods})   # Romanya

# DC gecis (switching) maliyeti — acma veya kapama basina
sc_dc = {1: 3_000, 2: 5_000, 3: 4_000}

# Fabrika -> DC birim tasima maliyeti  Cw[j,l,n,p]
_Cw_base = {
    (1, 1): [20, 30, 20, 30, 40, 20],  # Fab1 -> Ukrayna
    (1, 2): [25, 35, 25, 35, 45, 25],  # Fab1 -> Polonya
    (1, 3): [22, 32, 22, 32, 42, 22],  # Fab1 -> Romanya
    (2, 1): [30, 20, 40, 20, 30, 20],  # Fab2 -> Ukrayna
    (2, 2): [35, 25, 45, 25, 35, 25],  # Fab2 -> Polonya
    (2, 3): [28, 22, 38, 22, 32, 22],  # Fab2 -> Romanya
}
Cw = {(j, l, n, p): _Cw_base[(j, l)][p - 1] * _prod_cost[n]
      for j in factories for l in dcs
      for n in products for p in periods}

# DC -> Musteri birim tasima maliyeti  Cy[l,k,n,p]
_Cy_base = {
    (1, 1): [40, 20, 30, 30, 40, 20],  # Ukrayna -> Finlandiya
    (2, 1): [35, 18, 28, 28, 38, 18],  # Polonya -> Finlandiya (yakin)
    (3, 1): [42, 22, 32, 32, 42, 22],  # Romanya -> Finlandiya
    (1, 2): [30, 40, 30, 20, 40, 20],  # Ukrayna -> Turkiye
    (2, 2): [38, 45, 35, 25, 45, 25],  # Polonya -> Turkiye
    (3, 2): [25, 35, 25, 18, 35, 18],  # Romanya -> Turkiye (yakin)
    (1, 3): [30, 20, 40, 20, 30, 20],  # Ukrayna -> Rusya
    (2, 3): [32, 22, 42, 22, 32, 22],  # Polonya -> Rusya
    (3, 3): [28, 18, 38, 18, 28, 18],  # Romanya -> Rusya (yakin)
}
Cy = {(l, k, n, p): _Cy_base[(l, k)][p - 1] * _prod_cost[n]
      for l in dcs for k in customers
      for n in products for p in periods}

# Big-M: DC l icin maksimum toplam giris/cikis
M_dc = {l: sum(b_cap[j, p] for j in factories for p in periods) for l in dcs}

print(f"\n  Urunler      : {len(products)} ({', '.join([f'n={n}' for n in products])})")
print(f"  DC'ler       : {len(dcs)} ({', '.join(['Ukrayna','Polonya','Romanya'])})")
print(f"  Binary z_lp  : {len(dcs) * len(periods)} adet")
print(f"  Binary delta : {len(dcs) * len(periods)} adet (switching linearization)")

# =========================================================================
#  MODEL
# =========================================================================
print("\n" + "=" * 70)
print("  GUROBI MILP MODELi KURULUYOR")
print("=" * 70)

model = gp.Model("MultiProduct_MultiDC_MILP")
model.setParam("OutputFlag", 1)
model.setParam("MIPGap", 0.01)        # %1 MIP gap hedefi
model.setParam("TimeLimit", 300)      # 5 dakika zaman siniri
model.setParam("Cuts", 2)             # agresif cut generation
model.setParam("Presolve", 2)         # agresif presolve

# -------------------------------------------------------------------------
#  Karar Degiskenleri
# -------------------------------------------------------------------------
X = {(i, j, t, n, p): model.addVar(lb=0, name=f"X{i}{j}{t}{n}{p}")
     for i in suppliers for j in factories
     for t in modes for n in products for p in periods}

W = {(j, l, n, p): model.addVar(lb=0, name=f"W{j}{l}{n}{p}")
     for j in factories for l in dcs
     for n in products for p in periods}

Y = {(l, k, n, p): model.addVar(lb=0, name=f"Y{l}{k}{n}{p}")
     for l in dcs for k in customers
     for n in products for p in periods}

B = {(k, n, p): model.addVar(lb=0, name=f"B{k}{n}{p}")
     for k in customers for n in products for p in periods_ext}

Q = {(l, n, p): model.addVar(lb=0, name=f"Q{l}{n}{p}")
     for l in dcs for n in products for p in periods_ext}

# Binary: DC l donem p'de acik mi?
z = {(l, p): model.addVar(vtype=GRB.BINARY, name=f"z{l}{p}")
     for l in dcs for p in periods}

# Binary: switching linearization  delta_lp = max(z_lp - z_l(p-1), 0)
#   gecis = DC kapali->acik ya da acik->kapali
delta = {(l, p): model.addVar(vtype=GRB.BINARY, name=f"delta{l}{p}")
         for l in dcs for p in periods}

model.update()

# -------------------------------------------------------------------------
#  Amac Fonksiyonu
# -------------------------------------------------------------------------

# (1) Tedarikci -> Fabrika tasima + firsat maliyeti
obj_trans = gp.quicksum(
    (C[i, j, t, n, p] + N[i, j, t, n, p] * pi_rate) * X[i, j, t, n, p]
    for i in suppliers for j in factories
    for t in modes for n in products for p in periods
)

# (2) Fabrika -> DC tasima maliyeti
obj_w = gp.quicksum(
    Cw[j, l, n, p] * W[j, l, n, p]
    for j in factories for l in dcs for n in products for p in periods
)

# (3) DC -> Musteri tasima maliyeti
obj_y = gp.quicksum(
    Cy[l, k, n, p] * Y[l, k, n, p]
    for l in dcs for k in customers for n in products for p in periods
)

# (4) Stok elde tutma maliyeti  R[p] * Σ_n Q[l,n,p]
obj_hold = gp.quicksum(
    R[p] * Q[l, n, p]
    for l in dcs for n in products for p in periods
)

# (5) Yok satma maliyeti  T[p] * B[k,n,p]
obj_stock = gp.quicksum(
    T[p] * B[k, n, p]
    for k in customers for n in products for p in periods
)

# (6) DC sabit acilis maliyeti  f_lp * z_lp
obj_fixed = gp.quicksum(
    f_dc[l, p] * z[l, p]
    for l in dcs for p in periods
)

# (7) DC gecis (switching) maliyeti  sc_l * delta_lp
obj_switch = gp.quicksum(
    sc_dc[l] * delta[l, p]
    for l in dcs for p in periods
)

model.setObjective(
    obj_trans + obj_w + obj_y + obj_hold + obj_stock + obj_fixed + obj_switch,
    GRB.MINIMIZE
)

# -------------------------------------------------------------------------
#  Kisitlar
# -------------------------------------------------------------------------

# C1 — Tedarikci kapasite (urun bazli)
for i in suppliers:
    for n in products:
        for p in periods:
            model.addConstr(
                gp.quicksum(X[i, j, t, n, p]
                            for j in factories for t in modes)
                <= alpha[i, n, p],
                name=f"SC_{i}_{n}_{p}"
            )

# C2 — Fabrika kapasite (toplam, tum urunler ve DC'ler icin)
for j in factories:
    for p in periods:
        model.addConstr(
            gp.quicksum(W[j, l, n, p]
                        for l in dcs for n in products)
            <= b_cap[j, p],
            name=f"FC_{j}_{p}"
        )

# C3 — DC kapasite (binary ile bagli: Big-M)
#       Σ_n Q_lnp <= dc_cap[l,p] * z_lp
for l in dcs:
    for p in periods:
        model.addConstr(
            gp.quicksum(Q[l, n, p] for n in products)
            <= dc_cap[l, p] * z[l, p],
            name=f"DC_cap_{l}_{p}"
        )

# C4 — DC'ye giris yalnizca acikken (Big-M)
#       Σ_j Σ_n W_jlnp <= M_l * z_lp
for l in dcs:
    for p in periods:
        model.addConstr(
            gp.quicksum(W[j, l, n, p]
                        for j in factories for n in products)
            <= M_dc[l] * z[l, p],
            name=f"DC_in_{l}_{p}"
        )

# C5 — DC'den cikis yalnizca acikken (Big-M)
#       Σ_k Σ_n Y_lknp <= M_l * z_lp
for l in dcs:
    for p in periods:
        model.addConstr(
            gp.quicksum(Y[l, k, n, p]
                        for k in customers for n in products)
            <= M_dc[l] * z[l, p],
            name=f"DC_out_{l}_{p}"
        )

# C6 — Musteri talebi karsilama
#       Σ_l Y_lknp + B_kn(p-1) + B_knp >= d_knp  ∀k,n,p
for k in customers:
    for n in products:
        for p in periods:
            model.addConstr(
                gp.quicksum(Y[l, k, n, p] for l in dcs)
                + B[k, n, p - 1] + B[k, n, p]
                >= demand[k, n, p],
                name=f"Dem_{k}_{n}_{p}"
            )

# C7 — Mod kapasite (tum urunler, tum fabrika-DC rotalar)
for t in modes:
    for p in periods:
        model.addConstr(
            gp.quicksum(X[i, j, t, n, p]
                        for i in suppliers for j in factories for n in products)
            <= A_mode[t, p],
            name=f"MC_{t}_{p}"
        )

# C8 — Fabrika akis dengesi (urun bazli)
#       Σ_i Σ_t X_ijtnp = Σ_l W_jlnp  ∀j,n,p
for j in factories:
    for n in products:
        for p in periods:
            model.addConstr(
                gp.quicksum(X[i, j, t, n, p]
                            for i in suppliers for t in modes)
                == gp.quicksum(W[j, l, n, p] for l in dcs),
                name=f"FB_{j}_{n}_{p}"
            )

# C9 — DC stok dengesi (Fj=1 gecikme ile)
#       Q_lnp = Q_ln(p-1) + Σ_j W_jln(p-Fj) - Σ_k Y_lknp  ∀l,n,p
for l in dcs:
    for n in products:
        # Baslangic stoku
        model.addConstr(Q[l, n, 0] == Q0[l, n], name=f"Q0_{l}_{n}")
        for p in periods:
            w_arrived = (gp.quicksum(W[j, l, n, p - Fj] for j in factories)
                         if p > Fj else 0.0)
            model.addConstr(
                Q[l, n, p] == Q[l, n, p - 1] + w_arrived
                              - gp.quicksum(Y[l, k, n, p] for k in customers),
                name=f"IB_{l}_{n}_{p}"
            )

# C10 — DC stok ust siniri (binary ile bagli)
#        Q_lnp <= dc_cap[l,p] * z_lp (C3 zaten bunu sagliyor ama per-product ek)
for l in dcs:
    for n in products:
        for p in periods:
            model.addConstr(
                Q[l, n, p] <= dc_cap[l, p],
                name=f"Qub_{l}_{n}_{p}"
            )

# C11 — Switching lineerizasyonu
#        delta_lp >= z_lp - z_l(p-1)  (acma gostergesi)
#        delta_lp >= z_l(p-1) - z_lp  (kapama gostergesi)
#        delta_lp <= z_lp + z_l(p-1)  (her ikisi de 0 ise delta=0)
for l in dcs:
    for p in periods:
        if p == 1:
            # P1'de onceki donem referansi yok;
            # Ukrayna (l=1) baslangiçta acik kabul edilir
            z_prev = 1.0 if l == 1 else 0.0
            model.addConstr(delta[l, p] >= z[l, p] - z_prev, name=f"SW1a_{l}_{p}")
            model.addConstr(delta[l, p] >= z_prev - z[l, p], name=f"SW1b_{l}_{p}")
        else:
            model.addConstr(delta[l, p] >= z[l, p] - z[l, p - 1], name=f"SWa_{l}_{p}")
            model.addConstr(delta[l, p] >= z[l, p - 1] - z[l, p], name=f"SWb_{l}_{p}")
            model.addConstr(delta[l, p] <= z[l, p] + z[l, p - 1],  name=f"SWc_{l}_{p}")

# -------------------------------------------------------------------------
#  Valid Inequalities — Hesapsal performansi artirmak icin
# -------------------------------------------------------------------------

# VI1: Her donemde en az 1 DC acik olmali
for p in periods:
    model.addConstr(
        gp.quicksum(z[l, p] for l in dcs) >= 1,
        name=f"VI_minDC_{p}"
    )

# VI2: Toplam DC kapasitesi toplam talebi karsilamali
for p in periods:
    total_demand = sum(demand[k, n, p] for k in customers for n in products)
    model.addConstr(
        gp.quicksum(dc_cap[l, p] * z[l, p] for l in dcs) >= total_demand,
        name=f"VI_cap_{p}"
    )

# VI3: Symmetry breaking — DC'leri kapasiteye gore sirala
#       En buyuk DC (l=1, kap=40000) en kucukten once kapanmali
#       Eger l=1 kapali ise l=2 de kapali olmali
for p in periods:
    model.addConstr(z[2, p] <= z[1, p], name=f"SB_12_{p}")
    model.addConstr(z[3, p] <= z[2, p], name=f"SB_23_{p}")

# =========================================================================
#  COZUM
# =========================================================================
print("\n  Model ozeti:")
print(f"    Degiskenler : {model.NumVars}")
print(f"    Binary      : {model.NumBinVars}")
print(f"    Kisitlar    : {model.NumConstrs}")
print(f"\n  Gurobi cozuyor... (hedef MIP gap <= %1, limit 300s)\n")

start_time = time.time()
model.optimize()
elapsed = time.time() - start_time

# =========================================================================
#  SONUCLAR
# =========================================================================
def print_separator(char="=", width=70):
    print(char * width)


if model.status in [GRB.OPTIMAL, GRB.TIME_LIMIT] and model.SolCount > 0:
    Z_opt   = model.ObjVal
    Z_bound = model.ObjBound
    gap     = model.MIPGap * 100

    print_separator()
    print("  COZUM SONUCLARI")
    print_separator()
    print(f"  Status       : {'OPTIMAL' if model.status == GRB.OPTIMAL else 'TIME LIMIT (en iyi)'}")
    print(f"  Obj. Degeri  : {Z_opt:>15,.1f}")
    print(f"  Alt Sinir    : {Z_bound:>15,.1f}")
    print(f"  MIP Gap      : %{gap:.2f}")
    print(f"  Sure         : {elapsed:.1f}s")
    print(f"  Node Sayisi  : {int(model.NodeCount):,}")
    print_separator()

    # --- Maliyet bileşenleri ---
    z_trans_val = sum(
        (C[i,j,t,n,p] + N[i,j,t,n,p]*pi_rate) * X[i,j,t,n,p].X
        for i in suppliers for j in factories for t in modes
        for n in products for p in periods
    )
    z_w_val = sum(Cw[j,l,n,p]*W[j,l,n,p].X
                  for j in factories for l in dcs for n in products for p in periods)
    z_y_val = sum(Cy[l,k,n,p]*Y[l,k,n,p].X
                  for l in dcs for k in customers for n in products for p in periods)
    z_hold_val  = sum(R[p]*Q[l,n,p].X
                      for l in dcs for n in products for p in periods)
    z_stock_val = sum(T[p]*B[k,n,p].X
                      for k in customers for n in products for p in periods)
    z_fixed_val = sum(f_dc[l,p]*z[l,p].X for l in dcs for p in periods)
    z_sw_val    = sum(sc_dc[l]*delta[l,p].X for l in dcs for p in periods)

    print("\n  MALIYET BILESENLERI:")
    print(f"    Tasima (X+opp)   : {z_trans_val:>12,.1f}")
    print(f"    Fab->DC (W)      : {z_w_val:>12,.1f}")
    print(f"    DC->Mus (Y)      : {z_y_val:>12,.1f}")
    print(f"    Stok tutma       : {z_hold_val:>12,.1f}")
    print(f"    Yok satma        : {z_stock_val:>12,.1f}")
    print(f"    DC sabit maliyet : {z_fixed_val:>12,.1f}")
    print(f"    DC gecis maliyeti: {z_sw_val:>12,.1f}")
    print(f"    {'─'*30}")
    print(f"    TOPLAM           : {Z_opt:>12,.1f}")

    # --- DC acma/kapama plani ---
    print("\n  DC ACMA/KAPAMA PLANI  (z_lp):")
    dc_names = {1: "Ukrayna", 2: "Polonya", 3: "Romanya"}
    header = f"  {'DC':>10} |" + "".join(f" P{p:1d}  " for p in periods)
    print(header)
    print(f"  {'-'*10}-" + "-" * (len(periods)*5))
    for l in dcs:
        row = f"  {dc_names[l]:>10} |"
        for p in periods:
            val = round(z[l, p].X)
            row += f" {'ACIK' if val else 'KPLI':4} "
        print(row)

    # --- Urun bazli X akislari (ozet) ---
    print("\n  URUN BAZLI TEDARIKCI -> FABRIKA AKISLARI (X, ozet):")
    _s  = {1:"Rusya", 2:"Almanya", 3:"Norvec"}
    _f  = {1:"Ispanya", 2:"Rusya"}
    _m  = {1:"Tren", 2:"Kara", 3:"Deniz"}
    _pr = {1:"UrunA", 2:"UrunB", 3:"UrunC"}
    for n in products:
        total_n = sum(X[i,j,t,n,p].X
                      for i in suppliers for j in factories
                      for t in modes for p in periods)
        if total_n > 0.5:
            print(f"    {_pr[n]}: {total_n:,.0f} birim toplam")
            for i in suppliers:
                for j in factories:
                    for t in modes:
                        tot = sum(X[i,j,t,n,p].X for p in periods)
                        if tot > 0.5:
                            print(f"      {_s[i]:>8} -> {_f[j]:<8} [{_m[t]:<5}]: {tot:>10,.0f}")

    # --- W akislari (DC bazli ozet) ---
    print("\n  FABRiKA -> DC AKiSLARI (W, urun toplami):")
    for l in dcs:
        for p in periods:
            if round(z[l, p].X) == 1:
                w_total = sum(W[j,l,n,p].X
                              for j in factories for n in products)
                if w_total > 0.5:
                    print(f"    DC{l}({dc_names[l]:>8}), P{p}: {w_total:>10,.0f}")

    # =========================================================================
    #  EXCEL WORKBOOK — Her karar degiskeni icin ayri sekme + ozet + maliyet
    # =========================================================================
    out_path = os.path.join(RESULT_DIR, "milp_solution.xlsx")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:

        # ── Yardimci veriler ──────────────────────────────────────────────
        _sn  = {1:"Rusya(T1)",   2:"Almanya(T2)", 3:"Norvec(T3)"}
        _fn  = {1:"Ispanya(F1)", 2:"Rusya(F2)"}
        _mn  = {1:"Tren",        2:"Kara",         3:"Deniz"}
        _dn  = {1:"Ukrayna(L1)", 2:"Polonya(L2)",  3:"Romanya(L3)"}
        _kn  = {1:"Finlandiya(M1)", 2:"Turkiye(M2)", 3:"Rusya(M3)"}
        _prn = {1:"UrunA",       2:"UrunB",        3:"UrunC"}

        # ── SEKME 1: OZET ──────────────────────────────────────────────────
        summary_rows = [
            ["COZUM OZETL", ""],
            ["Status",     "OPTIMAL" if model.status == GRB.OPTIMAL else "TIME LIMIT"],
            ["Obj. Degeri", round(Z_opt, 1)],
            ["Alt Sinir",   round(Z_bound, 1)],
            ["MIP Gap (%)", round(gap, 2)],
            ["Sure (s)",    round(elapsed, 1)],
            ["Node Sayisi", int(model.NodeCount)],
            ["", ""],
            ["MALIYET BILESENLERI", "Deger ($)"],
            ["Tasima + Firsat (X)",  round(z_trans_val, 1)],
            ["Fab->DC Tasima (W)",   round(z_w_val, 1)],
            ["DC->Mus Tasima (Y)",   round(z_y_val, 1)],
            ["Stok Tutma",           round(z_hold_val, 1)],
            ["Yok Satma",            round(z_stock_val, 1)],
            ["DC Sabit Maliyet",     round(z_fixed_val, 1)],
            ["DC Gecis Maliyeti",    round(z_sw_val, 1)],
            ["TOPLAM",               round(Z_opt, 1)],
        ]
        pd.DataFrame(summary_rows, columns=["Parametre", "Deger"]).to_excel(
            writer, sheet_name="OZET", index=False
        )

        # ── SEKME 2: X — Tedarikci -> Fabrika akislari ───────────────────
        # Her satir: rota + donem bazli miktar + tasima maliyeti + firsat maliyeti
        x_rows = []
        for i in suppliers:
            for j in factories:
                for t in modes:
                    for n in products:
                        for p in periods:
                            v = round(X[i, j, t, n, p].X)
                            c_unit  = C[i, j, t, n, p]
                            n_unit  = N[i, j, t, n, p]
                            c_total = c_unit * v
                            n_total = n_unit * pi_rate * v
                            x_rows.append({
                                "Degisken"       : f"X{i}{j}{t}{n}{p}",
                                "Tedarikci"      : _sn[i],
                                "Fabrika"        : _fn[j],
                                "Mod"            : _mn[t],
                                "Urun"           : _prn[n],
                                "Donem"          : p,
                                "Miktar (birim)" : v,
                                "C_birim ($/b)"  : round(c_unit, 2),
                                "N_birim (saat)" : round(n_unit, 2),
                                "Tasima_Mal ($)" : round(c_total, 1),
                                "Firsat_Mal ($)" : round(n_total, 1),
                                "Toplam_Mal ($)" : round(c_total + n_total, 1),
                            })
        df_x = pd.DataFrame(x_rows)
        df_x.to_excel(writer, sheet_name="X_Ted_Fab", index=False)

        # ── SEKME 3: W — Fabrika -> DC akislari ──────────────────────────
        w_rows = []
        for j in factories:
            for l in dcs:
                for n in products:
                    for p in periods:
                        v      = round(W[j, l, n, p].X)
                        c_unit = Cw[j, l, n, p]
                        w_rows.append({
                            "Degisken"       : f"W{j}{l}{n}{p}",
                            "Fabrika"        : _fn[j],
                            "DC"             : _dn[l],
                            "Urun"           : _prn[n],
                            "Donem"          : p,
                            "DC_Acik"        : round(z[l, p].X),
                            "Miktar (birim)" : v,
                            "Cw_birim ($/b)" : round(c_unit, 2),
                            "Tasima_Mal ($)" : round(c_unit * v, 1),
                        })
        pd.DataFrame(w_rows).to_excel(writer, sheet_name="W_Fab_DC", index=False)

        # ── SEKME 4: Y — DC -> Musteri akislari ──────────────────────────
        y_rows = []
        for l in dcs:
            for k in customers:
                for n in products:
                    for p in periods:
                        v      = round(Y[l, k, n, p].X)
                        c_unit = Cy[l, k, n, p]
                        y_rows.append({
                            "Degisken"       : f"Y{l}{k}{n}{p}",
                            "DC"             : _dn[l],
                            "Musteri"        : _kn[k],
                            "Urun"           : _prn[n],
                            "Donem"          : p,
                            "DC_Acik"        : round(z[l, p].X),
                            "Miktar (birim)" : v,
                            "Cy_birim ($/b)" : round(c_unit, 2),
                            "Tasima_Mal ($)" : round(c_unit * v, 1),
                        })
        pd.DataFrame(y_rows).to_excel(writer, sheet_name="Y_DC_Mus", index=False)

        # ── SEKME 5: B — Karsilanamayan talep (backlog) ───────────────────
        b_rows = []
        for k in customers:
            for n in products:
                for p in periods_ext:
                    v       = round(B[k, n, p].X)
                    t_cost  = T.get(p, 0) * v   # p=0 icin T yok
                    b_rows.append({
                        "Degisken"        : f"B{k}{n}{p}",
                        "Musteri"         : _kn[k],
                        "Urun"            : _prn[n],
                        "Donem"           : p,
                        "Talep (birim)"   : round(demand.get((k, n, p), 0)) if p >= 1 else 0,
                        "Backlog (birim)" : v,
                        "T_p ($/b)"       : T.get(p, 0),
                        "Backlog_Mal ($)" : round(t_cost, 1),
                    })
        pd.DataFrame(b_rows).to_excel(writer, sheet_name="B_Backlog", index=False)

        # ── SEKME 6: Q — DC stok seviyeleri ──────────────────────────────
        q_rows = []
        for l in dcs:
            for n in products:
                for p in periods_ext:
                    v      = round(Q[l, n, p].X)
                    r_cost = R.get(p, 0) * v
                    q_rows.append({
                        "Degisken"        : f"Q{l}{n}{p}",
                        "DC"              : _dn[l],
                        "Urun"            : _prn[n],
                        "Donem"           : p,
                        "DC_Acik"         : round(z[l, p].X) if p >= 1 else 1,
                        "Stok (birim)"    : v,
                        "R_p ($/b)"       : R.get(p, 0),
                        "Tutma_Mal ($)"   : round(r_cost, 1),
                    })
        pd.DataFrame(q_rows).to_excel(writer, sheet_name="Q_Stok", index=False)

        # ── SEKME 7: z — DC acma/kapama kararlari ─────────────────────────
        z_rows = []
        for l in dcs:
            for p in periods:
                z_val_  = round(z[l, p].X)
                d_val_  = round(delta[l, p].X)
                fixed_c = f_dc[l, p] * z_val_
                sw_c    = sc_dc[l] * d_val_
                z_rows.append({
                    "Degisken"          : f"z{l}{p}",
                    "DC"                : _dn[l],
                    "Donem"             : p,
                    "Acik_mi (1/0)"     : z_val_,
                    "Gecis_mi (delta)"  : d_val_,
                    "f_lp (sabit $)"    : f_dc[l, p],
                    "sc_l (gecis $)"    : sc_dc[l],
                    "Sabit_Mal ($)"     : round(fixed_c, 1),
                    "Gecis_Mal ($)"     : round(sw_c, 1),
                    "Toplam_DC_Mal ($)" : round(fixed_c + sw_c, 1),
                })
        pd.DataFrame(z_rows).to_excel(writer, sheet_name="z_DC_Karar", index=False)

        # ── SEKME 8: MALIYET_OZET — Tum maliyet turlerini donem bazinda ───
        cost_rows = []
        for p in periods:
            # Taşıma (X)
            c_trans_p = sum(
                C[i,j,t,n,p]*X[i,j,t,n,p].X
                for i in suppliers for j in factories for t in modes for n in products
            )
            c_opp_p = sum(
                N[i,j,t,n,p]*pi_rate*X[i,j,t,n,p].X
                for i in suppliers for j in factories for t in modes for n in products
            )
            c_w_p = sum(Cw[j,l,n,p]*W[j,l,n,p].X
                        for j in factories for l in dcs for n in products)
            c_y_p = sum(Cy[l,k,n,p]*Y[l,k,n,p].X
                        for l in dcs for k in customers for n in products)
            c_hold_p  = sum(R[p]*Q[l,n,p].X for l in dcs for n in products)
            c_stock_p = sum(T[p]*B[k,n,p].X for k in customers for n in products)
            c_fixed_p = sum(f_dc[l,p]*z[l,p].X for l in dcs)
            c_sw_p    = sum(sc_dc[l]*delta[l,p].X for l in dcs)
            total_p   = c_trans_p+c_opp_p+c_w_p+c_y_p+c_hold_p+c_stock_p+c_fixed_p+c_sw_p
            cost_rows.append({
                "Donem"            : p,
                "Tasima_X ($)"     : round(c_trans_p, 1),
                "Firsat_N ($)"     : round(c_opp_p, 1),
                "FabDC_W ($)"      : round(c_w_p, 1),
                "DCMus_Y ($)"      : round(c_y_p, 1),
                "Tutma_R ($)"      : round(c_hold_p, 1),
                "YokSatma_T ($)"   : round(c_stock_p, 1),
                "DC_Sabit ($)"     : round(c_fixed_p, 1),
                "DC_Gecis ($)"     : round(c_sw_p, 1),
                "Donem_Toplam ($)" : round(total_p, 1),
            })
        # Toplam satiri
        df_cost = pd.DataFrame(cost_rows)
        total_row = {"Donem": "TOPLAM"}
        for col in df_cost.columns[1:]:
            total_row[col] = round(df_cost[col].sum(), 1)
        df_cost = pd.concat([df_cost, pd.DataFrame([total_row])], ignore_index=True)
        df_cost.to_excel(writer, sheet_name="MALIYET_OZET", index=False)

        # ── SEKME 9: URUN_BAZLI — Urun bazinda toplam maliyet ozeti ───────
        urun_rows = []
        for n in products:
            c_x  = sum((C[i,j,t,n,p]+N[i,j,t,n,p]*pi_rate)*X[i,j,t,n,p].X
                       for i in suppliers for j in factories for t in modes for p in periods)
            c_w  = sum(Cw[j,l,n,p]*W[j,l,n,p].X
                       for j in factories for l in dcs for p in periods)
            c_y  = sum(Cy[l,k,n,p]*Y[l,k,n,p].X
                       for l in dcs for k in customers for p in periods)
            c_h  = sum(R[p]*Q[l,n,p].X for l in dcs for p in periods)
            c_b  = sum(T[p]*B[k,n,p].X for k in customers for p in periods)
            tot_x = sum(X[i,j,t,n,p].X for i in suppliers for j in factories for t in modes for p in periods)
            tot_y = sum(Y[l,k,n,p].X for l in dcs for k in customers for p in periods)
            urun_rows.append({
                "Urun"                 : _prn[n],
                "Toplam_Tasinan(X)"    : round(tot_x),
                "Toplam_Dagitilan(Y)"  : round(tot_y),
                "Trans+Firsat ($)"     : round(c_x, 1),
                "FabDC ($)"            : round(c_w, 1),
                "DCMus ($)"            : round(c_y, 1),
                "Tutma ($)"            : round(c_h, 1),
                "YokSatma ($)"         : round(c_b, 1),
                "Urun_Toplam ($)"      : round(c_x+c_w+c_y+c_h+c_b, 1),
            })
        pd.DataFrame(urun_rows).to_excel(writer, sheet_name="URUN_BAZLI", index=False)

    print(f"\n  Excel workbook olusturuldu -> {out_path}")
    print(f"  Sekmeler: OZET | X_Ted_Fab | W_Fab_DC | Y_DC_Mus | B_Backlog | Q_Stok | z_DC_Karar | MALIYET_OZET | URUN_BAZLI")

elif model.status == GRB.INFEASIBLE:
    print("Model UYGUN DEGIL. IIS hesaplaniyor...")
    model.computeIIS()
    iis_path = os.path.join(RESULT_DIR, "milp.ilp")
    model.write(iis_path)
    print(f"  IIS -> {iis_path}")

else:
    print(f"Cozucu durumu: {model.status}")
