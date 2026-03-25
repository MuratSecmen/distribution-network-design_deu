import os
import math
import itertools
import gurobipy as gp
from gurobipy import GRB
import pandas as pd
from datetime import datetime

# =========================================================================
# 0.  CONFIGURATION
# =========================================================================
PARAMS_FILE    = os.path.join("data", "parameters_v2.xlsx")
TRANSPORT_FILE = os.path.join("data", "transportation_costs_v2.xlsx")
OUTPUT_DIR     = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================================================================
# 1.  SET LABELS  — genişletilmiş kardinaliteler
# =========================================================================
SUPPLIER_NAMES = {
    1:"Moskova", 2:"Berlin",   3:"Oslo",    4:"Astana",  5:"Pekin",
    6:"Mumbai",  7:"Dubai",    8:"Şanghay", 9:"Toronto", 10:"Kahire"
}
FACTORY_NAMES = {
    1:"Madrid",   2:"St.Pete", 3:"Varşova", 4:"Ankara",
    5:"Milano",   6:"Lyon",    7:"Hamburg", 8:"Budapeşte"
}
DC_NAMES = {
    1:"Ukrayna", 2:"Polonya", 3:"Romanya",
    4:"Çekya",   5:"Macaristan", 6:"Slovakya"
}
CUSTOMER_NAMES = {k: f"Müşteri_{k}" for k in range(1, 21)}
MODE_NAMES = {1:"Demiryolu", 2:"Karayolu", 3:"Denizyolu", 4:"Havayolu"}

SUPPLIERS = list(SUPPLIER_NAMES.keys())   # I = {1..10}
FACTORIES = list(FACTORY_NAMES.keys())    # J = {1..8}
DCS       = list(DC_NAMES.keys())         # L = {1..6}
CUSTOMERS = list(CUSTOMER_NAMES.keys())   # K = {1..20}
MODES     = list(MODE_NAMES.keys())       # M = {1..4}
PRODUCTS  = list(range(1, 6))             # N = {1..5}
PERIODS   = list(range(1, 13))            # P = {1..12}

# =========================================================================
# 2.  COĞRAFI KOORDİNATLAR (enlem, boylam)
# =========================================================================
COORDS = {
    # Tedarikçiler
    "sup_1":  (55.75,  37.62),
    "sup_2":  (52.52,  13.40),
    "sup_3":  (59.91,  10.75),
    "sup_4":  (51.18,  71.45),
    "sup_5":  (39.90, 116.40),
    "sup_6":  (19.08,  72.88),
    "sup_7":  (25.20,  55.27),
    "sup_8":  (31.23, 121.47),
    "sup_9":  (43.65, -79.38),
    "sup_10": (30.06,  31.25),
    # Fabrikalar
    "fac_1":  (40.42,  -3.70),
    "fac_2":  (59.95,  30.32),
    "fac_3":  (52.23,  21.01),
    "fac_4":  (39.93,  32.85),
    "fac_5":  (45.46,   9.19),
    "fac_6":  (45.75,   4.85),
    "fac_7":  (53.55,   9.99),
    "fac_8":  (47.50,  19.04),
    # DC'ler
    "dc_1":   (50.45,  30.52),
    "dc_2":   (52.23,  21.01),
    "dc_3":   (44.43,  26.10),
    "dc_4":   (50.08,  14.44),
    "dc_5":   (47.50,  19.04),
    "dc_6":   (48.15,  17.11),
}

def haversine(lat1, lon1, lat2, lon2):
    
    R = 6371.0
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi  = math.radians(lat2 - lat1)
    dlam  = math.radians(lon2 - lon1)
    a = math.sin(dphi/2)**2 + math.cos(phi1)*math.cos(phi2)*math.sin(dlam/2)**2
    return R * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))

def compute_distances():
    """
    Tedarikçi–fabrika mesafeleri (km).
    Nakliye maliyeti = dist[i,j] × rho[m] × (1 + 0.05*(n-1))
    Emisyon        = dist[i,j] × epsilon[m]
    """
    dist = {}
    for i in SUPPLIERS:
        c_i = COORDS[f"sup_{i}"]
        for j in FACTORIES:
            c_j = COORDS[f"fac_{j}"]
            dist[(i, j)] = round(haversine(*c_i, *c_j), 1)
    return dist

# =========================================================================
# 3.  COĞRAFI TAŞIMA MODU FEASİBİLİTESİ
# =========================================================================

FEASIBLE_MODES = {}
SEA_SUPPLIERS  = {3, 6, 7}
SEA_FACTORIES  = {1, 2, 5, 6, 7}
DIST_CACHE     = compute_distances()

for i in SUPPLIERS:
    for j in FACTORIES:
        modes = [1, 4]
        # Karayolu: mesafe ≤ 5000 km ise ekle
        if DIST_CACHE[(i, j)] <= 5000:
            modes.append(2)
        # Denizyolu: her iki taraf da deniz erişimliyse ekle
        if i in SEA_SUPPLIERS and j in SEA_FACTORIES:
            modes.append(3)
        FEASIBLE_MODES[(i, j)] = sorted(set(modes))

# =========================================================================
# 4.  VERİ YÜKLEME
# =========================================================================
def load_parameters_v2():
    print("  [INFO] Excel parametreleri okunuyor ...")

    df_s = pd.read_excel(PARAMS_FILE, sheet_name="SCALAR_PARAMS",
                         index_col=0, header=0)
    scalar   = df_s.iloc[:, 0].to_dict()
    big_M    = float(scalar.get("big_M",       1e7))
    eta      = float(scalar.get("eta",         0.95))
    E_cap    = float(scalar.get("E_cap",       1e9))

    def load_long(sheet, keys, val_col):
        df = pd.read_excel(PARAMS_FILE, sheet_name=sheet,
                           index_col=False, header=0)
        out = {}
        for _, row in df.iterrows():
            k = tuple(int(row[c]) for c in keys)
            out[k] = float(row[val_col])
        return out

    def load_long_tc(sheet, keys, val_col):
        df = pd.read_excel(TRANSPORT_FILE, sheet_name=sheet,
                           index_col=False, header=0)
        out = {}
        for _, row in df.iterrows():
            k = tuple(int(row[c]) for c in keys)
            out[k] = float(row[val_col])
        return out

    # ── Parametreler ───────────────────────────────────────────────────
    demand        = load_long("DEMAND",        ["k","n","p"], "demand")
    supplier_cap  = load_long("SUPPLIER_CAP",  ["i","n","p"], "capacity")
    factory_cap   = load_long("FACTORY_CAP",   ["j","n","p"], "capacity")
    mode_cap      = load_long("MODE_CAP",       ["m","p"],    "capacity")
    dc_throughput = load_long("DC_THROUGHPUT",  ["l","p"],    "throughput")
    dc_storage    = load_long("DC_STORAGE",     ["l","p"],    "storage")
    dc_invest     = load_long("DC_INVEST",      ["l","p"],    "cost")
    dc_opcost     = load_long("DC_OPCOST",      ["l","p"],    "cost")
    hold_cost     = load_long("HOLD_COST",      ["l","n","p"],"cost")
    back_cost     = load_long("BACK_COST",      ["k","n","p"],"cost")
    factory_dc_cost = load_long("FACTORY_DC_COST",["j","l","n","p"],"cost")
    dc_cust_cost  = load_long("DC_CUST_COST",   ["l","k","n","p"],"cost")
    # Fabrika açma/kapama maliyetleri
    fac_invest    = load_long("FAC_INVEST",     ["j","p"],    "cost")
    fac_switch    = load_long("FAC_SWITCH",     ["j"],        "sc")

    # ── Nakliye maliyeti: Haversine × rho[m] × ürün faktörü ──────────
    df_rho = pd.read_excel(PARAMS_FILE, sheet_name="MODE_FACTORS",
                           index_col=False, header=0)
    rho     = {int(r["m"]): float(r["rho"])     for _, r in df_rho.iterrows()}
    epsilon = {int(r["m"]): float(r["epsilon"]) for _, r in df_rho.iterrows()}

    trans_cost = {}
    emission_factor = {}
    for i in SUPPLIERS:
        for j in FACTORIES:
            for m in FEASIBLE_MODES.get((i, j), []):
                for n in PRODUCTS:
                    d_km = DIST_CACHE[(i, j)]
                    trans_cost[(i,j,m,n)]     = round(d_km * rho[m] * (1 + 0.05*(n-1)), 4)
                    emission_factor[(i,j,m,n)] = round(d_km * epsilon[m], 4)

    print(f"  [INFO] big_M={big_M:.0f} | eta={eta} | E_cap={E_cap:.0f}")
    print(f"  [INFO] demand={len(demand)} | trans_cost={len(trans_cost)}")

    return {
        "big_M":          big_M,
        "eta":            eta,
        "E_cap":          E_cap,
        "demand":         demand,
        "trans_cost":     trans_cost,
        "emission_factor":emission_factor,
        "factory_dc_cost":factory_dc_cost,
        "dc_cust_cost":   dc_cust_cost,
        "supplier_cap":   supplier_cap,
        "factory_cap":    factory_cap,
        "mode_cap":       mode_cap,
        "dc_throughput":  dc_throughput,
        "dc_storage":     dc_storage,
        "dc_invest":      dc_invest,
        "dc_opcost":      dc_opcost,
        "dc_switch":      pd.read_excel(PARAMS_FILE, sheet_name="DC_SWITCH",
                                        index_col=False).set_index("l")["sc"].to_dict(),
        "hold_cost":      hold_cost,
        "back_cost":      back_cost,
        "fac_invest":     fac_invest,
        "fac_switch":     fac_switch,
    }

# =========================================================================
# 5.  MODEL KURULUMU
# =========================================================================
def build_model(params):
    M_big   = params["big_M"]
    eta     = params["eta"]
    E_cap   = params["E_cap"]
    D       = params["demand"]
    c_S     = params["trans_cost"]
    em      = params["emission_factor"]
    c_F     = params["factory_dc_cost"]
    c_D     = params["dc_cust_cost"]
    alpha   = params["supplier_cap"]
    b_cap   = params["factory_cap"]
    A       = params["mode_cap"]
    u_cap   = params["dc_throughput"]
    v_cap   = params["dc_storage"]
    f_inv   = params["dc_invest"]
    g_op    = params["dc_opcost"]
    sc_dc   = params["dc_switch"]
    h       = params["hold_cost"]
    beta    = params["back_cost"]
    fo      = params["fac_invest"]
    sc_fac  = params["fac_switch"]

    model = gp.Model("DistNet_MILP_Extended_v2")
    model.Params.LogFile   = os.path.join(OUTPUT_DIR, "gurobi.log")
    model.Params.TimeLimit = 3600
    model.Params.MIPGap    = 0.01

    # ─────────────────────────────────────────────────────────────────
    # 5a. KARAR DEĞİŞKENLERİ
    # ─────────────────────────────────────────────────────────────────

    # x[i,j,m,n,p] : tedarikçi i → fabrika j, mod m, ürün n, dönem p
    x = {
        (i,j,m,n,p): model.addVar(lb=0.0, name=f"x_{i}_{j}_{m}_{n}_{p}")
        for i in SUPPLIERS for j in FACTORIES
        for m in MODES for n in PRODUCTS for p in PERIODS
        if m in FEASIBLE_MODES.get((i,j), [])
    }

    # w[j,l,n,p] : fabrika j → DC l, ürün n, dönem p
    w = {
        (j,l,n,p): model.addVar(lb=0.0, name=f"w_{j}_{l}_{n}_{p}")
        for j in FACTORIES for l in DCS
        for n in PRODUCTS for p in PERIODS
    }

    # y[l,k,n,p] : DC l → müşteri k, ürün n, dönem p
    y = {
        (l,k,n,p): model.addVar(lb=0.0, name=f"y_{l}_{k}_{n}_{p}")
        for l in DCS for k in CUSTOMERS
        for n in PRODUCTS for p in PERIODS
    }

    # q[l,n,p] : DC l stok seviyesi, ürün n, dönem p sonu (p=0 başlangıç)
    q = {
        (l,n,p): model.addVar(lb=0.0, name=f"q_{l}_{n}_{p}")
        for l in DCS for n in PRODUCTS for p in [0]+PERIODS
    }

    # b_bl[k,n,p] : müşteri k birikmiş talep, ürün n, dönem p sonu (p=0 başlangıç)
    b_bl = {
        (k,n,p): model.addVar(lb=0.0, name=f"b_{k}_{n}_{p}")
        for k in CUSTOMERS for n in PRODUCTS for p in [0]+PERIODS
    }

    # z[l,p] : DC l dönem p'de açık (binary)
    z = {
        (l,p): model.addVar(vtype=GRB.BINARY, name=f"z_{l}_{p}")
        for l in DCS for p in PERIODS
    }

    # delta[l,p] : DC l dönem p'de statü değiştirdi (binary)
    delta = {
        (l,p): model.addVar(vtype=GRB.BINARY, name=f"delta_{l}_{p}")
        for l in DCS for p in PERIODS
    }

    # phi[j,p] : fabrika j dönem p'de açık (binary)  — YENİ
    phi = {
        (j,p): model.addVar(vtype=GRB.BINARY, name=f"phi_{j}_{p}")
        for j in FACTORIES for p in PERIODS
    }

    # gamma[j,p] : fabrika j dönem p'de statü değiştirdi (binary)  — YENİ
    gamma = {
        (j,p): model.addVar(vtype=GRB.BINARY, name=f"gamma_{j}_{p}")
        for j in FACTORIES for p in PERIODS
    }

    model.update()

    # ─────────────────────────────────────────────────────────────────
    # 5b. AMAÇ FONKSİYONU  — LaTeX eq_1 ile birebir
    #
    # min Z = Σ c_S*x + Σ c_F*w + Σ c_D*y
    #       + Σ (f*z + g*Σy) + Σ h*q + Σ beta*b + Σ sc_dc*delta
    #       + Σ fo*phi + Σ sc_fac*gamma
    # ─────────────────────────────────────────────────────────────────
    obj = gp.LinExpr()

    for (i,j,m,n,p), var in x.items():
        obj += c_S.get((i,j,m,n), 0.0) * var

    for (j,l,n,p), var in w.items():
        obj += c_F.get((j,l,n,p), 0.0) * var

    for (l,k,n,p), var in y.items():
        obj += c_D.get((l,k,n,p), 0.0) * var

    for l in DCS:
        for p in PERIODS:
            obj += f_inv.get((l,p), 0.0) * z[l,p]
            obj += g_op.get((l,p), 0.0) * gp.quicksum(
                y[l,k,n,p] for k in CUSTOMERS for n in PRODUCTS)

    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                obj += h.get((l,n,p), 0.0) * q[l,n,p]

    for k in CUSTOMERS:
        for n in PRODUCTS:
            for p in PERIODS:
                obj += beta.get((k,n,p), 0.0) * b_bl[k,n,p]

    for l in DCS:
        for p in PERIODS:
            obj += sc_dc.get(l, 0.0) * delta[l,p]

    # Fabrika yatırım ve switching maliyetleri  — YENİ
    for j in FACTORIES:
        for p in PERIODS:
            obj += fo.get((j,p), 0.0)    * phi[j,p]
            obj += sc_fac.get((j,), 0.0) * gamma[j,p]

    model.setObjective(obj, GRB.MINIMIZE)

    # ─────────────────────────────────────────────────────────────────
    # 5c. KISITLAR
    # ─────────────────────────────────────────────────────────────────

    # ── eq_2 : Tedarikçi kapasite kısıtı ─────────────────────────────
    # Σ_{j,m} x[i,j,m,n,p] <= alpha[i,n,p]
    for i in SUPPLIERS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(x[i,j,m,n,p]
                                for j in FACTORIES for m in MODES
                                if (i,j,m,n,p) in x)
                    <= alpha.get((i,n,p), GRB.INFINITY),
                    name=f"eq2_{i}_{n}_{p}"
                )

    # ── eq_3 : Fabrika üretim kapasitesi ─────────────────────────────
    # Σ_l w[j,l,n,p] <= b[j,n,p]
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(w[j,l,n,p] for l in DCS)
                    <= b_cap.get((j,n,p), GRB.INFINITY),
                    name=f"eq3_{j}_{n}_{p}"
                )

    # ── eq_4 : Fabrikada akış dengesi ────────────────────────────────
    # Σ_{i,m} x[i,j,m,n,p] = Σ_l w[j,l,n,p]
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(x[i,j,m,n,p]
                                for i in SUPPLIERS for m in MODES
                                if (i,j,m,n,p) in x)
                    == gp.quicksum(w[j,l,n,p] for l in DCS),
                    name=f"eq4_{j}_{n}_{p}"
                )

    # ── eq_5 : DC envanter dengesi ────────────────────────────────────
    # q[l,n,p] = q[l,n,p-1] + Σ_j w[j,l,n,p] - Σ_k y[l,k,n,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    q[l,n,p]
                    == q[l,n,p-1]
                    + gp.quicksum(w[j,l,n,p] for j in FACTORIES)
                    - gp.quicksum(y[l,k,n,p] for k in CUSTOMERS),
                    name=f"eq5_{l}_{n}_{p}"
                )

    # ── eq_6 : DC depo kapasitesi ─────────────────────────────────────
    # Σ_n q[l,n,p] <= v[l,p]
    for l in DCS:
        for p in PERIODS:
            model.addConstr(
                gp.quicksum(q[l,n,p] for n in PRODUCTS)
                <= v_cap.get((l,p), GRB.INFINITY),
                name=f"eq6_{l}_{p}"
            )

    # ── eq_7 : DC throughput kapasitesi ──────────────────────────────
    # Σ_{k,n} y[l,k,n,p] <= u[l,p]
    for l in DCS:
        for p in PERIODS:
            model.addConstr(
                gp.quicksum(y[l,k,n,p] for k in CUSTOMERS for n in PRODUCTS)
                <= u_cap.get((l,p), GRB.INFINITY),
                name=f"eq7_{l}_{p}"
            )

    # ── eq_8 : Talep karşılama ────────────────────────────────────────
    # Σ_l y[l,k,n,p] + b[k,n,p-1] - b[k,n,p] = d[k,n,p]
    for k in CUSTOMERS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(y[l,k,n,p] for l in DCS)
                    + b_bl[k,n,p-1] - b_bl[k,n,p]
                    == D.get((k,n,p), 0.0),
                    name=f"eq8_{k}_{n}_{p}"
                )

    # ── eq_9 : Taşıma modu kapasitesi ────────────────────────────────
    # Σ_{i,j} x[i,j,m,n,p] <= A[m,p]
    for m in MODES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(x[i,j,m,n,p]
                                for i in SUPPLIERS for j in FACTORIES
                                if (i,j,m,n,p) in x)
                    <= A.get((m,p), GRB.INFINITY),
                    name=f"eq9_{m}_{n}_{p}"
                )

    # ── eq_10 : DC girişi → açık DC bağı (Big-M) ─────────────────────
    # Σ_j w[j,l,n,p] <= M_big * z[l,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(w[j,l,n,p] for j in FACTORIES)
                    <= M_big * z[l,p],
                    name=f"eq10_{l}_{n}_{p}"
                )

    # ── eq_11 : DC çıkışı → açık DC bağı (Big-M) ─────────────────────
    # Σ_k y[l,k,n,p] <= M_big * z[l,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(y[l,k,n,p] for k in CUSTOMERS)
                    <= M_big * z[l,p],
                    name=f"eq11_{l}_{n}_{p}"
                )

    # ── eq_12 & eq_13 : DC switching linearizasyonu ───────────────────
    # delta[l,p] >= z[l,p] - z[l,p-1]   (kapanmış→açılmış)
    # delta[l,p] >= z[l,p-1] - z[l,p]   (açılmış→kapanmış)
    for l in DCS:
        for p in PERIODS:
            z_prev = z[l,p-1] if p > 1 else gp.LinExpr(0)
            model.addConstr(delta[l,p] >= z[l,p]  - z_prev, name=f"eq12_{l}_{p}")
            model.addConstr(delta[l,p] >= z_prev  - z[l,p], name=f"eq13_{l}_{p}")

    # ── eq_16 : Fabrika girişi → açık fabrika bağı (Big-M)  — YENİ ───
    # Σ_{i,m} x[i,j,m,n,p] <= M_big * phi[j,p]
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(x[i,j,m,n,p]
                                for i in SUPPLIERS for m in MODES
                                if (i,j,m,n,p) in x)
                    <= M_big * phi[j,p],
                    name=f"eq16_{j}_{n}_{p}"
                )

    # ── eq_17 : Fabrika çıkışı → açık fabrika bağı (Big-M)  — YENİ ──
    # Σ_l w[j,l,n,p] <= M_big * phi[j,p]
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(w[j,l,n,p] for l in DCS)
                    <= M_big * phi[j,p],
                    name=f"eq17_{j}_{n}_{p}"
                )

    # ── eq_18 & eq_19 : Fabrika switching linearizasyonu  — YENİ ──────
    # gamma[j,p] >= phi[j,p] - phi[j,p-1]
    # gamma[j,p] >= phi[j,p-1] - phi[j,p]
    for j in FACTORIES:
        for p in PERIODS:
            phi_prev = phi[j,p-1] if p > 1 else gp.LinExpr(0)
            model.addConstr(gamma[j,p] >= phi[j,p]  - phi_prev, name=f"eq18_{j}_{p}")
            model.addConstr(gamma[j,p] >= phi_prev  - phi[j,p], name=f"eq19_{j}_{p}")

    # ── eq_20 : Karbon emisyon bütçesi kısıtı  — YENİ ─────────────────
    # Σ_{i,j,m,n,p} e[i,j,m,n] * x[i,j,m,n,p] <= E_cap
    model.addConstr(
        gp.quicksum(
            em.get((i,j,m,n), 0.0) * x[i,j,m,n,p]
            for (i,j,m,n,p) in x
        ) <= E_cap,
        name="eq20_carbon_budget"
    )

    # ── eq_21 : Minimum hizmet düzeyi kısıtı  — YENİ ──────────────────
    # Σ_l y[l,k,n,p] >= eta * d[k,n,p]
    for k in CUSTOMERS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(y[l,k,n,p] for l in DCS)
                    >= eta * D.get((k,n,p), 0.0),
                    name=f"eq21_{k}_{n}_{p}"
                )

    # ── eq_14 & eq_15 : Non-negativity & binary
    #    lb=0 ve vtype=BINARY değişken tanımında uygulandı.

    return model, x, w, y, q, b_bl, z, delta, phi, gamma


# =========================================================================
# 6.  SONUÇ ÇIKTILARI
# =========================================================================
def export_results(model, x, w, y, q, b_bl, z, delta, phi, gamma, params):
    status = model.Status
    if status not in (GRB.OPTIMAL, GRB.TIME_LIMIT, GRB.SUBOPTIMAL):
        print(f"  [WARN] Model status = {status}. Çözüm bulunamadı.")
        return

    print(f"\n  ── Amaç değeri  : {model.ObjVal:,.2f}")
    print(f"  ── MIP gap      : {model.MIPGap * 100:.4f} %")
    print(f"  ── Çözüm süresi : {model.Runtime:.2f} s\n")

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(OUTPUT_DIR, f"results_{ts}.xlsx")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:

        # Sheet 1: x – Tedarikçi→Fabrika akışları
        rows = []
        for (i,j,m,n,p), var in x.items():
            if var.X > 1e-6:
                rows.append({
                    "Supplier":SUPPLIER_NAMES[i], "Factory":FACTORY_NAMES[j],
                    "Mode":MODE_NAMES[m], "Product":n, "Period":p,
                    "Quantity":round(var.X,4),
                    "Distance_km":DIST_CACHE[(i,j)],
                    "Emission_kg":round(params["emission_factor"].get((i,j,m,n),0)*var.X,2),
                })
        pd.DataFrame(rows).to_excel(writer, sheet_name="x_flows", index=False)

        # Sheet 2: w – Fabrika→DC akışları
        rows = [{"Factory":FACTORY_NAMES[j],"DC":DC_NAMES[l],
                 "Product":n,"Period":p,"Quantity":round(w[j,l,n,p].X,4)}
                for j in FACTORIES for l in DCS for n in PRODUCTS for p in PERIODS
                if w[j,l,n,p].X > 1e-6]
        pd.DataFrame(rows).to_excel(writer, sheet_name="w_factory_dc", index=False)

        # Sheet 3: y – DC→Müşteri teslimatları
        rows = [{"DC":DC_NAMES[l],"Customer":CUSTOMER_NAMES[k],
                 "Product":n,"Period":p,"Delivery":round(y[l,k,n,p].X,4)}
                for l in DCS for k in CUSTOMERS for n in PRODUCTS for p in PERIODS
                if y[l,k,n,p].X > 1e-6]
        pd.DataFrame(rows).to_excel(writer, sheet_name="y_delivery", index=False)

        # Sheet 4: q – Stok seviyeleri
        rows = [{"DC":DC_NAMES[l],"Product":n,"Period":p,
                 "Inventory":round(q[l,n,p].X,4)}
                for l in DCS for n in PRODUCTS for p in [0]+PERIODS]
        pd.DataFrame(rows).to_excel(writer, sheet_name="q_inventory", index=False)

        # Sheet 5: b – Backlog seviyeleri
        rows = [{"Customer":CUSTOMER_NAMES[k],"Product":n,"Period":p,
                 "Backlog":round(b_bl[k,n,p].X,4)}
                for k in CUSTOMERS for n in PRODUCTS for p in [0]+PERIODS
                if b_bl[k,n,p].X > 1e-6]
        pd.DataFrame(rows).to_excel(writer, sheet_name="b_backlog", index=False)

        # Sheet 6: z – DC açık/kapalı
        rows = [{"DC":DC_NAMES[l],"Period":p,"Open":int(z[l,p].X+0.5)}
                for l in DCS for p in PERIODS]
        pd.DataFrame(rows).to_excel(writer, sheet_name="z_DC_status", index=False)

        # Sheet 7: phi – Fabrika açık/kapalı  — YENİ
        rows = [{"Factory":FACTORY_NAMES[j],"Period":p,"Open":int(phi[j,p].X+0.5)}
                for j in FACTORIES for p in PERIODS]
        pd.DataFrame(rows).to_excel(writer, sheet_name="phi_factory_status", index=False)

        # Sheet 8: delta – DC switching
        rows = [{"DC":DC_NAMES[l],"Period":p,"Switch":1}
                for l in DCS for p in PERIODS if delta[l,p].X > 0.5]
        pd.DataFrame(rows).to_excel(writer, sheet_name="delta_DC_switch", index=False)

        # Sheet 9: gamma – Fabrika switching  — YENİ
        rows = [{"Factory":FACTORY_NAMES[j],"Period":p,"Switch":1}
                for j in FACTORIES for p in PERIODS if gamma[j,p].X > 0.5]
        pd.DataFrame(rows).to_excel(writer, sheet_name="gamma_fac_switch", index=False)

        # Sheet 10: Emisyon özeti  — YENİ
        total_em = sum(
            params["emission_factor"].get((i,j,m,n),0) * x[i,j,m,n,p].X
            for (i,j,m,n,p) in x
        )
        rows = []
        for m in MODES:
            em_m = sum(
                params["emission_factor"].get((i,j,m,n),0) * x[i,j,m,n,p].X
                for (i,j,_m,n,p) in x if _m == m
            )
            rows.append({"Mode":MODE_NAMES[m], "Emission_kg":round(em_m,2)})
        rows.append({"Mode":"TOPLAM","Emission_kg":round(total_em,2)})
        rows.append({"Mode":"Bütçe (E_cap)","Emission_kg":params["E_cap"]})
        rows.append({"Mode":"Kullanım %","Emission_kg":round(100*total_em/params["E_cap"],2)})
        pd.DataFrame(rows).to_excel(writer, sheet_name="emission_summary", index=False)

        # Sheet 11: Dönem bazlı maliyet dökümü
        rows = []
        for p in PERIODS:
            trans_p   = sum(params["trans_cost"].get((i,j,m,n),0)*x[i,j,m,n,p].X
                            for (i,j,m,n,_p) in x if _p==p)
            fdc_p     = sum(params["factory_dc_cost"].get((j,l,n,p),0)*w[j,l,n,p].X
                            for j in FACTORIES for l in DCS for n in PRODUCTS)
            dcc_p     = sum(params["dc_cust_cost"].get((l,k,n,p),0)*y[l,k,n,p].X
                            for l in DCS for k in CUSTOMERS for n in PRODUCTS)
            hold_p    = sum(params["hold_cost"].get((l,n,p),0)*q[l,n,p].X
                            for l in DCS for n in PRODUCTS)
            back_p    = sum(params["back_cost"].get((k,n,p),0)*b_bl[k,n,p].X
                            for k in CUSTOMERS for n in PRODUCTS)
            inv_dc_p  = sum(params["dc_invest"].get((l,p),0)*z[l,p].X  for l in DCS)
            sw_dc_p   = sum(params["dc_switch"].get(l,0)*delta[l,p].X  for l in DCS)
            inv_fac_p = sum(params["fac_invest"].get((j,p),0)*phi[j,p].X for j in FACTORIES)
            sw_fac_p  = sum(params["fac_switch"].get((j,),0)*gamma[j,p].X for j in FACTORIES)
            rows.append({
                "Period":p, "Transport":round(trans_p,2),
                "Factory_DC":round(fdc_p,2), "DC_Cust":round(dcc_p,2),
                "Holding":round(hold_p,2),   "Backlog":round(back_p,2),
                "DC_Invest":round(inv_dc_p,2),"DC_Switch":round(sw_dc_p,2),
                "Fac_Invest":round(inv_fac_p,2),"Fac_Switch":round(sw_fac_p,2),
                "Total":round(trans_p+fdc_p+dcc_p+hold_p+back_p+
                              inv_dc_p+sw_dc_p+inv_fac_p+sw_fac_p,2),
            })
        rows.append({"Period":"GRAND TOTAL","Total":round(model.ObjVal,2)})
        pd.DataFrame(rows).to_excel(writer, sheet_name="CostByPeriod", index=False)

        # Sheet 12: Ürün özet tablosu
        rows = []
        for n in PRODUCTS:
            for p in PERIODS:
                rows.append({
                    "Product":n, "Period":p,
                    "Production":round(sum(w[j,l,n,p].X for j in FACTORIES for l in DCS),4),
                    "Delivery":  round(sum(y[l,k,n,p].X for l in DCS for k in CUSTOMERS),4),
                    "Inventory": round(sum(q[l,n,p].X   for l in DCS),4),
                    "Backlog":   round(sum(b_bl[k,n,p].X for k in CUSTOMERS),4),
                    "Demand":    round(sum(params["demand"].get((k,n,p),0) for k in CUSTOMERS),4),
                })
        pd.DataFrame(rows).to_excel(writer, sheet_name="ProductSummary", index=False)

    print(f"  [INFO] Sonuçlar → {out}")


# =========================================================================
# 7.  MAIN
# =========================================================================
def main():
    print("=" * 65)
    print("  Distribution Network Design  –  Extended MILP v2")
    print("  Seçmen, Öncan & Tuna (2015) | DEÜ Lojistik")
    print("=" * 65)

    params = load_parameters()

    print("\n  [INFO] Model kuruluyor ...")
    model, x, w, y, q, b_bl, z, delta, phi, gamma = build_model(params)

    print(f"  [INFO] Değişken sayısı  : {model.NumVars:,}")
    print(f"  [INFO] Kısıt sayısı     : {model.NumConstrs:,}")
    print(f"  [INFO] Binary değişken  : {model.NumBinVars:,}")
    print(f"  [INFO] Feasible x triples: {len(x):,}")
    print("\n  [INFO] Çözüm başlıyor ...\n")

    model.optimize()

    export_results(model, x, w, y, q, b_bl, z, delta, phi, gamma, params)

    print("\n  Bitti.\n")


if __name__ == "__main__":
    main()
