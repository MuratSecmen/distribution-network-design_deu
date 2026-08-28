import sys
import os
import gurobipy as gp
from gurobipy import GRB
import pandas as pd
import numpy as np
from datetime import datetime

# =========================================================================
# 0.  CONFIGURATION
# =========================================================================
PARAMS_FILE    = os.path.join("data", "parameters.xlsx")
TRANSPORT_FILE = os.path.join("data", "transportation_costs.xlsx")
OUTPUT_DIR     = "output"
os.makedirs(OUTPUT_DIR, exist_ok=True)

# =========================================================================
# 1.  SET LABELS
# =========================================================================
SUPPLIER_NAMES = {1: "Moskova", 2: "Berlin", 3: "Oslo", 4: "Astana", 5: "Pekin"}
FACTORY_NAMES  = {1: "Madrid", 2: "St.Pete", 3: "Varşova", 4: "Ankara"}
DC_NAMES       = {1: "Ukrayna", 2: "Polonya", 3: "Romanya"}
CUSTOMER_NAMES = {1: "Müşteri_1", 2: "Müşteri_2", 3: "Müşteri_3",
                  4: "Müşteri_4", 5: "Müşteri_5", 6: "Müşteri_6"}
MODE_NAMES     = {1: "Demiryolu", 2: "Karayolu", 3: "Denizyolu", 4: "Havayolu"}

SUPPLIERS  = list(SUPPLIER_NAMES.keys())   # I = {1..5}
FACTORIES  = list(FACTORY_NAMES.keys())    # J = {1..4}
DCS        = list(DC_NAMES.keys())         # L = {1..3}
CUSTOMERS  = list(CUSTOMER_NAMES.keys())   # K = {1..6}
MODES      = list(MODE_NAMES.keys())       # M = {1..4}
PRODUCTS   = [1, 2, 3]                     # N = {1..3}
PERIODS    = [1, 2, 3, 4]                  # P = {1..4}

# =========================================================================
# 2.  GEOGRAPHIC TRANSPORTATION-MODE FEASIBILITY
# =========================================================================
FEASIBLE_MODES = {
    # ── Moskova (landlocked) ──────────────────────────────────────────────
    (1, 1): [1, 2, 4],
    (1, 2): [1, 2, 4],
    (1, 3): [1, 2, 4],
    (1, 4): [1, 2, 4],
    # ── Berlin ────────────────────────────────────────────────────────────
    (2, 1): [1, 2, 4],
    (2, 2): [1, 2, 4],
    (2, 3): [1, 2, 4],
    (2, 4): [1, 2, 4],
    # ── Oslo (North Sea access) ───────────────────────────────────────────
    (3, 1): [1, 2, 3, 4],
    (3, 2): [1, 2, 3, 4],
    (3, 3): [1, 2, 4],
    (3, 4): [1, 2, 3, 4],
    # ── Astana (double landlocked) ────────────────────────────────────────
    (4, 1): [1, 4],
    (4, 2): [1, 2, 4],
    (4, 3): [1, 4],
    (4, 4): [1, 2, 4],
    # ── Pekin ─────────────────────────────────────────────────────────────
    (5, 1): [1, 3, 4],
    (5, 2): [1, 3, 4],
    (5, 3): [1, 4],
    (5, 4): [1, 3, 4],
}

# =========================================================================
# 3.  DATA LOADING FROM EXCEL
# =========================================================================
def load_parameters():
    """
    Tüm sayısal parametreler Excel dosyalarından okunur.
    Modelde hiçbir hardcoded sayısal değer bulunmaz.
    """
    print("  [INFO] Reading parameters from Excel workbooks ...")

    # ── Scalar parameters ─────────────────────────────────────────────────
    df_scalar = pd.read_excel(PARAMS_FILE, sheet_name="SCALAR_PARAMS",
                              index_col=0, header=0)
    scalar  = df_scalar.iloc[:, 0].to_dict()
    big_M   = float(scalar.get("big_M", 1e6))
    if "eta" not in scalar or pd.isna(scalar["eta"]):
        raise ValueError(
            "SCALAR_PARAMS sheet is missing 'eta' (minimum service level, "
            "e.g. 0.95). Required by eq_21_service_level: without it the "
            "model can satisfy demand entirely via backlog and never "
            "actually deliver anything or open a DC."
        )
    eta = float(scalar["eta"])

    # ── Demand  d[k, n, p] ────────────────────────────────────────────────
    # Sheet layout: columns = [k, n, p, demand]
    df_demand = pd.read_excel(PARAMS_FILE, sheet_name="DEMAND",
                              index_col=False, header=0)
    demand = {}
    for _, row in df_demand.iterrows():
        key = (int(row["k"]), int(row["n"]), int(row["p"]))
        demand[key] = float(row["demand"])

    # ── Transportation cost  c_S[i, j, m, n] ─────────────────────────────
    # Sheet layout: columns = [i, j, m, n, cost]
    df_tc = pd.read_excel(TRANSPORT_FILE, sheet_name="TRANS_COST",
                          index_col=False, header=0)
    trans_cost = {}
    for _, row in df_tc.iterrows():
        key = (int(row["i"]), int(row["j"]), int(row["m"]), int(row["n"]))
        trans_cost[key] = float(row["cost"])

    # ── Factory-to-DC cost  c_F[j, l, n, p] ─────────────────────────────
    # Sheet layout: columns = [j, l, n, p, cost]
    df_fc = pd.read_excel(PARAMS_FILE, sheet_name="FACTORY_DC_COST",
                          index_col=False, header=0)
    factory_dc_cost = {}
    for _, row in df_fc.iterrows():
        key = (int(row["j"]), int(row["l"]), int(row["n"]), int(row["p"]))
        factory_dc_cost[key] = float(row["cost"])

    # ── DC-to-customer cost  c_D[l, k, n, p] ────────────────────────────
    # Sheet layout: columns = [l, k, n, p, cost]
    df_dc = pd.read_excel(PARAMS_FILE, sheet_name="DC_CUST_COST",
                          index_col=False, header=0)
    dc_cust_cost = {}
    for _, row in df_dc.iterrows():
        key = (int(row["l"]), int(row["k"]), int(row["n"]), int(row["p"]))
        dc_cust_cost[key] = float(row["cost"])

    # ── Supplier capacity  alpha[i, n, p] ────────────────────────────────
    # Sheet layout: columns = [i, n, p, capacity]
    df_sup = pd.read_excel(PARAMS_FILE, sheet_name="SUPPLIER_CAP",
                           index_col=False, header=0)
    supplier_cap = {}
    for _, row in df_sup.iterrows():
        key = (int(row["i"]), int(row["n"]), int(row["p"]))
        supplier_cap[key] = float(row["capacity"])

    # ── Factory production capacity  b[j, n, p] ──────────────────────────
    # Sheet layout: columns = [j, n, p, capacity]
    df_jcap = pd.read_excel(PARAMS_FILE, sheet_name="FACTORY_CAP",
                            index_col=False, header=0)
    factory_cap = {}
    for _, row in df_jcap.iterrows():
        key = (int(row["j"]), int(row["n"]), int(row["p"]))
        factory_cap[key] = float(row["capacity"])

    # ── Mode capacity  A[m, p] ────────────────────────────────────────────
    # Sheet layout: columns = [m, p, capacity]
    df_mode = pd.read_excel(PARAMS_FILE, sheet_name="MODE_CAP",
                            index_col=False, header=0)
    mode_cap = {}
    for _, row in df_mode.iterrows():
        key = (int(row["m"]), int(row["p"]))
        mode_cap[key] = float(row["capacity"])

    # ── DC throughput capacity  u[l, p] ──────────────────────────────────
    # Sheet layout: columns = [l, p, throughput]
    df_utput = pd.read_excel(PARAMS_FILE, sheet_name="DC_THROUGHPUT",
                             index_col=False, header=0)
    dc_throughput = {}
    for _, row in df_utput.iterrows():
        key = (int(row["l"]), int(row["p"]))
        dc_throughput[key] = float(row["throughput"])

    # ── DC storage capacity  v[l, p] ─────────────────────────────────────
    # Sheet layout: columns = [l, p, storage]
    df_stor = pd.read_excel(PARAMS_FILE, sheet_name="DC_STORAGE",
                            index_col=False, header=0)
    dc_storage = {}
    for _, row in df_stor.iterrows():
        key = (int(row["l"]), int(row["p"]))
        dc_storage[key] = float(row["storage"])

    # ── DC fixed investment cost  f[l, p] ────────────────────────────────
    # Sheet layout: columns = [l, p, cost]
    df_invest = pd.read_excel(PARAMS_FILE, sheet_name="DC_INVEST",
                              index_col=False, header=0)
    dc_invest = {}
    for _, row in df_invest.iterrows():
        key = (int(row["l"]), int(row["p"]))
        dc_invest[key] = float(row["cost"])

    # ── DC variable operating cost  g[l, p] ──────────────────────────────
    # Sheet layout: columns = [l, p, cost]
    df_opcost = pd.read_excel(PARAMS_FILE, sheet_name="DC_OPCOST",
                              index_col=False, header=0)
    dc_opcost = {}
    for _, row in df_opcost.iterrows():
        key = (int(row["l"]), int(row["p"]))
        dc_opcost[key] = float(row["cost"])

    # ── DC switching cost  sc[l] ──────────────────────────────────────────
    # Sheet layout: columns = [l, sc]
    df_switch = pd.read_excel(PARAMS_FILE, sheet_name="DC_SWITCH",
                              index_col=False, header=0)
    dc_switch = {}
    for _, row in df_switch.iterrows():
        dc_switch[int(row["l"])] = float(row["sc"])

    # ── Inventory holding cost  h[l, n, p] ───────────────────────────────
    # Sheet layout: columns = [l, n, p, cost]
    df_hold = pd.read_excel(PARAMS_FILE, sheet_name="HOLD_COST",
                            index_col=False, header=0)
    hold_cost = {}
    for _, row in df_hold.iterrows():
        key = (int(row["l"]), int(row["n"]), int(row["p"]))
        hold_cost[key] = float(row["cost"])

    # ── Backlog penalty cost  beta[k, n, p] ──────────────────────────────
    # Sheet layout: columns = [k, n, p, cost]
    df_back = pd.read_excel(PARAMS_FILE, sheet_name="BACK_COST",
                            index_col=False, header=0)
    back_cost = {}
    for _, row in df_back.iterrows():
        key = (int(row["k"]), int(row["n"]), int(row["p"]))
        back_cost[key] = float(row["cost"])

    print(f"  [INFO] big_M={big_M:.0f} | "
          f"demand entries={len(demand)} | "
          f"trans_cost entries={len(trans_cost)}")

    return {
        "big_M":          big_M,
        "eta":            eta,
        "demand":         demand,
        "trans_cost":     trans_cost,
        "factory_dc_cost":factory_dc_cost,
        "dc_cust_cost":   dc_cust_cost,
        "supplier_cap":   supplier_cap,
        "factory_cap":    factory_cap,
        "mode_cap":       mode_cap,
        "dc_throughput":  dc_throughput,
        "dc_storage":     dc_storage,
        "dc_invest":      dc_invest,
        "dc_opcost":      dc_opcost,
        "dc_switch":      dc_switch,
        "hold_cost":      hold_cost,
        "back_cost":      back_cost,
    }


# =========================================================================
# 4.  MODEL CONSTRUCTION
# =========================================================================
def build_model(params):
    M_big  = params["big_M"]
    eta    = params["eta"]
    D      = params["demand"]           # d[k,n,p]
    c_S    = params["trans_cost"]       # c_S[i,j,m,n]
    c_F    = params["factory_dc_cost"]  # c_F[j,l,n,p]
    c_D    = params["dc_cust_cost"]     # c_D[l,k,n,p]
    alpha  = params["supplier_cap"]     # alpha[i,n,p]
    b_cap  = params["factory_cap"]      # b[j,n,p]
    A      = params["mode_cap"]         # A[m,p]
    u_cap  = params["dc_throughput"]    # u[l,p]
    v_cap  = params["dc_storage"]       # v[l,p]
    f_inv  = params["dc_invest"]        # f[l,p]
    g_op   = params["dc_opcost"]        # g[l,p]
    sc     = params["dc_switch"]        # sc[l]
    h      = params["hold_cost"]        # h[l,n,p]
    beta   = params["back_cost"]        # beta[k,n,p]

    model = gp.Model("DistNet_MILP_Extended")
    model.Params.LogFile   = os.path.join(OUTPUT_DIR, "gurobi.log")
    model.Params.TimeLimit = 3600
    model.Params.MIPGap    = 0.01

    # ─────────────────────────────────────────────────────────────────────
    # 4a. DECISION VARIABLES  (LaTeX notasyonu ile birebir eşleşme)
    # ─────────────────────────────────────────────────────────────────────

    # x[i,j,m,n,p] : tedarikçi i → fabrika j, mod m, ürün n, dönem p
    x = {
        (i, j, m, n, p): model.addVar(lb=0.0, name=f"x_{i}_{j}_{m}_{n}_{p}")
        for i in SUPPLIERS
        for j in FACTORIES
        for m in MODES
        for n in PRODUCTS
        for p in PERIODS
        if m in FEASIBLE_MODES.get((i, j), [])
    }

    # w[j,l,n,p] : fabrika j → DC l, ürün n, dönem p  (DC indeksi var)
    w = {
        (j, l, n, p): model.addVar(lb=0.0, name=f"w_{j}_{l}_{n}_{p}")
        for j in FACTORIES
        for l in DCS
        for n in PRODUCTS
        for p in PERIODS
    }

    # y[l,k,n,p] : DC l → müşteri k, ürün n, dönem p
    y = {
        (l, k, n, p): model.addVar(lb=0.0, name=f"y_{l}_{k}_{n}_{p}")
        for l in DCS
        for k in CUSTOMERS
        for n in PRODUCTS
        for p in PERIODS
    }

    # q[l,n,p] : DC l'deki stok, ürün n, dönem p sonu
    # p=0 → başlangıç stoku (serbest değişken, sıfıra zorlanmaz)
    q = {
        (l, n, p): model.addVar(lb=0.0, name=f"q_{l}_{n}_{p}")
        for l in DCS
        for n in PRODUCTS
        for p in [0] + PERIODS
    }

    # b_bl[k,n,p] : müşteri k'nın birikmiş talebi, ürün n, dönem p sonu
    # p=0 → ufkun BAŞINDAKİ birikim, sabit 0 (lb=ub=0). Serbest bırakılırsa
    # (önceki sürümdeki hata) solver b_bl[k,n,0]'ı ufkun toplam talebi
    # kadar "hayali geçmiş borç" olarak seçip eq_8'in muhasebe özdeşliğiyle
    # bunu sıfıra indirebilir; bu da HİÇBİR teslimat yapılmadan (y=0) ve
    # HİÇBİR DC açılmadan modelin "talep karşılandı" gibi görünmesine yol
    # açar. b_bl[k,n,0]=0 sabitlemesi, ufka giren gerçek bir geçmiş borç
    # verisi olmadığını (ki bu veri setinde yok) ifade eder ve talebin
    # GERÇEKTEN teslim edilmesini (y>0, dolayısıyla açık bir DC) zorunlu
    # kılar.
    b_bl = {
        (k, n, p): model.addVar(
            lb=0.0, ub=(0.0 if p == 0 else GRB.INFINITY),
            name=f"b_{k}_{n}_{p}")
        for k in CUSTOMERS
        for n in PRODUCTS
        for p in [0] + PERIODS
    }

    # z[l,p] : DC l dönem p'de açık mı (binary)
    z = {
        (l, p): model.addVar(vtype=GRB.BINARY, name=f"z_{l}_{p}")
        for l in DCS
        for p in PERIODS
    }

    # u[l] : DC l'ye yatırım yapıldı mı (binary, tek seferlik — dönem
    # indeksi YOK). z[l,p] ile karıştırılmamalı: z dönemlik açık/kapalı
    # operasyon durumu, u ise ufkun başında verilen tek seferlik sermaye
    # yatırımı kararıdır (bkz. eq_10b bağlantı kısıtı ve amaç fonksiyonu).
    u = {
        l: model.addVar(vtype=GRB.BINARY, name=f"u_{l}")
        for l in DCS
    }

    # delta[l,p] : DC l dönem p'de statü değiştirdi mi (binary)
    delta = {
        (l, p): model.addVar(vtype=GRB.BINARY, name=f"delta_{l}_{p}")
        for l in DCS
        for p in PERIODS
    }

    model.update()

    # ─────────────────────────────────────────────────────────────────────
    # 4b. OBJECTIVE FUNCTION  — LaTeX eq_1 ile birebir
    #
    # min Z = Σ c_S * x  +  Σ c_F * w  +  Σ c_D * y
    #       + Σ f*u  +  Σ g*Σy  +  Σ h*q  +  Σ beta*b  +  Σ sc*delta
    # ─────────────────────────────────────────────────────────────────────
    obj = gp.LinExpr()

    # Tedarikçi → fabrika nakliye maliyeti
    for (i, j, m, n, p), var in x.items():
        obj += c_S.get((i, j, m, n), 0.0) * var

    # Fabrika → DC nakliye maliyeti
    for (j, l, n, p), var in w.items():
        obj += c_F.get((j, l, n, p), 0.0) * var

    # DC → müşteri nakliye maliyeti
    for (l, k, n, p), var in y.items():
        obj += c_D.get((l, k, n, p), 0.0) * var

    # DC sabit yatırım maliyeti  f[l,1] * u[l]  — TEK SEFERLİK.
    # u[l]=1 <=> DC l'ye ufkun başında sermaye yatırımı yapılmış demektir;
    # bu maliyet, DC'nin kaç dönem açık kaldığından bağımsız olarak sadece
    # bir kez tahsil edilir (DC_INVEST sayfasındaki dönem-1 maliyeti
    # kullanılır). Önceki sürümde bu terim f[l,p]*z[l,p] olarak HER
    # dönem tekrar tekrar tahsil ediliyordu, ki bu "yatırım" değil
    # "işletme gideri" anlamına geliyordu.
    for l in DCS:
        obj += f_inv.get((l, 1), 0.0) * u[l]

    # DC değişken işletme maliyeti  g[l,p] * Σ_{n,k} y[l,k,n,p]
    for l in DCS:
        for p in PERIODS:
            total_y = gp.quicksum(y[l, k, n, p]
                                  for k in CUSTOMERS
                                  for n in PRODUCTS)
            obj += g_op.get((l, p), 0.0) * total_y

    # Stok tutma maliyeti  h[l,n,p] * q[l,n,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                obj += h.get((l, n, p), 0.0) * q[l, n, p]

    # Gecikme ceza maliyeti  beta[k,n,p] * b[k,n,p]
    for k in CUSTOMERS:
        for n in PRODUCTS:
            for p in PERIODS:
                obj += beta.get((k, n, p), 0.0) * b_bl[k, n, p]

    # DC statü değişim maliyeti  sc[l] * delta[l,p]
    for l in DCS:
        for p in PERIODS:
            obj += sc.get(l, 0.0) * delta[l, p]

    model.setObjective(obj, GRB.MINIMIZE)

    # ─────────────────────────────────────────────────────────────────────
    # 4c. CONSTRAINTS  — LaTeX kısıtlarıyla birebir eşleşme
    # ─────────────────────────────────────────────────────────────────────

    # ── Constraint eq_2 : Tedarikçi kapasite kısıtı ──────────────────────
    # Σ_{j,m} x[i,j,m,n,p] <= alpha[i,n,p]
    for i in SUPPLIERS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(x[i, j, m, n, p]
                                for j in FACTORIES
                                for m in MODES
                                if (i, j, m, n, p) in x)
                    <= alpha.get((i, n, p), GRB.INFINITY),
                    name=f"eq2_supplier_cap_{i}_{n}_{p}"
                )

    # ── Constraint eq_3 : Fabrika üretim kapasitesi ───────────────────────
    # Σ_l w[j,l,n,p] <= b[j,n,p]
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(w[j, l, n, p] for l in DCS)
                    <= b_cap.get((j, n, p), GRB.INFINITY),
                    name=f"eq3_factory_cap_{j}_{n}_{p}"
                )

    # ── Constraint eq_4 : Fabrikada akış dengesi ──────────────────────────
    # Σ_{i,m} x[i,j,m,n,p] = Σ_l w[j,l,n,p]
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                inbound = gp.quicksum(x[i, j, m, n, p]
                                      for i in SUPPLIERS
                                      for m in MODES
                                      if (i, j, m, n, p) in x)
                outbound = gp.quicksum(w[j, l, n, p] for l in DCS)
                model.addConstr(
                    inbound == outbound,
                    name=f"eq4_flow_balance_{j}_{n}_{p}"
                )

    # ── Constraint eq_5 : DC envanter dengesi ─────────────────────────────
    # q[l,n,p] = q[l,n,p-1] + Σ_j w[j,l,n,p] - Σ_k y[l,k,n,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    q[l, n, p]
                    == q[l, n, p - 1]
                    + gp.quicksum(w[j, l, n, p] for j in FACTORIES)
                    - gp.quicksum(y[l, k, n, p] for k in CUSTOMERS),
                    name=f"eq5_inv_balance_{l}_{n}_{p}"
                )

    # ── Constraint eq_6 : DC depo kapasitesi ─────────────────────────────
    # Σ_n q[l,n,p] <= v[l,p]
    for l in DCS:
        for p in PERIODS:
            model.addConstr(
                gp.quicksum(q[l, n, p] for n in PRODUCTS)
                <= v_cap.get((l, p), GRB.INFINITY),
                name=f"eq6_storage_cap_{l}_{p}"
            )

    # ── Constraint eq_7 : DC işlem kapasitesi ────────────────────────────
    # Σ_{k,n} y[l,k,n,p] <= u[l,p]
    for l in DCS:
        for p in PERIODS:
            model.addConstr(
                gp.quicksum(y[l, k, n, p]
                            for k in CUSTOMERS
                            for n in PRODUCTS)
                <= u_cap.get((l, p), GRB.INFINITY),
                name=f"eq7_throughput_cap_{l}_{p}"
            )

    # ── Constraint eq_8 : Talep karşılama (backlog ile) ───────────────────
    # Σ_l y[l,k,n,p] + b[k,n,p-1] - b[k,n,p] = d[k,n,p]
    for k in CUSTOMERS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(y[l, k, n, p] for l in DCS)
                    + b_bl[k, n, p - 1]
                    - b_bl[k, n, p]
                    == D.get((k, n, p), 0.0),
                    name=f"eq8_demand_{k}_{n}_{p}"
                )

    # ── Constraint eq_9 : Taşıma modu kapasitesi ─────────────────────────
    # Σ_{i,j} x[i,j,m,n,p] <= A[m,p]
    for m in MODES:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(x[i, j, m, n, p]
                                for i in SUPPLIERS
                                for j in FACTORIES
                                if (i, j, m, n, p) in x)
                    <= A.get((m, p), GRB.INFINITY),
                    name=f"eq9_mode_cap_{m}_{n}_{p}"
                )

    # ── Constraint eq_10 : DC girişi → açık DC bağı (Big-M) ──────────────
    # Σ_j w[j,l,n,p] <= M_big * z[l,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(w[j, l, n, p] for j in FACTORIES)
                    <= M_big * z[l, p],
                    name=f"eq10_dc_inflow_link_{l}_{n}_{p}"
                )

    # ── Constraint eq_11 : DC çıkışı → açık DC bağı (Big-M) ─────────────
    # Σ_k y[l,k,n,p] <= M_big * z[l,p]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(y[l, k, n, p] for k in CUSTOMERS)
                    <= M_big * z[l, p],
                    name=f"eq11_dc_outflow_link_{l}_{n}_{p}"
                )

    # ── Constraints eq_12 & eq_13 : Switching linearizasyonu (her iki yön)
    # delta[l,p] >= z[l,p] - z[l,p-1]   (kapanmış → açılmış)
    # delta[l,p] >= z[l,p-1] - z[l,p]   (açılmış → kapanmış)
    for l in DCS:
        for p in PERIODS:
            z_prev = z[l, p - 1] if p > 1 else gp.LinExpr(0)
            model.addConstr(
                delta[l, p] >= z[l, p] - z_prev,
                name=f"eq12_switch_open_{l}_{p}"
            )
            model.addConstr(
                delta[l, p] >= z_prev - z[l, p],
                name=f"eq13_switch_close_{l}_{p}"
            )

    # ── Constraint eq_10b : açık DC → yatırım bağı ────────────────────────
    # z[l,p] <= u[l]   (yatırım yapılmamış bir DC hiçbir dönem açık olamaz)
    for l in DCS:
        for p in PERIODS:
            model.addConstr(
                z[l, p] <= u[l],
                name=f"eq10b_dc_invest_link_{l}_{p}"
            )

    # ── Constraint eq_21 : Minimum hizmet düzeyi kısıtı ───────────────────
    # Σ_l y[l,k,n,p] >= eta * d[k,n,p]
    # b_bl[k,n,0]=0 sabitlemesi tek başına modelin talebi tamamen (dönemler
    # boyunca artan) backlog'a yıkıp hiç teslimat yapmamasını ENGELLEMEZ —
    # bu, veri setindeki backlog cezası DC yatırımından ucuzsa yine de
    # matematiksel olarak "geçerli" (kazık değil, gerçek) bir optimal
    # sonuç olabilir. eta, bunun bir iş kuralı olarak kabul edilemeyeceğini
    # (müşterilere süresiz servis reddi olamayacağını) açıkça zorunlu kılar
    # ve dolaylı olarak gerekli DC'lerin açılmasını (z, dolayısıyla u) tetikler.
    for k in CUSTOMERS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    gp.quicksum(y[l, k, n, p] for l in DCS)
                    >= eta * D.get((k, n, p), 0.0),
                    name=f"eq21_service_level_{k}_{n}_{p}"
                )

    # ── Constraints eq_14 & eq_15 : Non-negativity & binary integrality ───
    # lb=0 değişken tanımında zaten uygulandı.
    # Binary kısıtı vtype=GRB.BINARY ile uygulandı.

    model.update()  # NumConstrs/NumVars main()'de doğru basılsın diye
    return model, x, w, y, q, b_bl, z, u, delta


# =========================================================================
# 5.  RESULTS EXTRACTION & EXCEL OUTPUT
# =========================================================================
def export_results(model, x, w, y, q, b_bl, z, u, delta, params):
    status = model.Status
    if status not in (GRB.OPTIMAL, GRB.TIME_LIMIT, GRB.SUBOPTIMAL):
        print(f"  [WARN] Model status = {status}. No feasible solution found.")
        return

    obj_val = model.ObjVal
    print(f"\n  ── Objective value : {obj_val:,.2f}")
    print(f"  ── MIP gap         : {model.MIPGap * 100:.4f} %")
    print(f"  ── Solve time      : {model.Runtime:.2f} s\n")

    ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
    out = os.path.join(OUTPUT_DIR, f"results_{ts}.xlsx")

    with pd.ExcelWriter(out, engine="openpyxl") as writer:

        # Sheet 1: x – Tedarikçi → fabrika akışları
        rows_x = []
        for (i, j, m, n, p), var in x.items():
            val = var.X
            if val > 1e-6:
                rows_x.append({
                    "Supplier": SUPPLIER_NAMES[i],
                    "Factory":  FACTORY_NAMES[j],
                    "Mode":     MODE_NAMES[m],
                    "Product":  n,
                    "Period":   p,
                    "Quantity": round(val, 4),
                })
        pd.DataFrame(rows_x).to_excel(writer, sheet_name="x_flows", index=False)

        # Sheet 2: w – Fabrika → DC akışları
        rows_w = []
        for (j, l, n, p), var in w.items():
            val = var.X
            if val > 1e-6:
                rows_w.append({
                    "Factory": FACTORY_NAMES[j],
                    "DC":      DC_NAMES[l],
                    "Product": n,
                    "Period":  p,
                    "Quantity":round(val, 4),
                })
        pd.DataFrame(rows_w).to_excel(writer, sheet_name="w_factory_dc", index=False)

        # Sheet 3: y – DC → müşteri teslimatları
        rows_y = []
        for (l, k, n, p), var in y.items():
            val = var.X
            if val > 1e-6:
                rows_y.append({
                    "DC":       DC_NAMES[l],
                    "Customer": CUSTOMER_NAMES[k],
                    "Product":  n,
                    "Period":   p,
                    "Delivery": round(val, 4),
                })
        pd.DataFrame(rows_y).to_excel(writer, sheet_name="y_delivery", index=False)

        # Sheet 4: q – Stok seviyeleri
        rows_q = []
        for l in DCS:
            for n in PRODUCTS:
                for p in [0] + PERIODS:
                    rows_q.append({
                        "DC":        DC_NAMES[l],
                        "Product":   n,
                        "Period":    p,
                        "Inventory": round(q[l, n, p].X, 4),
                    })
        pd.DataFrame(rows_q).to_excel(writer, sheet_name="q_inventory", index=False)

        # Sheet 5: b – Backlog seviyeleri
        rows_b = []
        for k in CUSTOMERS:
            for n in PRODUCTS:
                for p in [0] + PERIODS:
                    val = b_bl[k, n, p].X
                    if val > 1e-6:
                        rows_b.append({
                            "Customer": CUSTOMER_NAMES[k],
                            "Product":  n,
                            "Period":   p,
                            "Backlog":  round(val, 4),
                        })
        pd.DataFrame(rows_b).to_excel(writer, sheet_name="b_backlog", index=False)

        # Sheet 6: z – DC açık/kapalı durumu
        rows_z = []
        for l in DCS:
            for p in PERIODS:
                rows_z.append({
                    "DC":     DC_NAMES[l],
                    "Period": p,
                    "Open":   int(z[l, p].X + 0.5),
                })
        pd.DataFrame(rows_z).to_excel(writer, sheet_name="z_DC_status", index=False)

        # Sheet 6b: u – DC yatırım kararı (tek seferlik, dönemsiz)
        rows_u = []
        for l in DCS:
            rows_u.append({
                "DC":       DC_NAMES[l],
                "Invested": int(u[l].X + 0.5),
            })
        pd.DataFrame(rows_u).to_excel(writer, sheet_name="u_DC_invest", index=False)

        # Sheet 7: delta – DC statü değişim olayları
        rows_d = []
        for l in DCS:
            for p in PERIODS:
                if delta[l, p].X > 0.5:
                    rows_d.append({
                        "DC":     DC_NAMES[l],
                        "Period": p,
                        "Switch": 1,
                    })
        pd.DataFrame(rows_d).to_excel(writer, sheet_name="delta_switch", index=False)

        # Sheet 8: Dönem bazlı maliyet dökümü
        rows_cost = []
        for p in PERIODS:
            trans_p = sum(
                params["trans_cost"].get((i, j, m, n), 0.0) * x[i, j, m, n, p].X
                for i in SUPPLIERS for j in FACTORIES
                for m in MODES for n in PRODUCTS
                if (i, j, m, n, p) in x
            )
            fdc_p = sum(
                params["factory_dc_cost"].get((j, l, n, p), 0.0) * w[j, l, n, p].X
                for j in FACTORIES for l in DCS for n in PRODUCTS
            )
            dcc_p = sum(
                params["dc_cust_cost"].get((l, k, n, p), 0.0) * y[l, k, n, p].X
                for l in DCS for k in CUSTOMERS for n in PRODUCTS
            )
            hold_p = sum(
                params["hold_cost"].get((l, n, p), 0.0) * q[l, n, p].X
                for l in DCS for n in PRODUCTS
            )
            back_p = sum(
                params["back_cost"].get((k, n, p), 0.0) * b_bl[k, n, p].X
                for k in CUSTOMERS for n in PRODUCTS
            )
            # DC yatırım maliyeti tek seferlik: sadece dönem 1'in
            # satırında görünür (amaç fonksiyonuyla tutarlı kalması için).
            invest_p = sum(
                params["dc_invest"].get((l, 1), 0.0) * u[l].X
                for l in DCS
            ) if p == 1 else 0.0
            switch_p = sum(
                params["dc_switch"].get(l, 0.0) * delta[l, p].X
                for l in DCS
            )
            rows_cost.append({
                "Period":    p,
                "Transport": round(trans_p,  2),
                "Factory_DC":round(fdc_p,    2),
                "DC_Cust":   round(dcc_p,    2),
                "Holding":   round(hold_p,   2),
                "Backlog":   round(back_p,   2),
                "DC_Invest": round(invest_p, 2),
                "DC_Switch": round(switch_p, 2),
                "Total":     round(trans_p + fdc_p + dcc_p +
                                   hold_p + back_p +
                                   invest_p + switch_p, 2),
            })
        rows_cost.append({"Period": "GRAND TOTAL",
                          "Total": round(model.ObjVal, 2)})
        pd.DataFrame(rows_cost).to_excel(writer, sheet_name="CostByPeriod", index=False)

        # Sheet 9: Ürün özet tablosu
        rows_prod = []
        for n in PRODUCTS:
            for p in PERIODS:
                total_prod = sum(
                    w[j, l, n, p].X for j in FACTORIES for l in DCS
                )
                total_del = sum(
                    y[l, k, n, p].X for l in DCS for k in CUSTOMERS
                )
                total_inv = sum(q[l, n, p].X for l in DCS)
                total_bl  = sum(b_bl[k, n, p].X for k in CUSTOMERS)
                total_dem = sum(
                    params["demand"].get((k, n, p), 0.0) for k in CUSTOMERS
                )
                rows_prod.append({
                    "Product":    n,
                    "Period":     p,
                    "Production": round(total_prod, 4),
                    "Delivery":   round(total_del,  4),
                    "Inventory":  round(total_inv,  4),
                    "Backlog":    round(total_bl,   4),
                    "Demand":     round(total_dem,  4),
                })
        pd.DataFrame(rows_prod).to_excel(writer, sheet_name="ProductSummary", index=False)

    print(f"  [INFO] Results exported → {out}")


# =========================================================================
# 6.  MAIN
# =========================================================================
def main():
    print("=" * 65)
    print("  Distribution Network Design  –  Extended MILP")
    print("  Seçmen, Öncan & Tuna (2015) | DEÜ Lojistik")
    print("=" * 65)

    params = load_parameters()

    print("\n  [INFO] Building model ...")
    model, x, w, y, q, b_bl, z, u, delta = build_model(params)

    print(f"  [INFO] Variables   : {model.NumVars:,}")
    print(f"  [INFO] Constraints : {model.NumConstrs:,}")
    print(f"  [INFO] Binary vars : {model.NumBinVars:,}")
    print(f"  [INFO] Feasible x-triples (geographic filter): {len(x):,}")
    print("\n  [INFO] Solving ...\n")

    model.optimize()

    export_results(model, x, w, y, q, b_bl, z, u, delta, params)

    print("\n  Done.\n")


if __name__ == "__main__":
    main()
