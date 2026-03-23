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
# 1.  SET LABELS (for readability in output)
# =========================================================================
SUPPLIER_NAMES = {1: "Moskova", 2: "Berlin", 3: "Oslo", 4: "Astana", 5: "Pekin"}
FACTORY_NAMES  = {1: "Madrid", 2: "St.Pete", 3: "Varşova", 4: "Ankara"}
DC_NAMES       = {1: "Ukrayna", 2: "Polonya", 3: "Romanya"}
MODE_NAMES     = {1: "Demiryolu", 2: "Karayolu", 3: "Denizyolu", 4: "Havayolu"}

SUPPLIERS  = list(SUPPLIER_NAMES.keys())   # I = {1..5}
FACTORIES  = list(FACTORY_NAMES.keys())    # J = {1..4}
DCS        = list(DC_NAMES.keys())         # L = {1..3}
MODES      = list(MODE_NAMES.keys())       # T = {1..4}
PRODUCTS   = [1, 2, 3]                     # N = {1..3}
PERIODS    = [1, 2, 3, 4]                  # P = {1..4}

# =========================================================================
# 2.  GEOGRAPHIC TRANSPORTATION-MODE FEASIBILITY
#     Key   : (supplier_index, factory_index)
#     Value : list of feasible mode indices
#     Source: Geographic network analysis (Eurasian supply routes)
# =========================================================================
FEASIBLE_MODES = {
    # ── Moskova (landlocked) ──────────────────────────────────────────────
    (1, 1): [1, 2, 4],       # Moskova → Madrid   : rail, road, air
    (1, 2): [1, 2, 4],       # Moskova → St.Pete  : rail, road, air
    (1, 3): [1, 2, 4],       # Moskova → Varşova  : rail, road, air
    (1, 4): [1, 2, 4],       # Moskova → Ankara   : rail, road, air
    # ── Berlin ────────────────────────────────────────────────────────────
    (2, 1): [1, 2, 4],       # Berlin  → Madrid   : rail, road, air
    (2, 2): [1, 2, 4],       # Berlin  → St.Pete  : rail, road, air
    (2, 3): [1, 2, 4],       # Berlin  → Varşova  : rail, road, air
    (2, 4): [1, 2, 4],       # Berlin  → Ankara   : rail, road, air
    # ── Oslo (North Sea access) ───────────────────────────────────────────
    (3, 1): [1, 2, 3, 4],    # Oslo    → Madrid   : ALL modes
    (3, 2): [1, 2, 3, 4],    # Oslo    → St.Pete  : ALL modes
    (3, 3): [1, 2, 4],       # Oslo    → Varşova  : rail, road, air (inland)
    (3, 4): [1, 2, 3, 4],    # Oslo    → Ankara   : ALL modes (Mediterranean)
    # ── Astana (double landlocked) ────────────────────────────────────────
    (4, 1): [1, 4],          # Astana  → Madrid   : rail, air only
    (4, 2): [1, 2, 4],       # Astana  → St.Pete  : rail, road, air
    (4, 3): [1, 4],          # Astana  → Varşova  : rail, air only
    (4, 4): [1, 2, 4],       # Astana  → Ankara   : rail, road, air
    # ── Pekin (China – sea / rail / air feasible) ─────────────────────────
    (5, 1): [1, 3, 4],       # Pekin   → Madrid   : rail, sea, air
    (5, 2): [1, 3, 4],       # Pekin   → St.Pete  : rail, sea, air
    (5, 3): [1, 4],          # Pekin   → Varşova  : rail, air (no practical sea)
    (5, 4): [1, 3, 4],       # Pekin   → Ankara   : rail, sea, air
}

# =========================================================================
# 3.  DATA LOADING FROM EXCEL
# =========================================================================
def load_parameters():
    """
    Read all numerical parameters from the two Excel workbooks.
    All scalar rates, costs, capacities, and demand values must be
    defined in the workbooks – no hardcoded numerics in model logic.
    """
    print("  [INFO] Reading parameters from Excel workbooks …")

    # ── Scalar parameters (pi_rate, big_M, etc.) ─────────────────────────
    df_scalar = pd.read_excel(PARAMS_FILE, sheet_name="SCALAR_PARAMS",
                              index_col=0, header=0)
    scalar = df_scalar.iloc[:, 0].to_dict()
    pi_rate = float(scalar.get("pi_rate", 0.10))
    big_M   = float(scalar.get("big_M",  1e6))

    # ── Demand  D[n, p] ───────────────────────────────────────────────────
    df_demand = pd.read_excel(PARAMS_FILE, sheet_name="DEMAND",
                              index_col=0, header=0)
    # rows = products, cols = periods
    demand = {(int(n), int(p)): float(df_demand.loc[n, p])
              for n in df_demand.index
              for p in df_demand.columns}

    # ── Production cost  c_prod[j, n, p] ─────────────────────────────────
    df_pc = pd.read_excel(PARAMS_FILE, sheet_name="PROD_COST",
                          index_col=0, header=0)
    prod_cost = {(int(j), int(n)): float(df_pc.loc[j, n])
                 for j in df_pc.index
                 for n in df_pc.columns}

    # ── Factory capacity  CAP_J[j, p] ─────────────────────────────────────
    df_jcap = pd.read_excel(PARAMS_FILE, sheet_name="FACTORY_CAP",
                            index_col=0, header=0)
    factory_cap = {(int(j), int(p)): float(df_jcap.loc[j, p])
                   for j in df_jcap.index
                   for p in df_jcap.columns}

    # ── DC capacity  CAP_L[l, p] ──────────────────────────────────────────
    df_lcap = pd.read_excel(PARAMS_FILE, sheet_name="DC_CAP",
                            index_col=0, header=0)
    dc_cap = {(int(l), int(p)): float(df_lcap.loc[l, p])
              for l in df_lcap.index
              for p in df_lcap.columns}

    # ── DC investment cost  K[l] ──────────────────────────────────────────
    df_invest = pd.read_excel(PARAMS_FILE, sheet_name="DC_INVEST",
                              index_col=0, header=0)
    dc_invest = {int(l): float(df_invest.loc[l, "K"])
                 for l in df_invest.index}

    # ── DC open/close switching cost  SC[l] ───────────────────────────────
    df_switch = pd.read_excel(PARAMS_FILE, sheet_name="DC_SWITCH",
                              index_col=0, header=0)
    dc_switch = {int(l): float(df_switch.loc[l, "SC"])
                 for l in df_switch.index}

    # ── Inventory holding cost  h[l, n] ──────────────────────────────────
    df_hold = pd.read_excel(PARAMS_FILE, sheet_name="HOLD_COST",
                            index_col=0, header=0)
    hold_cost = {(int(l), int(n)): float(df_hold.loc[l, n])
                 for l in df_hold.index
                 for n in df_hold.columns}

    # ── Backlog penalty  b_pen[l, n] ─────────────────────────────────────
    df_back = pd.read_excel(PARAMS_FILE, sheet_name="BACK_COST",
                            index_col=0, header=0)
    back_cost = {(int(l), int(n)): float(df_back.loc[l, n])
                 for l in df_back.index
                 for n in df_back.columns}

    # ── Transportation cost  c_trans[i, j, t, n] ─────────────────────────
    df_tc = pd.read_excel(TRANSPORT_FILE, sheet_name="TRANS_COST",
                          index_col=False, header=0)
    # Multi-index column (supplier, factory, mode, product) packed in rows
    # Expected sheet layout: columns = [i, j, t, n, cost]
    trans_cost = {}
    for _, row in df_tc.iterrows():
        key = (int(row["i"]), int(row["j"]), int(row["t"]), int(row["n"]))
        trans_cost[key] = float(row["cost"])

    print(f"  [INFO] pi_rate={pi_rate:.4f} | big_M={big_M:.0f} | "
          f"demand entries={len(demand)} | trans_cost entries={len(trans_cost)}")

    return {
        "pi_rate":    pi_rate,
        "big_M":      big_M,
        "demand":     demand,
        "prod_cost":  prod_cost,
        "factory_cap":factory_cap,
        "dc_cap":     dc_cap,
        "dc_invest":  dc_invest,
        "dc_switch":  dc_switch,
        "hold_cost":  hold_cost,
        "back_cost":  back_cost,
        "trans_cost": trans_cost,
    }


# =========================================================================
# 4.  MODEL CONSTRUCTION
# =========================================================================
def build_model(params):
    pi   = params["pi_rate"]
    M    = params["big_M"]
    D    = params["demand"]
    c_p  = params["prod_cost"]
    c_t  = params["trans_cost"]
    CAP_J= params["factory_cap"]
    CAP_L= params["dc_cap"]
    K    = params["dc_invest"]
    SC   = params["dc_switch"]
    h    = params["hold_cost"]
    b    = params["back_cost"]

    model = gp.Model("DistNet_MILP_Extended")
    model.Params.LogFile   = os.path.join(OUTPUT_DIR, "gurobi.log")
    model.Params.TimeLimit = 3600      # 1-hour wall-clock limit
    model.Params.MIPGap    = 0.01      # 1 % optimality gap

    # ─────────────────────────────────────────────────────────────────────
    # 4a. DECISION VARIABLES
    # ─────────────────────────────────────────────────────────────────────

    # X[i,j,t,n,p] : shipment quantity from supplier i to factory j
    #                via mode t, product n, period p
    #                (only geographically feasible (i,j,t) triples)
    X = {
        (i, j, t, n, p): model.addVar(lb=0.0, name=f"X{i}{j}{t}{n}{p}")
        for i in SUPPLIERS
        for j in FACTORIES
        for t in MODES
        for n in PRODUCTS
        for p in PERIODS
        if t in FEASIBLE_MODES.get((i, j), [])   # ← geographic filter
    }

    # W[j,n,p] : production quantity at factory j, product n, period p
    W = {
        (j, n, p): model.addVar(lb=0.0, name=f"W{j}{n}{p}")
        for j in FACTORIES
        for n in PRODUCTS
        for p in PERIODS
    }

    # Y[l,n,p] : delivery from DC l to market, product n, period p
    Y = {
        (l, n, p): model.addVar(lb=0.0, name=f"Y{l}{n}{p}")
        for l in DCS
        for n in PRODUCTS
        for p in PERIODS
    }

    # B[l,n,p] : backlog at DC l, product n, period p
    #            B[l,n,0] is a free variable – not forced to zero
    B = {
        (l, n, p): model.addVar(lb=0.0, name=f"B{l}{n}{p}")
        for l in DCS
        for n in PRODUCTS
        for p in [0] + PERIODS          # p=0 is the initial backlog
    }

    # Q[l,n,p] : inventory at DC l, product n, end of period p
    #            Q[l,n,0] is the initial inventory
    Q = {
        (l, n, p): model.addVar(lb=0.0, name=f"Q{l}{n}{p}")
        for l in DCS
        for n in PRODUCTS
        for p in [0] + PERIODS
    }

    # y[l] : binary – invest in (open) DC l (strategic, period-0 decision)
    y = {
        l: model.addVar(vtype=GRB.BINARY, name=f"y{l}")
        for l in DCS
    }

    # z[l,p] : binary – DC l is operational in period p
    z = {
        (l, p): model.addVar(vtype=GRB.BINARY, name=f"z{l}{p}")
        for l in DCS
        for p in PERIODS
    }

    # delta[l,p] : binary – DC l opens in period p (switching indicator)
    #              1 if z[l,p]=1 and z[l,p-1]=0
    delta = {
        (l, p): model.addVar(vtype=GRB.BINARY, name=f"delta{l}{p}")
        for l in DCS
        for p in PERIODS
    }

    model.update()

    # ─────────────────────────────────────────────────────────────────────
    # 4b. OBJECTIVE FUNCTION
    #     Minimise: transportation + production + inventory holding
    #               + backlog penalty + DC investment + DC switching
    # ─────────────────────────────────────────────────────────────────────
    obj = gp.LinExpr()

    # Transportation cost
    for (i, j, t, n, p), var in X.items():
        key = (i, j, t, n)
        if key in c_t:
            obj += c_t[key] * var

    # Production cost
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                key = (j, n)
                if key in c_p:
                    obj += c_p[key] * W[j, n, p]

    # Inventory holding cost  (end-of-period inventory × rate × unit_value)
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                obj += h.get((l, n), pi) * Q[l, n, p]

    # Backlog penalty
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                obj += b.get((l, n), 0.0) * B[l, n, p]

    # DC one-time investment cost (paid at start of planning horizon)
    for l in DCS:
        obj += K.get(l, 0.0) * y[l]

    # DC open/close switching cost
    for l in DCS:
        for p in PERIODS:
            obj += SC.get(l, 0.0) * delta[l, p]

    model.setObjective(obj, GRB.MINIMIZE)

    # ─────────────────────────────────────────────────────────────────────
    # 4c. CONSTRAINTS
    # ─────────────────────────────────────────────────────────────────────

    # C1 – Raw material supply balance at factory j
    #      Sum of inbound shipments = production (with 1-period lead time)
    for j in FACTORIES:
        for n in PRODUCTS:
            for p in PERIODS:
                inbound = gp.quicksum(
                    X[i, j, t, n, p]
                    for i in SUPPLIERS
                    for t in MODES
                    if (i, j, t, n, p) in X
                )
                if p == 1:
                    model.addConstr(inbound == W[j, n, p],
                                    name=f"C1_supply_{j}_{n}_{p}")
                else:
                    model.addConstr(inbound == W[j, n, p],
                                    name=f"C1_supply_{j}_{n}_{p}")

    # C2 – Inventory balance at DC l (Fj = 1 period production delay)
    #      Q[l,n,p] = Q[l,n,p-1] + (sum_j W[j,n,p-1]) - Y[l,n,p] + B[l,n,p] - B[l,n,p-1]
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                production_prev = gp.quicksum(W[j, n, p - 1] for j in FACTORIES) \
                    if p > 1 else gp.LinExpr(0)
                # (for p=1 there is no W[j,n,0] – assume zero initial production)
                model.addConstr(
                    Q[l, n, p] == Q[l, n, p - 1] + production_prev
                                  - Y[l, n, p] + B[l, n, p] - B[l, n, p - 1],
                    name=f"C2_inv_balance_{l}_{n}_{p}"
                )

    # C3 – Demand satisfaction at DC l
    for l in DCS:
        for n in PRODUCTS:
            for p in PERIODS:
                model.addConstr(
                    Y[l, n, p] + B[l, n, p] >= D.get((n, p), 0),
                    name=f"C3_demand_{l}_{n}_{p}"
                )

    # C4 – Factory production capacity
    for j in FACTORIES:
        for p in PERIODS:
            model.addConstr(
                gp.quicksum(W[j, n, p] for n in PRODUCTS)
                <= CAP_J.get((j, p), GRB.INFINITY),
                name=f"C4_factory_cap_{j}_{p}"
            )

    # C5 – DC throughput capacity (only if DC is open)
    for l in DCS:
        for p in PERIODS:
            cap_lp = CAP_L.get((l, p), GRB.INFINITY)
            model.addConstr(
                gp.quicksum(Y[l, n, p] for n in PRODUCTS)
                <= cap_lp * z[l, p],
                name=f"C5_dc_cap_{l}_{p}"
            )

    # C6 – DC can only operate if investment was made
    for l in DCS:
        for p in PERIODS:
            model.addConstr(z[l, p] <= y[l],
                            name=f"C6_invest_link_{l}_{p}")

    # C7 – Switching indicator (open in p but not in p-1)
    #      delta[l,p] >= z[l,p] - z[l,p-1]
    for l in DCS:
        for p in PERIODS:
            z_prev = z[l, p - 1] if p > 1 else gp.LinExpr(0)
            model.addConstr(
                delta[l, p] >= z[l, p] - z_prev,
                name=f"C7_switch_{l}_{p}"
            )

    # C8 – Symmetry breaking (if DC l is open, prefer lower-indexed DCs)
    for l in DCS[:-1]:
        model.addConstr(y[l] >= y[l + 1],
                        name=f"C8_symmetry_{l}")

    # C9 – Valid inequality: at least one DC must be open per period
    for p in PERIODS:
        model.addConstr(
            gp.quicksum(z[l, p] for l in DCS) >= 1,
            name=f"C9_min_dc_{p}"
        )

    # C10 – Non-negativity of inventory and backlog are handled by lb=0
    #        except Q[l,n,0] and B[l,n,0] which are initial conditions
    #        (left as free variables – values driven by data)

    return model, X, W, Y, B, Q, y, z, delta


# =========================================================================
# 5.  RESULTS EXTRACTION & EXCEL OUTPUT  (10 sheets)
# =========================================================================
def export_results(model, X, W, Y, B, Q, y, z, delta, params):
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

        # ── Sheet 1: X – Transportation flows ────────────────────────────
        rows_x = []
        for (i, j, t, n, p), var in X.items():
            val = var.X
            if val > 1e-6:
                rows_x.append({
                    "Supplier": SUPPLIER_NAMES[i],
                    "Factory":  FACTORY_NAMES[j],
                    "Mode":     MODE_NAMES[t],
                    "Product":  n,
                    "Period":   p,
                    "Quantity": round(val, 4),
                })
        pd.DataFrame(rows_x).to_excel(writer, sheet_name="X_flows",       index=False)

        # ── Sheet 2: W – Production quantities ───────────────────────────
        rows_w = [{"Factory": FACTORY_NAMES[j], "Product": n,
                   "Period": p, "Quantity": round(W[j, n, p].X, 4)}
                  for j in FACTORIES for n in PRODUCTS for p in PERIODS
                  if W[j, n, p].X > 1e-6]
        pd.DataFrame(rows_w).to_excel(writer, sheet_name="W_production",  index=False)

        # ── Sheet 3: Y – DC deliveries ────────────────────────────────────
        rows_y = [{"DC": DC_NAMES[l], "Product": n, "Period": p,
                   "Delivery": round(Y[l, n, p].X, 4)}
                  for l in DCS for n in PRODUCTS for p in PERIODS
                  if Y[l, n, p].X > 1e-6]
        pd.DataFrame(rows_y).to_excel(writer, sheet_name="Y_delivery",    index=False)

        # ── Sheet 4: B – Backlogs ─────────────────────────────────────────
        rows_b = [{"DC": DC_NAMES[l], "Product": n, "Period": p,
                   "Backlog": round(B[l, n, p].X, 4)}
                  for l in DCS for n in PRODUCTS for p in [0] + PERIODS
                  if B[l, n, p].X > 1e-6]
        pd.DataFrame(rows_b).to_excel(writer, sheet_name="B_backlog",     index=False)

        # ── Sheet 5: Q – Inventory levels ────────────────────────────────
        rows_q = [{"DC": DC_NAMES[l], "Product": n, "Period": p,
                   "Inventory": round(Q[l, n, p].X, 4)}
                  for l in DCS for n in PRODUCTS for p in [0] + PERIODS]
        pd.DataFrame(rows_q).to_excel(writer, sheet_name="Q_inventory",   index=False)

        # ── Sheet 6: z – DC open/close status ────────────────────────────
        rows_z = [{"DC": DC_NAMES[l], "Period": p, "Open": int(z[l, p].X + 0.5)}
                  for l in DCS for p in PERIODS]
        pd.DataFrame(rows_z).to_excel(writer, sheet_name="z_DC_status",   index=False)

        # ── Sheet 7: y – DC investment decisions ─────────────────────────
        rows_y2 = [{"DC": DC_NAMES[l], "Invest": int(y[l].X + 0.5)}
                   for l in DCS]
        pd.DataFrame(rows_y2).to_excel(writer, sheet_name="y_invest",     index=False)

        # ── Sheet 8: delta – DC switching events ─────────────────────────
        rows_d = [{"DC": DC_NAMES[l], "Period": p,
                   "Switch": int(delta[l, p].X + 0.5)}
                  for l in DCS for p in PERIODS
                  if delta[l, p].X > 0.5]
        pd.DataFrame(rows_d).to_excel(writer, sheet_name="delta_switch",  index=False)

        # ── Sheet 9: Cost breakdown by period ────────────────────────────
        rows_cost = []
        for p in PERIODS:
            trans_p = sum(
                params["trans_cost"].get((i, j, t, n), 0.0) * X[i, j, t, n, p].X
                for i in SUPPLIERS for j in FACTORIES
                for t in MODES for n in PRODUCTS
                if (i, j, t, n, p) in X
            )
            prod_p = sum(
                params["prod_cost"].get((j, n), 0.0) * W[j, n, p].X
                for j in FACTORIES for n in PRODUCTS
            )
            hold_p = sum(
                params["hold_cost"].get((l, n), 0.0) * Q[l, n, p].X
                for l in DCS for n in PRODUCTS
            )
            back_p = sum(
                params["back_cost"].get((l, n), 0.0) * B[l, n, p].X
                for l in DCS for n in PRODUCTS
            )
            rows_cost.append({
                "Period":       p,
                "Transport":    round(trans_p, 2),
                "Production":   round(prod_p,  2),
                "Holding":      round(hold_p,  2),
                "Backlog":      round(back_p,  2),
                "Total":        round(trans_p + prod_p + hold_p + back_p, 2),
            })
        # Investment & switching (one-time)
        invest_total = sum(params["dc_invest"].get(l, 0.0) * y[l].X for l in DCS)
        switch_total = sum(
            params["dc_switch"].get(l, 0.0) * delta[l, p].X
            for l in DCS for p in PERIODS
        )
        rows_cost.append({"Period": "DC_Invest",  "Total": round(invest_total, 2)})
        rows_cost.append({"Period": "DC_Switch",  "Total": round(switch_total, 2)})
        rows_cost.append({"Period": "GRAND TOTAL","Total": round(obj_val, 2)})
        pd.DataFrame(rows_cost).to_excel(writer, sheet_name="CostByPeriod", index=False)

        # ── Sheet 10: Product summary by period ──────────────────────────
        rows_prod = []
        for n in PRODUCTS:
            for p in PERIODS:
                total_prod = sum(W[j, n, p].X for j in FACTORIES)
                total_del  = sum(Y[l, n, p].X for l in DCS)
                total_inv  = sum(Q[l, n, p].X for l in DCS)
                total_bl   = sum(B[l, n, p].X for l in DCS)
                rows_prod.append({
                    "Product":    n,
                    "Period":     p,
                    "Production": round(total_prod, 4),
                    "Delivery":   round(total_del,  4),
                    "Inventory":  round(total_inv,  4),
                    "Backlog":    round(total_bl,   4),
                    "Demand":     params["demand"].get((n, p), 0.0),
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

    # 1. Load data
    params = load_parameters()

    # 2. Build and solve model
    print("\n  [INFO] Building model …")
    model, X, W, Y, B, Q, y, z, delta = build_model(params)

    print(f"  [INFO] Variables : {model.NumVars:,}")
    print(f"  [INFO] Constraints: {model.NumConstrs:,}")
    print(f"  [INFO] Binary vars: {model.NumBinVars:,}")
    print(f"  [INFO] Feasible X-triples (geographic filter): {len(X):,}")
    print("\n  [INFO] Solving …\n")

    model.optimize()

    # 3. Export results
    export_results(model, X, W, Y, B, Q, y, z, delta, params)

    print("\n  Done.\n")


if __name__ == "__main__":
    main()
