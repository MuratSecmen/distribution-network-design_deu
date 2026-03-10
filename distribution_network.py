"""
Distribution Network Design and Optimization in a Supply Chain
==============================================================
Published : Beykoz Akademi Dergisi, 2015; 3(1), 67-84
Authors   : Murat Secmen, Temel Oncan, Okan Tuna
MSc Thesis: Dokuz Eylul University, Logistics Engineering (2009-2013)
Solver    : Gurobi Python API (gurobipy)

Problem:
    Multi-period, multi-modal supply chain network optimization.
    Minimize: transportation + opportunity + inventory holding + stockout costs.

Network:
    Suppliers  : T1=Russia(i=1), T2=Germany(i=2), T3=Norway(i=3)
    Factories  : F1=Spain(j=1),  F2=Russia(j=2)
    Dist. Ctr  : Ukraine (single DC)
    Markets    : M1=Finland(k=1), M2=Turkey(k=2), M3=Russia(k=3)
    Modes      : t=1 Rail, t=2 Road, t=3 Sea
    Periods    : p = 1..6 (monthly, 6-month planning horizon)

Variable Notation (thesis):
    X_ijtp : units from supplier i to factory j via mode t in period p
    W_jp   : units from factory j to DC in period p
    Y_kp   : units from DC to customer k in period p
    B_kp   : backlogged (unmet) demand of customer k in period p
    Q_p    : DC inventory level at end of period p

Optimal Solution: Min Z = 922,396.5
"""

import gurobipy as gp
from gurobipy import GRB
import pandas as pd
import os

os.makedirs("results", exist_ok=True)

# ── INDEX SETS ────────────────────────────────────────────────────────────────
suppliers = [1, 2, 3]
factories = [1, 2]
modes     = [1, 2, 3]   # 1=Rail, 2=Road, 3=Sea
periods   = [1, 2, 3, 4, 5, 6]
customers = [1, 2, 3]

# ── PARAMETERS — Transportation cost C_ijtp (Table 3.2 Rail, 3.3 Road, 3.4 Sea)
# C[(i,j,t,p)] = unit cost
C = {}
# Rail (t=1) costs — Table 3.2
rail = {
    (1,1): [3,4,5,6,5,5],  (1,2): [7,6,3,2,2,3],
    (2,1): [6,7,7,8,7,7],  (2,2): [3,5,9,6,6,9],
    (3,1): [4,6,3,6,3,3],  (3,2): [6,6,7,5,5,7],
}
# Road (t=2) costs — Table 3.3
road = {
    (1,1): [2,2,3,2,3,3],  (1,2): [3,1,3,2,2,3],
    (2,1): [3,3,3,2,3,3],  (2,2): [1,3,3,4,4,3],
    (3,1): [3,4,3,2,3,3],  (3,2): [5,4,3,1,1,3],
}
# Sea (t=3) costs — Table 3.4
sea = {
    (1,1): [5,9,8,12,9,5],   (1,2): [13,11,13,11,7,8],
    (2,1): [8,10,9,14,8,7],  (2,2): [12,14,11,14,9,9],
    (3,1): [7,11,7,15,6,8],  (3,2): [9,13,9,11,8,7],
}
for (i,j), vals in rail.items():
    for pi, v in enumerate(vals, 1): C[(i,j,1,pi)] = v
for (i,j), vals in road.items():
    for pi, v in enumerate(vals, 1): C[(i,j,2,pi)] = v
for (i,j), vals in sea.items():
    for pi, v in enumerate(vals, 1): C[(i,j,3,pi)] = v

# Opportunity loss N_ijtp (Table 3.1, second value in "time-loss" pairs)
# Format from thesis: each cell "beta-N" where beta=transit time, N=opp.loss
# Rail opportunity losses:
N_rail = {
    (1,1): [48,72,96,120,110,145],  (1,2): [70,114,172,104,125,100],
    (2,1): [70,114,172,104,125,100],(2,2): [80,75,140,112,70,56],
    (3,1): [100,108,100,170,144,130],(3,2): [100,75,140,112,70,56],
}
N_road = {
    (i,j): [0]*6 for i in [1,2,3] for j in [1,2]
}  # Road = reference mode, opportunity loss = 0
N_sea = {
    (1,1): [192,240,288,384,350,395], (1,2): [180,224,302,334,360,400],
    (2,1): [180,224,302,334,360,400], (2,2): [195,250,230,422,302,211],
    (3,1): [180,220,290,300,432,200], (3,2): [190,227,320,390,160,195],
}
N = {}
for (i,j), vals in N_rail.items():
    for pi, v in enumerate(vals, 1): N[(i,j,1,pi)] = v
for (i,j), vals in N_road.items():
    for pi, v in enumerate(vals, 1): N[(i,j,2,pi)] = v
for (i,j), vals in N_sea.items():
    for pi, v in enumerate(vals, 1): N[(i,j,3,pi)] = v

pi_rate = 1.0  # opportunity cost rate

# Factory -> DC cost Cw_jp (Table 3.5, rows Fab1/Fab2)
Cw = {
    (1,1):20,(1,2):30,(1,3):20,(1,4):30,(1,5):40,(1,6):20,
    (2,1):30,(2,2):20,(2,3):40,(2,4):20,(2,5):30,(2,6):20,
}

# DC -> Customer cost Cy_kp (Table 3.5, rows Must1/2/3)
Cy = {
    (1,1):40,(1,2):20,(1,3):30,(1,4):30,(1,5):40,(1,6):20,
    (2,1):30,(2,2):40,(2,3):30,(2,4):20,(2,5):40,(2,6):20,
    (3,1):30,(3,2):20,(3,3):40,(3,4):20,(3,5):30,(3,6):20,
}

# Holding cost R_p and Stockout cost T_p (Table 3.5)
R = {1:20, 2:25, 3:20, 4:10, 5:20, 6:20}
T = {1:20, 2:20, 3:20, 4:30, 5:25, 6:15}

# Supplier capacity alpha_ip (Table 3.6)
alpha = {
    (1,1):10000,(1,2):25000,(1,3):20000,(1,4):10000,(1,5):10000,(1,6):10000,
    (2,1):15000,(2,2):30000,(2,3):15000,(2,4):10000,(2,5):15000,(2,6):10000,
    (3,1):25000,(3,2):20000,(3,3):30000,(3,4):20000,(3,5):25000,(3,6):10000,
}

# Factory capacity b_jp (Table 3.6)
b_cap = {
    (1,1):40000,(1,2):30000,(1,3):34000,(1,4):28000,(1,5):40000,(1,6):36000,
    (2,1):40000,(2,2):20000,(2,3):30000,(2,4):31000,(2,5):38000,(2,6):40000,
}

# DC capacity c_p (Table 3.6)
c_cap = {p: 40000 for p in periods}

# Mode capacity A_tp (Table 3.6)
A_mode = {(1,p):55000 for p in periods}
A_mode.update({(2,p):35000 for p in periods})
A_mode.update({(3,p):75000 for p in periods})

# Customer demand d_kp (Table 3.6)
d = {
    (1,1):10000,(1,2):10000,(1,3):10000,(1,4):10000,(1,5):10000,(1,6):10000,
    (2,1):15000,(2,2):10000,(2,3):10000,(2,4):15000,(2,5):15000,(2,6):15000,
    (3,1):15000,(3,2):20000,(3,3):15000,(3,4):15000,(3,5):15000,(3,6):15000,
}

# Lead times: DC -> customer G_k = 2 periods (from thesis)
G = {1:2, 2:2, 3:2}
Q_init = 40000

# ── MODEL ─────────────────────────────────────────────────────────────────────
m = gp.Model("DistributionNetworkDesign_Secmen2015")
m.setParam("OutputFlag", 1)

# ── DECISION VARIABLES ────────────────────────────────────────────────────────
X = {(i,j,t,p): m.addVar(lb=0, name=f"X_{i}{j}{t}{p}")
     for i in suppliers for j in factories for t in modes for p in periods}

W = {(j,p): m.addVar(lb=0, name=f"W_{j}{p}")
     for j in factories for p in periods}

Y = {(k,p): m.addVar(lb=0, name=f"Y_{k}{p}")
     for k in customers for p in periods}

B = {(k,p): m.addVar(lb=0, name=f"B_{k}{p}")
     for k in customers for p in [0]+periods}

Q = {p: m.addVar(lb=0, name=f"Q_{p}") for p in [0]+periods}

m.update()

# ── OBJECTIVE FUNCTION ────────────────────────────────────────────────────────
# (1) Transportation cost
obj_trans = gp.quicksum(C[i,j,t,p] * X[i,j,t,p]
                        for i in suppliers for j in factories
                        for t in modes for p in periods)

# (2) Factory→DC + DC→Customer costs
obj_w = gp.quicksum(Cw[j,p] * W[j,p] for j in factories for p in periods)
obj_y = gp.quicksum(Cy[k,p] * Y[k,p] for k in customers for p in periods)

# (3) Opportunity cost
obj_opp = gp.quicksum(N.get((i,j,t,p),0) * pi_rate * X[i,j,t,p]
                      for i in suppliers for j in factories
                      for t in modes for p in periods)

# (4) Inventory holding cost (period-varying R_p)
obj_hold = gp.quicksum(
    R[p] * (gp.quicksum(W[j,p] for j in factories)
            - gp.quicksum(Y[k,p] for k in customers)
            - gp.quicksum(B[k,p] for k in customers)
            + Q[p-1])
    for p in periods)

# (5) Stockout/backlog cost (period-varying T_p)
obj_stock = gp.quicksum(
    T[p] * (gp.quicksum(Y[k,p] for k in customers)
            + gp.quicksum(B[k,p] for k in customers)
            - gp.quicksum(W[j,p] for j in factories)
            - Q[p-1])
    for p in periods)

m.setObjective(obj_trans + obj_w + obj_y + obj_opp + obj_hold + obj_stock,
               GRB.MINIMIZE)

# ── CONSTRAINTS ──────────────────────────────────────────────────────────────
# C1: Supplier capacity (3.1)
for i in suppliers:
    for p in periods:
        m.addConstr(gp.quicksum(X[i,j,t,p] for j in factories for t in modes)
                    <= alpha[i,p], name=f"SC_{i}_{p}")

# C2: Factory capacity (3.2)
for j in factories:
    for p in periods:
        m.addConstr(W[j,p] <= b_cap[j,p], name=f"FC_{j}_{p}")

# C3: DC capacity (3.3)
for p in periods:
    m.addConstr(gp.quicksum(Y[k,p] for k in customers)
                <= c_cap[p], name=f"DC_{p}")

# C4: Customer demand satisfaction with backlog (3.4)
for k in customers:
    for p in periods:
        m.addConstr(Y[k,p] + B[k,p-1] + B[k,p] >= d[k,p],
                    name=f"Dem_{k}_{p}")

# C5: Mode capacity (3.5)
for t in modes:
    for p in periods:
        m.addConstr(gp.quicksum(X[i,j,t,p] for i in suppliers for j in factories)
                    <= A_mode[t,p], name=f"MC_{t}_{p}")

# C6: Flow balance — supplier inflow to factory = W_jp
for j in factories:
    for p in periods:
        m.addConstr(gp.quicksum(X[i,j,t,p] for i in suppliers for t in modes)
                    == W[j,p], name=f"FB_{j}_{p}")

# C7: DC inventory balance
m.addConstr(Q[0] == Q_init, name="Q_init")
for p in periods:
    m.addConstr(Q[p] == Q[p-1]
                + gp.quicksum(W[j,p] for j in factories)
                - gp.quicksum(Y[k,p] for k in customers),
                name=f"IB_{p}")

# C8: Initial backlog = 0
for k in customers:
    m.addConstr(B[k,0] == 0, name=f"B0_{k}")

# C9: DC inventory upper bound per period
for p in periods:
    m.addConstr(Q[p] <= 40000, name=f"Qub_{p}")

# ── SOLVE ─────────────────────────────────────────────────────────────────────
m.optimize()

# ── OUTPUT ────────────────────────────────────────────────────────────────────
if m.status == GRB.OPTIMAL:
    print(f"\n{'='*60}")
    print(f"  OPTIMAL SOLUTION — Min Z = {m.ObjVal:,.1f}")
    print(f"  (Thesis reference: 922,396.5)")
    print(f"{'='*60}")

    snames = {1:"Russia(T1)", 2:"Germany(T2)", 3:"Norway(T3)"}
    fnames = {1:"Spain(F1)",  2:"Russia(F2)"}
    mnames = {1:"Rail", 2:"Road", 3:"Sea"}
    cnames = {1:"Finland(M1)", 2:"Turkey(M2)", 3:"Russia(M3)"}

    rows = []
    for i in suppliers:
        for j in factories:
            for t in modes:
                for p in periods:
                    v = X[i,j,t,p].X
                    if v > 0.5:
                        print(f"  X[{i},{j},{t},{p}] = {v:>8,.0f}  "
                              f"{snames[i]} -> {fnames[j]} via {mnames[t]}, P{p}")
                        rows.append({"Var":f"X_{i}{j}{t}{p}","Value":v,
                                     "From":snames[i],"To":fnames[j],
                                     "Mode":mnames[t],"Period":p})
    for j in factories:
        for p in periods:
            v = W[j,p].X
            if v > 0.5:
                print(f"  W[{j},{p}]   = {v:>8,.0f}  {fnames[j]} -> Ukraine DC, P{p}")
                rows.append({"Var":f"W_{j}{p}","Value":v,
                             "From":fnames[j],"To":"Ukraine DC",
                             "Mode":"—","Period":p})
    for k in customers:
        for p in periods:
            v = Y[k,p].X
            if v > 0.5:
                print(f"  Y[{k},{p}]   = {v:>8,.0f}  DC -> {cnames[k]}, P{p}")
                rows.append({"Var":f"Y_{k}{p}","Value":v,
                             "From":"Ukraine DC","To":cnames[k],
                             "Mode":"—","Period":p})
    for p in [0]+periods:
        v = Q[p].X
        if v > 0.5:
            print(f"  Q[{p}]     = {v:>8,.0f}  DC Inventory")
            rows.append({"Var":f"Q_{p}","Value":v,
                         "From":"DC Stock","To":"—","Mode":"—","Period":p})
    for k in customers:
        for p in [0]+periods:
            v = B[k,p].X
            if v > 0.5:
                print(f"  B[{k},{p}]   = {v:>8,.0f}  Backlog {cnames[k]}")
                rows.append({"Var":f"B_{k}{p}","Value":v,
                             "From":"Backlog","To":cnames[k],
                             "Mode":"—","Period":p})

    pd.DataFrame(rows).to_excel("results/optimal_solution_gurobi.xlsx", index=False)
    print("\nExported: results/optimal_solution_gurobi.xlsx")
else:
    print(f"Status: {m.status}")
