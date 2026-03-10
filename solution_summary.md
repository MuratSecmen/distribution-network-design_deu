# Solution Summary

## Optimal Objective Value

**Min. Z = 922,396.5**

Solved with IBM CPLEX Optimization Studio.

## Optimal Flow Variables (Nonzero)

### Supplier → Factory (X_ijtp)

| Variable | Supplier | Factory | Mode | Period | Quantity |
|----------|----------|---------|------|--------|----------|
| X1111 | Russia (T1) | Spain (F1) | Rail | 1 | 5,000 |
| X1112 | Russia (T1) | Spain (F1) | Rail | 2 | 5,000 |
| X1113 | Russia (T1) | Spain (F1) | Rail | 3 | 5,000 |
| X1121 | Russia (T1) | Spain (F1) | Road | 1 | 5,000 |
| X1123 | Russia (T1) | Spain (F1) | Road | 3 | 15,000 |
| X2121 | Germany (T2) | Spain (F1) | Road | 1 | 5,000 |
| X2222 | Germany (T2) | Russia (F2) | Road | 2 | 20,000 |
| X3121 | Norway (T3) | Spain (F1) | Road | 1 | 25,000 |
| X3122 | Norway (T3) | Spain (F1) | Road | 2 | 15,000 |
| X3123 | Norway (T3) | Spain (F1) | Road | 3 | 14,000 |
| X3223 | Norway (T3) | Russia (F2) | Road | 3 | 6,000 |

### Factory → Distribution Center (W_jp)

| Variable | Factory | Period | Quantity |
|----------|---------|--------|----------|
| W11 | Spain (F1) | 1 | 40,000 |
| W12 | Spain (F1) | 2 | 20,000 |
| W13 | Spain (F1) | 3 | 34,000 |
| W22 | Russia (F2) | 2 | 20,000 |
| W23 | Russia (F2) | 3 | 6,000 |

### Distribution Center → Customers (Y_kp)

| Variable | Customer | Period | Quantity |
|----------|----------|--------|----------|
| Y12 | Finland (M1) | 2 | 10,000 |
| Y13 | Finland (M1) | 3 | 15,000 |
| Y14 | Finland (M1) | 4 | 10,000 |
| Y15 | Finland (M1) | 5 | 10,000 |
| Y16 | Finland (M1) | 6 | 10,000 |
| Y22 | Turkey (M2) | 2 | 10,000 |
| Y23 | Turkey (M2) | 3 | 10,000 |
| Y24 | Turkey (M2) | 4 | 15,000 |
| Y25 | Turkey (M2) | 5 | 15,000 |
| Y26 | Turkey (M2) | 6 | 15,000 |
| Y32 | Russia (M3) | 2 | 20,000 |
| Y33 | Russia (M3) | 3 | 15,000 |
| Y34 | Russia (M3) | 4 | 15,000 |
| Y35 | Russia (M3) | 5 | 15,000 |
| Y36 | Russia (M3) | 6 | 15,000 |

### Initial Conditions

| Variable | Value | Description |
|----------|-------|-------------|
| Q0 | 40,000 | Initial DC inventory |
| B10 | 10,000 | Finland unmet demand P1 |
| B20 | 15,000 | Turkey unmet demand P1 |
| B30 | 15,000 | Russia unmet demand P1 |

## Interpretation

- **Period 1 stockouts** occur for all markets because the DC-to-customer lead time is 2 periods. Initial inventory (Q0 = 40,000) covers Period 2 demand entirely.
- **Norway (T3)** is the dominant supplier due to its low unit transport cost (1.35–2.00) despite higher lead time opportunity cost.
- **Spain (F1)** handles the majority of production and DC supply.
- From Period 4 onwards the supply chain operates in steady state with no backlog.
