# Distribution Network Design and Optimization in a Supply Chain

[![Published](https://img.shields.io/badge/Published-Beykoz%20Akademi%20Dergisi%202015-blue)](https://dergipark.org.tr/en/download/article-file/949560)
[![Language](https://img.shields.io/badge/Solver-IBM%20CPLEX%20%7C%20Gurobi-orange)]()
[![Model](https://img.shields.io/badge/Model-Linear%20Programming-green)]()
[![University](https://img.shields.io/badge/University-Dokuz%20Eyl%C3%BCl%20University-red)]()

> **Published in:** Beykoz Akademi Dergisi, 2015; 3(1), 67–84  
> **Authors:** Murat Seçmen, Temel Öncan, Okan Tuna  
> **Thesis:** M.Sc. in Logistics Engineering, Dokuz Eylül University (2009–2013)

---

## Problem Description

This study addresses the **multi-period, multi-modal distribution network design** problem for a retail supply chain operating across international markets. The objective is to determine optimal product flow quantities across the network while minimizing total supply chain costs.

### Network Structure

```
[Suppliers]          [Factories]     [Distribution]    [Markets]
                                        Center
Russia (T1) ──┐                    ┌─────────────┐──→ Finland (M1)
              ├──→ Spain (F1) ──→  │  Ukraine    │──→ Türkiye  (M2)
Germany (T2)──┤                    │    (DM)     │──→ Russia  (M3)
              └──→ Russia (F2) ──→ └─────────────┘
Norway (T3) ──┘

Transportation modes: Rail · Road · Air  |  Planning horizon: 6 periods (months)
```

### Decision Variables

| Variable | Description |
|----------|-------------|
| `X_ijtp` | Units shipped from supplier `i` to factory `j` via mode `t` in period `p` |
| `W_jp`   | Units shipped from factory `j` to distribution center in period `p` |
| `Y_kp`   | Units shipped from distribution center to customer `k` in period `p` |
| `B_kp`   | Backlogged (unmet) demand of customer `k` in period `p` |
| `Q_p`    | Inventory level at distribution center at end of period `p` |

### Objective Function

Minimize total cost across all periods:

```
Min Z = Transportation Cost
      + Opportunity Cost (delay penalty π × lead time N_ijtp × X_ijtp)
      + Inventory Holding Cost (R × excess stock at DC)
      + Stockout/Backlog Cost (T × unmet demand)
```

**Optimal Solution: Z = 922,396.5**

### Key Constraints

1. **Supplier capacity:** Total shipments from each supplier ≤ capacity `α_ip`
2. **Factory capacity:** Shipments from factory to DC ≤ factory capacity `b_jp`
3. **DC capacity:** Total outflow from DC ≤ DC capacity `c_p`
4. **Customer demand satisfaction:** Deliveries + backlog coverage ≥ demand `d_kp`
5. **Transport mode capacity:** Shipments per mode ≤ mode capacity `A_tp`

---

## Repository Structure

```
📁 distribution-network-design/
├── 📄 README.md
│
├── 📁 model/
│   ├── distribution_network.py
│   ├── distribution_network.mod
│   └── distribution_network.dat
│
├── 📁 data/
│   ├── parameters.xlsx
│   ├── transportation_costs.xlsx
│   └── optimal_solution.xlsx
│
├── 📁 results/
│   └── solution_summary.md
│
└── 📁 docs/
    └── paper.pdf


## Solvers & Dependencies

### Gurobi (Python)
```bash
pip install gurobipy pandas openpyxl
python model/distribution_network.py
```
> Requires a valid Gurobi license. Free academic licenses available at [gurobi.com](https://www.gurobi.com/academia/academic-program-and-licenses/).

### IBM CPLEX OPL
Open `model/distribution_network.mod` and `model/distribution_network.dat`  
in IBM ILOG CPLEX Optimization Studio and run.

---

## Results Summary

| Variable | Value | Description |
|----------|-------|-------------|
| **Min. Z** | **922,396.5** | Total minimum cost |
| X1111 | 5,000 | Russia→Spain via Rail, Period 1 |
| X1112 | 5,000 | Russia→Spain via Rail, Period 2 |
| X3121 | 25,000 | Norway→Spain via Road, Period 1 |
| X3122 | 15,000 | Norway→Spain via Road, Period 2 |
| W11 | 40,000 | Spain→Ukraine, Period 1 |
| W12 | 20,000 | Spain→Ukraine, Period 2 |
| Q0 | 40,000 | Initial DC inventory |

**Key finding:** Under current transportation costs, sourcing primarily from Norway and Russia (supported by Spanish production) yields the minimum-cost distribution plan. The 2-month lead time from the DC to markets necessitates a 40,000-unit initial buffer stock to prevent stockouts in Period 1.

---

## Citation

```bibtex
@article{secmen2015distribution,
  title   = {Tedarik Zincirinde Dağıtım Ağları Tasarımı Üzerine Bir Uygulama},
  author  = {Seçmen, Murat and Öncan, Temel and Tuna, Okan},
  journal = {Beykoz Akademi Dergisi},
  volume  = {3},
  number  = {1},
  pages   = {67--84},
  year    = {2015}
}
```

---

## Author

**Murat Seçmen**  
M.Sc. Logistics Engineering — Dokuz Eylül University  
Production Planning & Control Lead Engineer — Turkish Aerospace Industries (TUSAŞ)  
M.Sc. Candidate — Industrial Engineering, Hacettepe University  

[![LinkedIn](https://img.shields.io/badge/LinkedIn-muratsecmen-blue?logo=linkedin)](https://linkedin.com/in/muratsecmen)
[![GitHub](https://img.shields.io/badge/GitHub-MuratSecmen-black?logo=github)](https://github.com/MuratSecmen)
