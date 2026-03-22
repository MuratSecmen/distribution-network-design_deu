# Distribution Network Design — Extended MILP Model
### Dokuz Eylül Üniversitesi | Lojistik Mühendisliği M.Sc.

![Solver](https://img.shields.io/badge/Solver-Gurobi-orange)
![Language](https://img.shields.io/badge/Language-Python-blue)
![Model](https://img.shields.io/badge/Model-MILP-green)
![University](https://img.shields.io/badge/University-Dokuz%20Eylül-red)

> Multi-period, multi-product Mixed-Integer Linear Programming (MILP) model  
> for strategic supply chain network design. Implemented in Python/Gurobi.

---

## Problem Tanımı

Bu çalışma, çok dönemli çok ürünlü bir tedarik zinciri ağında **dağıtım merkezi yatırım kararları**, **açma/kapama kararları** ve **çok modlu ürün akış optimizasyonu**nu eş zamanlı ele alan bir MILP modeli sunmaktadır.

### Ağ Yapısı
```
[Tedarikçiler]     [Fabrikalar]    [Dağıtım Mrk.]    [Müşteriler]

T1 (Rusya)  ──┐                   ┌── DM1 ──┐
T2 (Almanya)──┼──→ F1 (İspanya)──→├── DM2 ──┼──→ M1..M6
T3 (Norveç) ──┤                   └── DM3 ──┘
T4 (Çin)    ──┼──→ F2 (Rusya) ──→
T5 (Türkiye)──┘──→ F3 (Türkiye)──→
               └──→ F4 (Polonya)──→

Ulaşım modları: Kara · Demir · Hava · Deniz
Planlama ufku:  4 dönem (takvim çeyrekleri)
Ürün sayısı:    3
```

---

## Karar Değişkenleri

| Değişken | Açıklama |
|----------|----------|
| `x[i,j,m,p,r]` | Tedarikçi→Fabrika akışı (ürün r, mod m, dönem p) |
| `w[j,l,m,p,r]` | Fabrika→DM akışı |
| `y[l,k,m,p,r]` | DM→Müşteri akışı |
| `q[l,p,r]` | DM'de dönem sonu envanter düzeyi |
| `b[k,p,r]` | Müşteri k'nın karşılanamayan talebi (backlog) |
| `z[l,p]` | DM l'nin dönem p'de açık olup olmadığı (binary) |
| `u[l]` | DM l'ye yatırım yapılıp yapılmadığı (binary) |
| `delta[l,p]` | DM açma/kapama geçiş değişkeni (switching) |

---

## Amaç Fonksiyonu
```
Min Z = Taşıma Maliyeti
      + Üretim Maliyeti
      + Envanter Tutma Maliyeti
      + Backlog Ceza Maliyeti
      + DM Yatırım Maliyeti
      + DM Açma/Kapama Geçiş Maliyeti
```

---

## Orijinal Modele Kıyasla Geliştirmeler

| # | Boyut | Orijinal | Bu Çalışma |
|---|-------|----------|------------|
| 1 | Ürün sayısı | 1 | 3 |
| 2 | DM sayısı | 1 (sabit) | 3 (dinamik) |
| 3 | Ağ büyüklüğü | 3T/2F/1DM/3M | 5T/4F/3DM/6M |
| 4 | DM kararı | Sadece açık/kapalı | Yatırım + açık/kapalı |
| 5 | Ulaşım modu | Kısıtsız | Coğrafi fizibilite filtreli |
| 6 | Parametre mimarisi | Kısmen hardcoded | %100 Excel-driven |
| 7 | Dönem tanımı | Soyut | Takvim çeyrekleri |

---

## Veri Dosyaları

Tüm parametreler Excel'den okunur. `main.py` içinde hiçbir sayısal değer hardcoded değildir.

### `parameters.xlsx` — 9 Sayfa

| Sayfa | İçerik |
|-------|--------|
| `SCALAR_PARAMS` | Big-M, MIP gap, zaman limiti |
| `DEMAND` | Müşteri talebi (ürün × dönem) |
| `PROD_COST` | Fabrika üretim maliyeti |
| `FACTORY_CAP` | Fabrika kapasitesi |
| `DC_CAP` | DM depolama kapasitesi |
| `DC_INVEST` | DM yatırım maliyeti |
| `DC_SWITCH` | DM açma/kapama maliyeti |
| `HOLD_COST` | Birim envanter tutma maliyeti |
| `BACK_COST` | Birim backlog ceza maliyeti |

### `transportation_costs.xlsx`

180 satırlık maliyet matrisi. Sütunlar: `i, j, l, k, mode, product, cost_per_unit`  
Yalnızca coğrafi olarak fizibil `(düğüm, mod)` kombinasyonları içerir.

---

## Kurulum ve Çalıştırma

### Gereksinimler
```bash
pip install gurobipy pandas openpyxl
```

> Geçerli bir Gurobi lisansı gerekmektedir.  
> Akademik lisans için: [gurobi.com/academia](https://www.gurobi.com/academia/academic-program-and-licenses/)

### Çalıştırma
```bash
cd extended_model
python main.py
```

---

## Repo Yapısı
```
distribution-network-design_deu/
├── extended_model/
│   ├── main.py                     # Model + solver
│   ├── parameters.xlsx             # 9 sayfalık parametre seti
│   └── transportation_costs.xlsx   # Maliyet matrisi
├── results/
│   ├── milp_solution.xlsx          # MILP çözüm detayları
│   └── optimal_solution.xlsx       # Optimal çözüm özeti
├── LICENSE                         # MIT
└── README.md
```

---


[![LinkedIn](https://img.shields.io/badge/LinkedIn-muratsecmen-blue?logo=linkedin)](https://linkedin.com/in/muratsecmen)
[![GitHub](https://img.shields.io/badge/GitHub-MuratSecmen-black?logo=github)](https://github.com/MuratSecmen)
