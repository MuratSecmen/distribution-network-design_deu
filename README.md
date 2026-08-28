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
| `x[i,j,m,p,r]` | Tedarikçi→Fabrika akışı (ürün r, mod m, dönem p) — çok modlu |
| `w[j,l,p,r]` | Fabrika→DM akışı (mod ayrımı yok) |
| `y[l,k,p,r]` | DM→Müşteri akışı (mod ayrımı yok) |
| `q[l,p,r]` | DM'de dönem sonu envanter düzeyi |
| `b[k,p,r]` | Müşteri k'nın karşılanamayan talebi (backlog) |
| `z[l,p]` | DM l'nin dönem p'de açık olup olmadığı (binary) |
| `u[l]` | DM l'ye yatırım yapılıp yapılmadığı (binary, tek seferlik — dönem indeksi yok) |
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
| 6 | Parametre mimarisi | Kısmen hardcoded | Maliyet/kapasite/talep parametreleri Excel-driven (küme tanımları ve mod-fizibilite tablosu kodda sabit) |
| 7 | Dönem tanımı | Soyut | Takvim çeyrekleri |

---

## Veri Dosyaları

Maliyet/kapasite/talep parametrelerinin tamamı Excel'den okunur. Küme
tanımları (tedarikçi/fabrika/DM/müşteri isimleri, ürün ve dönem sayısı) ve
coğrafi taşıma-modu uygunluk tablosu (`FEASIBLE_MODES`) ise `main.py` /
`main_v2.py` içinde sabit kodlu — bunlar Excel'den okunmaz.

### `extended_model/data/parameters.xlsx` — `main.py` için, 14 Sayfa

| Sayfa | İçerik |
|-------|--------|
| `SCALAR_PARAMS` | Big-M, MIP gap, zaman limiti |
| `DEMAND` | Müşteri talebi (müşteri × ürün × dönem) |
| `SUPPLIER_CAP` | Tedarikçi kapasitesi |
| `FACTORY_CAP` | Fabrika üretim kapasitesi |
| `MODE_CAP` | Taşıma modu kapasitesi |
| `DC_THROUGHPUT` | DM işlem (throughput) kapasitesi |
| `DC_STORAGE` | DM depolama kapasitesi |
| `DC_INVEST` | DM yatırım maliyeti (dönem 1'deki değer tek seferlik yatırım maliyeti olarak kullanılır) |
| `DC_OPCOST` | DM değişken işletme maliyeti |
| `DC_SWITCH` | DM açma/kapama geçiş maliyeti |
| `HOLD_COST` | Birim envanter tutma maliyeti |
| `BACK_COST` | Birim backlog ceza maliyeti |
| `FACTORY_DC_COST` | Fabrika→DM birim taşıma maliyeti |
| `DC_CUST_COST` | DM→Müşteri birim taşıma maliyeti |

### `extended_model/data/transportation_costs.xlsx` — `main.py` için

`TRANS_COST` sayfası, sütunlar: `i, j, m, n, cost`. Yalnızca coğrafi olarak
fizibil `(tedarikçi, fabrika, mod)` kombinasyonları içerir.

### `parameters_v2.xlsx` / `transportation_costs_v2.xlsx` — `main_v2.py` için

Genişletilmiş kardinaliteler (10T/8F/6DM/20M), Haversine tabanlı mesafe ve
mod-faktörü ile hesaplanan taşıma maliyeti/emisyonu, fabrika açma/kapama
kararı (`phi`/`gamma`), karbon emisyon bütçesi ve minimum hizmet düzeyi
kısıtları için ek sayfalar (`FAC_INVEST`, `FAC_SWITCH`, `MODE_FACTORS`)
içerir.

---

## Kurulum ve Çalıştırma

### Gereksinimler
```bash
pip install -r requirements.txt
```

> Geçerli bir Gurobi lisansı gerekmektedir.  
> Akademik lisans için: [gurobi.com/academia](https://www.gurobi.com/academia/academic-program-and-licenses/)
>
> `main_v2.py` (10T/8F/6DM/20M, 5 ürün, 12 dönem) boyutu, Gurobi'nin
> size-limited (ücretsiz/kısıtlı) lisansının değişken sınırını aşar —
> tam lisans gerektirir. `main.py` küçük ölçekli olduğu için kısıtlı
> lisansla da çalışır.

### Çalıştırma
```bash
cd extended_model
python main.py       # temel genişletilmiş model (3T/4F/3DM/6M)
python main_v2.py     # genişletilmiş kardinalite + emisyon/hizmet-düzeyi kısıtları
```

---

## Repo Yapısı
```
distribution-network-design_deu/
├── extended_model/
│   ├── main.py                        # temel genişletilmiş model
│   ├── main_v2.py                     # genişletilmiş kardinalite + emisyon/hizmet-düzeyi
│   ├── data/
│   │   ├── parameters.xlsx            # main.py için
│   │   ├── transportation_costs.xlsx  # main.py için
│   │   ├── parameters_v2.xlsx         # main_v2.py için
│   │   └── transportation_costs_v2.xlsx
│   └── output/                        # çalıştırma çıktıları (git'e commit edilmez)
├── requirements.txt
├── LICENSE
└── README.md
```

---


[![LinkedIn](https://img.shields.io/badge/LinkedIn-muratsecmen-blue?logo=linkedin)](https://linkedin.com/in/muratsecmen)
[![GitHub](https://img.shields.io/badge/GitHub-MuratSecmen-black?logo=github)](https://github.com/MuratSecmen)
