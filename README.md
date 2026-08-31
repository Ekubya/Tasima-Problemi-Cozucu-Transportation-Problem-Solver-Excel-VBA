##ENGLISH
# Transportation Problem Solver — Excel + VBA

This project is an Excel (`.xlsm`) model supported by VBA macros that solves a **transportation (distribution) problem** using the least cost approach.

> When you open the file, if you see an **"Enable Content"** warning at the top of Excel, you must approve it for the macros to work.

## 📊 Problem Definition

The model is a classic **transportation problem** that aims to minimize the transportation cost between **5 sources (Factories: F1–F5)** and **6 destinations (Warehouses/Regions: D1–D6)**.

- **Unit Transportation Cost Table:** Unit cost for each source-destination pair (B3:G7)
- **Capacity:** Maximum supply amount of each source (H3:H7)
- **Demand:** Required amount for each destination (B8:G8)
- **Objective:** To find a distribution plan that minimizes the total transportation cost

In this example, total capacity = total demand = **3,650** units (balanced transportation problem).

## 🧮 Named Ranges Used

The VBA codes work through the following named ranges instead of using cell addresses directly:

| Name | Range | Meaning |
|---|---|---|
| `MyCost` | `B3:G7` | Unit transportation cost matrix |
| `Kapasite` | `H3:H7` | Capacities of the sources |
| `Talep` | `B8:G8` | Demands of the destinations |
| `Miktar` | `B13:G17` | Solution/distribution plan (filled by macros) |
| `Used` | `H13:H17` | Total amount used from each source |
| `Meet` | `B18:G18` | Total amount met for each destination |
| `Demand` | `B20:G20` | Demand check row |
| `ToplamMaliyet` | `H19` | Calculated total cost |

## ⚙️ VBA Macros

The workbook contains multiple macros that heuristically solve the transportation problem using the **Least Cost Method**. Each applies the same logic with a different loop structure:

| Module | Macro | Approach |
|---|---|---|
| `Module1` | `MiktarAta1` | Row-based least cost assignment — `For` loop |
| `Module2` | `MiktarAta2` | Same logic — `Do While` loop |
| `Module2` | `MiktarAta3` | Same logic — `Do Until` loop |
| `Module2` | `MiktarAta4` | **Column-based** least cost assignment |
| `Module3` | `MiktarAta6` | Clean, optimized version with type definitions (`Dim`) and `Range` objects |

**Working logic (summary):**
1. The `Miktar` (Amount) range is cleared.
2. For each source (row), the destination with the lowest cost that can utilize the remaining capacity is selected.
3. The assigned amount is the minimum of the remaining capacity and the remaining demand.
4. The process is repeated until the total used amount equals the total capacity.

### Running the macros
You can press `Alt + F8` in Excel, select the desired macro (e.g., `MiktarAta1`, `MiktarAta4`, `MiktarAta6`), and click **Run**. Since different macros use different sorting logics, they may produce **different (but valid) distribution plans** depending on the execution order; the total costs can be compared.

## ✅ Example Result

The distribution plan obtained when `MiktarAta1` is run:

| | D1 | D2 | D3 | D4 | D5 | D6 | Used |
|---|---|---|---|---|---|---|---|
| **F1** | 60 | – | – | – | 590 | – | 650 |
| **F2** | – | – | – | 550 | – | – | 550 |
| **F3** | – | – | 600 | – | – | – | 600 |
| **F4** | 650 | – | – | – | – | 200 | 850 |
| **F5** | 100 | 700 | 50 | 150 | – | – | 1,000 |
| **Met** | 810 | 700 | 650 | 700 | 590 | 200 | |

**Total Cost = 180,440**

> Note: This is an **initial feasible solution** provided by the least cost method (heuristic); it may not be the same as the **exact optimal solution** found through linear programming (e.g., Excel Solver, Simplex method).

## 🛠️ Requirements

- Microsoft Excel (desktop version — macros may not work in mobile/web Excel)
- Macros must be enabled (File > Options > Trust Center > Macro Settings)
##TURKISH
# Taşıma Problemi Çözücü (Transportation Problem Solver) — Excel + VBA

Bu proje, bir **taşıma (dağıtım) problemini** en küçük maliyet yaklaşımıyla çözen, VBA makroları ile desteklenmiş bir Excel (`.xlsm`) modelidir. 

> Dosyayı açtığınızda Excel'in üst tarafında **"İçeriği Etkinleştir / Enable Content"** uyarısı çıkarsa, makroların çalışması için bunu onaylamanız gerekir.

## 📊 Problem Tanımı

Model, **5 kaynak (Fabrika: F1–F5)** ile **6 hedef (Depo/Bölge: D1–D6)** arasındaki taşıma maliyetini minimize etmeyi amaçlayan klasik bir **taşıma problemi**dir.

- **Birim Taşıma Maliyeti Tablosu:** Her kaynak-hedef çifti için birim maliyet (B3:G7)
- **Kapasite:** Her kaynağın maksimum arz miktarı (H3:H7)
- **Talep:** Her hedefin ihtiyaç duyduğu miktar (B8:G8)
- **Amaç:** Toplam taşıma maliyetini minimize eden bir dağıtım planı bulmak

Bu örnekte toplam kapasite = toplam talep = **3.650** birim (dengeli taşıma problemi).

## 🧮 Kullanılan İsimlendirilmiş Aralıklar (Named Ranges)

VBA kodları, hücre adreslerini doğrudan kullanmak yerine aşağıdaki isimlendirilmiş aralıklar üzerinden çalışır:

| İsim | Aralık | Anlamı |
|---|---|---|
| `MyCost` | `B3:G7` | Birim taşıma maliyetleri matrisi |
| `Kapasite` | `H3:H7` | Kaynakların kapasiteleri |
| `Talep` | `B8:G8` | Hedeflerin talepleri |
| `Miktar` | `B13:G17` | Çözüm/dağıtım planı (makrolar tarafından doldurulur) |
| `Used` | `H13:H17` | Her kaynaktan kullanılan toplam miktar |
| `Meet` | `B18:G18` | Her hedefe karşılanan toplam miktar |
| `Demand` | `B20:G20` | Talep kontrol satırı |
| `ToplamMaliyet` | `H19` | Hesaplanan toplam maliyet |

## ⚙️ VBA Makroları

Çalışma kitabı, **en küçük maliyet yöntemi (Least Cost Method)** ile taşıma problemini sezgisel olarak çözen birden fazla makro içerir. Her biri aynı mantığı farklı bir döngü yapısıyla uygular:

| Modül | Makro | Yaklaşım |
|---|---|---|
| `Module1` | `MiktarAta1` | Satır bazlı en küçük maliyet ataması — `For` döngüsü |
| `Module2` | `MiktarAta2` | Aynı mantık — `Do While` döngüsü |
| `Module2` | `MiktarAta3` | Aynı mantık — `Do Until` döngüsü |
| `Module2` | `MiktarAta4` | **Sütun bazlı** en küçük maliyet ataması |
| `Module3` | `MiktarAta6` | Tip tanımlı (`Dim`) ve `Range` nesneleriyle optimize edilmiş temiz versiyon |

**Çalışma mantığı (özet):**
1. `Miktar` aralığı temizlenir.
2. Her kaynak (satır) için, kalan kapasiteyi kullanabileceği en düşük maliyetli hedef seçilir.
3. Kalan kapasite ve kalan talebin minimumu kadar miktar atanır.
4. Toplam kullanılan miktar, toplam kapasiteye eşitlenene kadar işlem tekrarlanır.

### Makroları çalıştırma
Excel'de `Alt + F8` tuşlarına basıp istediğiniz makroyu (`MiktarAta1`, `MiktarAta4`, `MiktarAta6` vb.) seçip **Çalıştır**'a tıklayabilirsiniz. Farklı makrolar farklı sıralama mantığı kullandığından, uygulama sırasına bağlı olarak **farklı (ama geçerli) dağıtım planları** üretebilirler; toplam maliyetler karşılaştırılabilir.

## ✅ Örnek Sonuç

`MiktarAta1` çalıştırıldığında elde edilen dağıtım planı:

| | D1 | D2 | D3 | D4 | D5 | D6 | Kullanılan |
|---|---|---|---|---|---|---|---|
| **F1** | 60 | – | – | – | 590 | – | 650 |
| **F2** | – | – | – | 550 | – | – | 550 |
| **F3** | – | – | 600 | – | – | – | 600 |
| **F4** | 650 | – | – | – | – | 200 | 850 |
| **F5** | 100 | 700 | 50 | 150 | – | – | 1000 |
| **Karşılanan** | 810 | 700 | 650 | 700 | 590 | 200 | |

**Toplam Maliyet = 180.440**

> Not: Bu, en küçük maliyet yönteminin (sezgisel/heuristic) verdiği bir **başlangıç uygun çözümüdür**; doğrusal programlama ile (örn. Excel Solver, Simpleks yöntemi) bulunacak **kesin optimal çözümle** aynı olmayabilir.

## 🛠️ Gereksinimler

- Microsoft Excel (masaüstü sürüm — makrolar mobil/web Excel'de çalışmayabilir)
- Makroların etkinleştirilmiş olması (Dosya > Seçenekler > Güven Merkezi > Makro Ayarları)

## 📌 Notlar
- `Miktar` aralığını sıfırlayıp farklı makroları sırayla deneyerek sonuçları karşılaştırabilirsiniz.

