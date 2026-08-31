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

