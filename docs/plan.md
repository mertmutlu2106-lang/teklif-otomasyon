# Teklif → Excel Otomasyonu (v1) — Plan

> Not: Bu dosyada gerçek müşteri bilgisi yoktur; örnekler maskelidir (KVKK).

## Bağlam
Danışman OpenHizliTeklif'te KASKO+TRAFİK teklifi çekiyor, sonra şirketlerin fiyatını **elle** TECDİT Excel'inin not sütununa yazıyor (aylık ~1.770 satır → gecikme + hata). Amaç: bu elle yazımı otomatikleştirmek. En kırılgan parça (programa otomatik veri girişi) kapsam dışı: kullanıcı teklifi elle çeker → "PDF Aktar" → araç, PDF'i okuyup şirket tekliflerini MERT formatında Excel'e geçirir.

## Kapsam (v1)
- **Girdi:** "PDF Aktar" ile çıkan 2 dosya: `... - KASKO.pdf`, `... - TRAFİK.pdf`. Metin katmanı temiz, sıralı şirket→fiyat tablosu + araç bilgileri okunuyor.
- **KASKO → R sütunu:** PDF'teki **TÜM şirketler**, ucuzdan pahalıya, MERT formatında. Yenileme teklifi + araç başlığı **elle** eklenir.
- **TRAFİK → N sütunu:** en ucuz **3** teklif, MERT formatında.
- Eşleştirme: **plaka (+TC)**; not sütununa **ekleyerek** yazar.

## MERT formatı
- `{ŞİRKET} : {FİYAT} --- {ŞİRKET} : {FİYAT} --- ...` (ucuzdan pahalıya, kuruş atılır → tam TL).
- Şirket adı dönüşümü: `HDI_PLUS→HDI`, `TURKIYE→TÜRKİYE`, `HEPIYI→HEPİYİ` (eşleştirme tablosu).
- Örnek (maskeli): KASKO → `ORIENT : 12087 --- HEPİYİ : 13562 --- ... --- AXA : 50955`  •  TRAFİK (ilk 3) → `ANADOLU : 10210 --- EUREKO : 10635 --- ANA : 11188`

## PDF'ten çıkan alanlar
plaka, TC, belge seri, şasi, motor, marka/tip, model yılı, kasko değeri, branş, sıralı şirket→fiyat tablosu, hazırlayan+tarih.

## Teknik
- **C# / .NET 9** → tek kurulumsuz `.exe` (`dotnet publish -p:PublishSingleFile`).
- PDF: UglyToad.PdfPig (saf managed).
- Excel: COM interop ile açık paylaşımlı çalışma kitabına yazım (çakışma riski en düşük yol).

## Akış
PDF sürükle-bırak veya klasör izle → oku → MERT metnini kur (kasko=tümü, trafik=ilk 3) → açık Excel'de satırı bul (plaka+TC) → R/N hücresine ekle → log tut.

## Doğrulama
1. **Kuru mod:** örnek PDF'lerle çalıştır → beklenen KASKO (tümü) ve TRAFİK (ilk 3) satırlarını ekrana bassın.
2. TECDİT'in **test kopyasında** çalıştır → doğru satır bulunup R/N'ye eklendiğini doğrula.

## Açık maddeler (v1'i bloklamaz)
- TRAFİK için MERT `ŞİRKET : fiyat` formatı teyidi.
- Yenileme/araç başlığı KASKO'da manuel (kararlaştırıldı).

## KVKK
Tamamen yerel/ofis-içi. Test daima kopya dosyada. Müşteri verisi repoya/buluta gitmez.
