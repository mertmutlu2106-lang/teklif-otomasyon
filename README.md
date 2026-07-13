# Teklif Otomasyonu

OpenHizliTeklif'ten "PDF Aktar" ile çıkan **KASKO / TRAFİK teklif PDF'lerini** okuyup, sigorta şirketi tekliflerini **MERT formatında** TECDİT Excel'inin not sütununa otomatik geçiren masaüstü aracı.

> Amaç: danışmanın her müşteri için ~12 şirketin fiyatını **elle** Excel'e yazma zahmetini ortadan kaldırmak.

## Ne yapar (v1)

1. `... - KASKO.pdf` ve `... - TRAFİK.pdf` dosyalarını okur (PdfPig).
2. Şirket → fiyat tablosunu çıkarır, ucuzdan pahalıya sıralar.
3. **MERT formatında** metni kurar:
   - **KASKO → R sütunu:** PDF'teki **tüm** şirketler.
   - **TRAFİK → N sütunu:** en ucuz **3** teklif.
   - Biçim: `ŞİRKET : FİYAT --- ŞİRKET : FİYAT ...` (kuruş atılır, tam TL).
4. Açık olan TECDİT Excel'inde **plaka (+TC)** ile satırı bulur, not sütununa **üzerine yazmadan ekler**.

> KASKO'da yenileme teklifi + araç başlığı **elle** eklenir (kapsam dışı, kararlaştırıldı).

## Teknik

- **C# / .NET 9** — tek, kurulumsuz `.exe` olarak yayınlanır (`dotnet publish -c Release -p:PublishSingleFile=true`).
- PDF: [UglyToad.PdfPig](https://github.com/UglyToad/PdfPig)
- Excel: COM interop (zaten açık olan paylaşımlı çalışma kitabına güvenli yazım).

## KVKK / Gizlilik

- Tamamen **yerel / ofis içi** çalışır; hiçbir veri buluta gönderilmez.
- Müşteri verisi içeren dosyalar (`*.pdf`, `*.xlsx`, TECDİT, örnekler) `.gitignore` ile **repoya alınmaz**.
- Test daima TECDİT'in bir **kopyası** üzerinde yapılır.

## Durum

🚧 Geliştirme aşamasında — v1 iskeleti. Ayrıntılı plan: [`docs/plan.md`](docs/plan.md)
