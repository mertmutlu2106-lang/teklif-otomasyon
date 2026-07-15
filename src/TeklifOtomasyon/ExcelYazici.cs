using System.Globalization;
using System.Runtime.InteropServices;

namespace TeklifOtomasyon;

public class YazimSonucu
{
    public bool Bulundu;
    public int Satir;
    public string SayfaAdi = "";
    public string HedefHucre = "";
    public string? EskiIcerik;
    public string? YeniIcerik;
    public string Mesaj = "";
    public string PdfAdi = "";
    public bool Coklu;                        // birden fazla eşleşme (güvenlik için yazılmaz)
    public bool Atlandi;                       // liste zaten yazılmış → tekrar eklenmedi
    public bool AdUyusmazligi;                 // dosya adı ile PDF içeriği çelişiyor (güvenlik için yazılmaz)
    public string? GecmisTarihi;               // bu plaka+branş için daha önce yazılmışsa tarihi
    public List<int> EslesenSatirlar = new();
    public string KaynakMod = "";             // "açık dosya" / "dosya açıldı"
}

/// <summary>TECDİT Excel'ine (COM interop) MERT metnini ekler. Tek oturumda toplu işler.</summary>
public static class ExcelYazici
{
    // Hücredeki mevcut içerik ile aracın bloğu arasına girer. Kullanıcının elle yazdığı
    // blok ayracıyla aynı stil (MertFormatter'ın şirket arası " --- " ayracı bundan ayrı).
    private const string Ayrac = "-----";

    public static List<YazimSonucu> Isle(
        string excelYolu,
        IReadOnlyList<(TeklifCalismasi t, string mert, string pdfAdi)> ogeler,
        bool yaz)
    {
        // Türkçe kültürde COM interop hatası (0x80028018) yaşanmaması için:
        Thread.CurrentThread.CurrentCulture = CultureInfo.GetCultureInfo("en-US");

        // Excel "meşgul" durumunda çağrıları otomatik yeniden dene:
        MessageFilter.Register();

        var sonuclar = new List<YazimSonucu>();
        try
        {
        if (!File.Exists(excelYolu))
        {
            foreach (var o in ogeler)
                sonuclar.Add(new YazimSonucu { PdfAdi = o.pdfAdi, Mesaj = "Excel dosyası bulunamadı: " + excelYolu });
            return sonuclar;
        }

        // 1) Önce ZATEN AÇIK olan workbook'u ara (paylaşımlı/canlı dosya için güvenli yol)
        dynamic? acikWb = RotHelper.AcikCalismaKitabiBul(excelYolu);
        bool bagliMod = acikWb is not null;         // kullanıcının açık oturumu → biz kapatmayız
        string kaynakMod = bagliMod ? "açık dosya (canlı)" : "dosya açıldı";
        Log.Ayrac();
        Log.Yaz($"Isle: yaz={yaz} bagliMod={bagliMod} yol={excelYolu}");

        dynamic? excel = null;
        dynamic wb = null!;

        if (bagliMod)
        {
            wb = acikWb!;
            try { Log.Yaz($"  bagliWB FullName={(string)wb.FullName} Visible={(bool)wb.Application.Visible} ReadOnly={(bool)wb.ReadOnly}"); }
            catch (Exception e) { Log.Yaz("  bagliWB bilgi hatasi: " + e.Message); }
        }
        else
        {
            Type? xlType = Type.GetTypeFromProgID("Excel.Application");
            if (xlType is null)
            {
                foreach (var o in ogeler)
                    sonuclar.Add(new YazimSonucu { PdfAdi = o.pdfAdi, Mesaj = "Excel COM bulunamadı." });
                return sonuclar;
            }
            excel = Activator.CreateInstance(xlType)!;
            excel.Visible = false;
            excel.DisplayAlerts = false;
            try
            {
                wb = excel.Workbooks.Open(excelYolu);
                Log.Yaz("  kendi actik (yeni gizli instance)");
            }
            catch (Exception ex)
            {
                try { excel.Quit(); Marshal.FinalReleaseComObject(excel); } catch { }
                foreach (var o in ogeler)
                    sonuclar.Add(new YazimSonucu { PdfAdi = o.pdfAdi, Mesaj = "Excel açılamadı: " + ex.Message });
                return sonuclar;
            }
        }

        bool degisti = false;
        try
        {
            foreach (var (t, mert, pdfAdi) in ogeler)
            {
                var s = TekOge(wb, t, mert, pdfAdi, yaz, ref degisti);
                s.KaynakMod = kaynakMod;
                sonuclar.Add(s);
            }
            if (yaz && degisti) { wb.Save(); Log.Yaz("  wb.Save() cagrildi"); }
        }
        finally
        {
            if (bagliMod)
            {
                // Kullanıcının oturumu — KAPATMA, sadece kendi referansımızı bırak.
                try { Marshal.ReleaseComObject(wb); } catch { }
            }
            else
            {
                try { wb.Close(false); } catch { }
                try { excel!.Quit(); } catch { }
                try { Marshal.FinalReleaseComObject(excel!); } catch { }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        return sonuclar;
        }
        finally { MessageFilter.Revoke(); }
    }

    private static YazimSonucu TekOge(dynamic wb, TeklifCalismasi t, string mert, string pdfAdi, bool yaz, ref bool degisti)
    {
        (string sayfa, int notCol, string harf) = t.Brans == Brans.Trafik
            ? ("TRAFİK", 14, "N")   // TRAFİK sayfası: AÇIKLAMA = N
            : ("KASKO", 18, "R");   // KASKO sayfası: NOT = R

        var s = new YazimSonucu { SayfaAdi = sayfa, PdfAdi = pdfAdi };

        // Branş güvenliği: PDF KASKO/TRAFİK değilse hiç yazma
        if (t.Brans == Brans.Bilinmiyor)
        {
            s.Mesaj = "Branş belirlenemedi (KASKO/TRAFİK PDF'i değil?).";
            return s;
        }

        // Dosya adı çapraz kontrolü: Open adın içine plaka/branş/TC yazıyor. Parser'ın PDF
        // içinden okuduğuyla çelişiyorsa yanlış satıra yazma riski var → hiç yazma.
        var adBilgi = DosyaAdi.Coz(pdfAdi);
        string? cakisma = AdCakismasi(adBilgi, t);
        if (cakisma is not null)
        {
            s.AdUyusmazligi = true;
            s.Mesaj = $"Dosya adı PDF içeriğiyle çelişiyor ({cakisma}) — güvenlik için YAZILMADI.";
            Log.Yaz($"  {pdfAdi}: AD UYUSMAZLIGI {cakisma}");
            return s;
        }

        s.GecmisTarihi = YazimGecmisi.Bul(t.Plaka, t.Brans)?.Tarih;

        dynamic ws = wb.Worksheets[sayfa];
        dynamic used = ws.UsedRange;
        int lastRow = (int)used.Row + (int)used.Rows.Count - 1;

        dynamic plakaVals = ws.Range["E1:E" + lastRow].Value2;
        dynamic tcVals = ws.Range["D1:D" + lastRow].Value2;
        dynamic pturVals = ws.Range["N1:N" + lastRow].Value2;

        string hedefPlaka = Norm(t.Plaka);
        string hedefTc = (t.Tc ?? "").Trim();

        // TÜM eşleşmeleri topla (çoklu eşleşme yakalanır)
        var eslesen = new List<int>();
        for (int r = 3; r <= lastRow; r++)
        {
            if (Norm(Str(plakaVals[r, 1])) != hedefPlaka) continue;
            string tc = Str(tcVals[r, 1]).Trim();
            if (hedefTc.Length > 0 && tc.Length > 0 && tc != hedefTc) continue;
            if (t.Brans == Brans.Kasko)
            {
                string ptur = Str(pturVals[r, 1]).Trim().ToUpperInvariant();
                if (ptur != "KASKO") continue;
            }
            eslesen.Add(r);
        }
        s.EslesenSatirlar = eslesen;
        Log.Yaz($"  {pdfAdi}: sayfa={sayfa} eslesen=[{string.Join(",", eslesen)}]");

        if (eslesen.Count == 0)
        {
            s.Mesaj = $"{sayfa} sayfasında {t.Plaka} bulunamadı.";
            return s;
        }
        if (eslesen.Count > 1)
        {
            s.Coklu = true;
            s.Mesaj = $"{eslesen.Count} eşleşme (satır {string.Join(", ", eslesen)}) — güvenlik için YAZILMADI, elle kontrol et.";
            return s;
        }

        int satir = eslesen[0];
        s.Bulundu = true;
        s.Satir = satir;
        s.HedefHucre = $"{sayfa}!{harf}{satir}";

        dynamic hucre = ws.Cells[satir, notCol];
        string eski = Str(hucre.Value2);
        s.EskiIcerik = eski;

        // Mükerrer koruması: aynı liste zaten hücredeyse tekrar ekleme
        if (!string.IsNullOrWhiteSpace(eski) && eski.Contains(mert, StringComparison.Ordinal))
        {
            s.Atlandi = true;
            s.YeniIcerik = eski;
            s.Mesaj = "Bu liste zaten yazılmış — atlandı.";
            Log.Yaz($"  {pdfAdi}: ATLANDI (zaten var) {s.HedefHucre}");
            return s;
        }

        string yeni = string.IsNullOrWhiteSpace(eski) ? mert : eski + Ayrac + mert;
        s.YeniIcerik = yeni;

        if (yaz)
        {
            hucre.Value2 = yeni;
            degisti = true;
            s.Mesaj = "Yazıldı.";
            try
            {
                string kontrol = Str(ws.Cells[satir, notCol].Value2);
                Log.Yaz($"  {pdfAdi}: YAZILDI {s.HedefHucre} yeniLen={yeni.Length} okunanLen={kontrol.Length} wb={(string)wb.FullName}");
            }
            catch (Exception e) { Log.Yaz("  yazim-sonrasi okuma hatasi: " + e.Message); }
        }
        else
        {
            s.Mesaj = "Önizleme (yazılmadı).";
        }
        return s;
    }

    /// <summary>Dosya adındaki plaka/TC ile PDF'ten okunanları karşılaştırır. İkisi de doluyken
    /// farklıysa çakışma adını döner; biri boşsa (ör. TcGizle açıksa) kontrol atlanır.</summary>
    private static string? AdCakismasi(DosyaAdiBilgi ad, TeklifCalismasi t)
    {
        if (ad.Plaka is not null && !string.IsNullOrWhiteSpace(t.Plaka) && Norm(ad.Plaka) != Norm(t.Plaka))
            return $"plaka: ad={ad.Plaka} pdf={t.Plaka}";

        if (ad.Tc is not null && !string.IsNullOrWhiteSpace(t.Tc) && ad.Tc != t.Tc.Trim())
            return "TC";

        if (ad.Brans != Brans.Bilinmiyor && ad.Brans != t.Brans)
            return $"branş: ad={ad.Brans} pdf={t.Brans}";

        return null;
    }

    private static string Str(object? o) => o?.ToString() ?? "";

    internal static string Norm(string? s) =>
        (s ?? "").Replace(" ", "").Replace("-", "").ToUpperInvariant();
}
