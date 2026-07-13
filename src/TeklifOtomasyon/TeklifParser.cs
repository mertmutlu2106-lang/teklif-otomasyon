using System.Text.RegularExpressions;
using UglyToad.PdfPig;

namespace TeklifOtomasyon;

/// <summary>OpenHizliTeklif "PDF Aktar" çıktısını yapısal veriye çevirir.</summary>
public static class TeklifParser
{
    // "1 ORIENT SİGORTA 12,087.90 ₺ ..." → rank, ad, fiyat
    private static readonly Regex SirketSatiri = new(
        @"^\s*(?<rank>\d{1,2})\s+(?<ad>.+?)\s+S[İI]GORTA\s+(?<fiyat>[\d.,]+)\s*₺",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static TeklifCalismasi Parse(string pdfYolu)
    {
        var lines = SayfaSatirlari(pdfYolu);
        var metin = string.Join("\n", lines);
        var t = new TeklifCalismasi { Brans = BransBul(metin) };

        // --- Araç / sigortalı bilgileri (best-effort, regex ile) ---
        t.MusteriAdi = Ilk(metin, @"Sigortal[ıi]\s*[ÜU]nvan:?\s*([^\n]+)")
                       ?? Ilk(metin, @"Sayın\s+([^\n]+)");
        t.Plaka     = Ilk(metin, @"\b(\d{2}[A-ZÇĞİÖŞÜ]{1,3}\d{2,4})\b");
        t.Tc        = Ilk(metin, @"(?:TC\s*/\s*Vergi\s*Kimlik\s*No:?)\s*(\d{10,11})") ?? Ilk(metin, @"\b(\d{11})\b");
        t.BelgeSeri = Ilk(metin, @"Ruhsat\s*No:?\s*([A-Z]{1,3}\d{4,})");
        t.Sasi      = Ilk(metin, @"\b([A-HJ-NPR-Z0-9]{17})\b");
        t.Motor     = Ilk(metin, @"Motor\s*No:?\s*([A-Z0-9]{6,})");
        t.KaskoDegeri = Ilk(metin, @"Kasko\s*De[ğg]eri:?\s*(\d+)");
        t.Hazirlayan  = Ilk(metin, @"([A-ZÇĞİÖŞÜ]+\s+[A-ZÇĞİÖŞÜ]+)\s*-\s*\d{2}\.\d{2}\.\d{4}");
        t.Tarih       = Ilk(metin, @"(\d{2}\.\d{2}\.\d{4}\s+\d{2}:\d{2})");

        // Marka - Tip:  "144 - TOYOTA | 1004 - YARIS 1.33 STYLE | 2012"
        var markaTip = Ilk(metin, @"Ara[çc]\s*Marka\s*-\s*Tip\s*(.+?)(?:\n|$)");
        if (markaTip is not null)
        {
            var parcalar = markaTip.Split('|', StringSplitOptions.TrimEntries);
            if (parcalar.Length >= 3)
            {
                var marka = KoduAt(parcalar[0]);   // "144 - TOYOTA" → "TOYOTA"
                var tip   = KoduAt(parcalar[1]);   // "1004 - YARIS 1.33 STYLE" → "YARIS 1.33 STYLE"
                t.MarkaTip  = $"{marka} {tip}".Trim();
                t.ModelYili = parcalar[2].Trim();
            }
        }

        // --- Şirket → fiyat tablosu ---
        foreach (var line in lines)
        {
            var m = SirketSatiri.Match(line);
            if (!m.Success) continue;
            var ad = m.Groups["ad"].Value.Trim();
            if (!decimal.TryParse(m.Groups["fiyat"].Value,
                    System.Globalization.NumberStyles.Number,
                    System.Globalization.CultureInfo.InvariantCulture, out var fiyat))
                continue;
            t.Teklifler.Add(new SirketTeklif(ad, fiyat));
        }

        return t;
    }

    /// <summary>Kelimeleri Y koordinatına göre satırlara gruplar (görsel sıra).</summary>
    private static List<string> SayfaSatirlari(string pdfYolu)
    {
        var result = new List<string>();
        using var doc = PdfDocument.Open(pdfYolu);
        foreach (var page in doc.GetPages())
        {
            var lines = page.GetWords()
                .GroupBy(w => Math.Round(w.BoundingBox.Bottom / 3.0))
                .OrderByDescending(g => g.Key)
                .Select(g => string.Join(" ", g.OrderBy(w => w.BoundingBox.Left).Select(w => w.Text)));
            result.AddRange(lines);
        }
        return result;
    }

    private static Brans BransBul(string metin)
    {
        if (Regex.IsMatch(metin, @"KASKO\s+fiyat", RegexOptions.IgnoreCase)) return Brans.Kasko;
        if (Regex.IsMatch(metin, @"TRAF[İI]K\s+fiyat", RegexOptions.IgnoreCase)) return Brans.Trafik;
        return Brans.Bilinmiyor;
    }

    private static string? Ilk(string metin, string pattern)
    {
        var m = Regex.Match(metin, pattern, RegexOptions.IgnoreCase);
        return m.Success ? m.Groups[m.Groups.Count > 1 ? 1 : 0].Value.Trim() : null;
    }

    private static string KoduAt(string s)
    {
        // "144 - TOYOTA" → "TOYOTA"
        var idx = s.IndexOf('-');
        return (idx >= 0 ? s[(idx + 1)..] : s).Trim();
    }
}
