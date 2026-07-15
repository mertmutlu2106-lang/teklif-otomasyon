using System.Text.RegularExpressions;

namespace TeklifOtomasyon;

/// <summary>Dosya adından çözülen bilgiler. Çözülemezse alanlar boş / Bilinmiyor döner.</summary>
public record DosyaAdiBilgi(string? Plaka, Brans Brans, string? Tc);

/// <summary>
/// Open'ın export dosya adını çözer: <c>06FIK201-KASKO-17374155632-logosuz.pdf</c>
/// → plaka, branş, TC. Parser'ın PDF içinden okuduklarına karşı çapraz kontrol için kullanılır.
/// </summary>
public static class DosyaAdi
{
    /// <summary>Klasör izleme filtresi. Bu kalıba başka hiçbir PDF uymuyor (poliçe, makbuz, dekont vb.).</summary>
    public const string Filtre = "*-logosuz.pdf";

    private static readonly Regex Kalip = new(
        @"^(?<plaka>[A-Z0-9]+)-(?<brans>KASKO|TRAF[İIiı]K)-(?<tc>\d{11})-logosuz\.pdf$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    public static DosyaAdiBilgi Coz(string yol)
    {
        var m = Kalip.Match(Path.GetFileName(yol));
        if (!m.Success) return new DosyaAdiBilgi(null, Brans.Bilinmiyor, null);

        var brans = m.Groups["brans"].Value.StartsWith("KASKO", StringComparison.OrdinalIgnoreCase)
            ? Brans.Kasko
            : Brans.Trafik;

        return new DosyaAdiBilgi(m.Groups["plaka"].Value, brans, m.Groups["tc"].Value);
    }
}
