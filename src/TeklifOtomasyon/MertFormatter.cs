using System.Text;

namespace TeklifOtomasyon;

/// <summary>Teklifleri MERT'in Excel not sütunu formatına çevirir.</summary>
public static class MertFormatter
{
    // Ayarlanabilir küçük tercihler (kuru modda görüp ince ayar yapılacak)
    private const string Ayrac = " --- ";
    private const string FiyatSonEki = " TL";

    // Ham şirket adı → MERT'in yazdığı görünen ad
    private static readonly Dictionary<string, string> AdEslesme = new(StringComparer.OrdinalIgnoreCase)
    {
        ["HDI_PLUS"] = "HDI",
        ["TURKIYE"]  = "TÜRKİYE",
        ["HEPIYI"]   = "HEPİYİ",
    };

    /// <summary>
    /// KASKO → tüm şirketler; TRAFİK → en ucuz 3. Ucuzdan pahalıya, kuruş atılmış.
    /// </summary>
    public static string Uret(TeklifCalismasi t)
    {
        int adet = t.Brans == Brans.Trafik ? 3 : int.MaxValue;
        var secili = t.Teklifler
            .OrderBy(x => x.Fiyat)
            .Take(adet)
            .Select(x => $"{Ad(x.HamAd)} : {Fiyat(x.Fiyat)}{FiyatSonEki}");
        return string.Join(Ayrac, secili);
    }

    private static string Ad(string hamAd) =>
        AdEslesme.TryGetValue(hamAd, out var d) ? d : hamAd.ToUpperInvariant();

    // Kuruş atılır → tam TL, binlik ayırıcı yok. Örn: 12.087,90 → 12087
    private static string Fiyat(decimal fiyat) => ((long)Math.Floor(fiyat)).ToString();
}
