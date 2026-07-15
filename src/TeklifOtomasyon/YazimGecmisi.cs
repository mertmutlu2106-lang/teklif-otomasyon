using System.Text.Json;

namespace TeklifOtomasyon;

public class GecmisKayit
{
    public string Tarih { get; set; } = "";
    public string Mert { get; set; } = "";
}

/// <summary>
/// Aracın yazdıklarının yerel kaydı. TECDİT hücresindeki bloklar ayrıştırılamadığı için
/// (ayraç 2-20 tire arası elle yazılıyor) "bu müşteriye zaten yazdım mı?" sorusu buradan cevaplanır.
/// KVKK: plaka içerir → <c>*.local.json</c> gitignore'da.
/// </summary>
public static class YazimGecmisi
{
    private static string Yol => Path.Combine(AppContext.BaseDirectory, "yazim-gecmisi.local.json");
    private static Dictionary<string, GecmisKayit>? _kayitlar;

    private static Dictionary<string, GecmisKayit> Yukle()
    {
        if (_kayitlar is not null) return _kayitlar;
        try
        {
            if (File.Exists(Yol))
            {
                var okunan = JsonSerializer.Deserialize<Dictionary<string, GecmisKayit>>(File.ReadAllText(Yol));
                if (okunan is not null)
                    _kayitlar = new Dictionary<string, GecmisKayit>(okunan, StringComparer.OrdinalIgnoreCase);
            }
        }
        catch { /* bozuk dosya → boş geçmiş */ }
        return _kayitlar ??= new Dictionary<string, GecmisKayit>(StringComparer.OrdinalIgnoreCase);
    }

    private static string Anahtar(string? plaka, Brans brans) => $"{ExcelYazici.Norm(plaka)}|{brans}";

    public static GecmisKayit? Bul(string? plaka, Brans brans)
        => Yukle().TryGetValue(Anahtar(plaka, brans), out var k) ? k : null;

    public static void Kaydet(string? plaka, Brans brans, string mert)
    {
        var d = Yukle();
        d[Anahtar(plaka, brans)] = new GecmisKayit
        {
            Tarih = DateTime.Now.ToString("dd.MM.yyyy HH:mm"),
            Mert = mert
        };
        try
        {
            File.WriteAllText(Yol, JsonSerializer.Serialize(d, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch (Exception ex) { Log.Yaz("Gecmis yazilamadi: " + ex.Message); }
    }
}
