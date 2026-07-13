using System.Text.Json;

namespace TeklifOtomasyon;

/// <summary>Basit yerel ayar (hedef Excel yolu). exe yanında json olarak saklanır.</summary>
public class Config
{
    public string? ExcelYolu { get; set; }

    private static string Yol => Path.Combine(AppContext.BaseDirectory, "appsettings.local.json");

    public static Config Yukle()
    {
        try
        {
            if (File.Exists(Yol))
                return JsonSerializer.Deserialize<Config>(File.ReadAllText(Yol)) ?? new Config();
        }
        catch { /* bozuk ayar → varsayılan */ }
        return new Config();
    }

    public void Kaydet()
    {
        try
        {
            File.WriteAllText(Yol, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
        }
        catch { /* yazılamazsa sessiz geç */ }
    }
}
