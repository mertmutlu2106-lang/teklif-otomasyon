namespace TeklifOtomasyon;

/// <summary>Basit teşhis günlüğü (exe yanında teklif-log.log).</summary>
internal static class Log
{
    private static readonly string Yol = Path.Combine(AppContext.BaseDirectory, "teklif-log.log");
    private static readonly object _kilit = new();

    public static void Yaz(string mesaj)
    {
        try
        {
            lock (_kilit)
                File.AppendAllText(Yol, DateTime.Now.ToString("HH:mm:ss.fff") + "  " + mesaj + Environment.NewLine);
        }
        catch { /* log yazılamazsa sessiz geç */ }
    }

    public static void Ayrac() => Yaz(new string('-', 50));
}
