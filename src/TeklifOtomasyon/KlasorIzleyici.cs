namespace TeklifOtomasyon;

/// <summary>
/// TEKLIF-GELEN klasörüne düşen Open export PDF'lerini yakalar. Yakalayınca <c>yeniPdf</c>
/// callback'ini çağırır — callback ARKA PLAN thread'inde çalışır, UI'ya marshal etmek çağıranın işi.
/// </summary>
public sealed class KlasorIzleyici : IDisposable
{
    private readonly Action<string> _yeniPdf;
    private readonly Dictionary<string, DateTime> _sonGoruldu = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _kilit = new();
    private FileSystemWatcher? _fsw;

    public KlasorIzleyici(Action<string> yeniPdf) => _yeniPdf = yeniPdf;

    public bool Calisiyor => _fsw?.EnableRaisingEvents == true;

    public void Basla(string? klasor)
    {
        Dur();
        if (string.IsNullOrWhiteSpace(klasor) || !Directory.Exists(klasor))
        {
            Log.Yaz("Izleyici: klasor yok, baslamadi -> " + (klasor ?? "(bos)"));
            return;
        }

        _fsw = new FileSystemWatcher(klasor, DosyaAdi.Filtre)
        {
            IncludeSubdirectories = false,
            NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite | NotifyFilters.Size
        };
        // Kaydeden uygulamaya göre dosya ya doğrudan oluşur ya da geçici addan yeniden adlandırılır.
        _fsw.Created += (_, e) => Kuyrukla(e.FullPath);
        _fsw.Renamed += (_, e) => Kuyrukla(e.FullPath);
        _fsw.Error += (_, e) => Log.Yaz("Izleyici hatasi: " + e.GetException().Message);
        _fsw.EnableRaisingEvents = true;
        Log.Yaz("Izleyici basladi: " + klasor);
    }

    public void Dur()
    {
        if (_fsw is null) return;
        try
        {
            _fsw.EnableRaisingEvents = false;
            _fsw.Dispose();
            Log.Yaz("Izleyici durdu.");
        }
        catch { /* zaten kapanmış */ }
        _fsw = null;
    }

    private void Kuyrukla(string yol)
    {
        lock (_kilit)
        {
            // FileSystemWatcher tek dosya için birden fazla olay üretir.
            if (_sonGoruldu.TryGetValue(yol, out var t) && DateTime.UtcNow - t < TimeSpan.FromSeconds(2)) return;
            _sonGoruldu[yol] = DateTime.UtcNow;
        }

        Task.Run(() =>
        {
            if (DosyaHazir(yol))
            {
                Log.Yaz("Izleyici yakaladi: " + Path.GetFileName(yol));
                try { _yeniPdf(yol); } catch (Exception ex) { Log.Yaz("Izleyici callback hatasi: " + ex.Message); }
            }
            else
            {
                Log.Yaz("Izleyici: dosya acilamadi, birakildi -> " + Path.GetFileName(yol));
            }
        });
    }

    /// <summary>Kaydeden uygulama dosyayı hâlâ yazıyor olabilir; kilit açılana kadar bekler (en fazla ~10sn).</summary>
    private static bool DosyaHazir(string yol)
    {
        for (int i = 0; i < 40; i++)
        {
            try
            {
                using var fs = File.Open(yol, FileMode.Open, FileAccess.Read, FileShare.None);
                if (fs.Length > 0) return true;
            }
            catch (IOException) { /* hâlâ yazılıyor ya da silinmiş */ }
            catch (UnauthorizedAccessException) { }
            Thread.Sleep(250);
        }
        return false;
    }

    public void Dispose() => Dur();
}
