using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace TeklifOtomasyon;

/// <summary>
/// O an AÇIK olan bir Excel çalışma kitabını (Workbook) bulur.
/// ROT'ta workbook'lar path ile kayıtlı DEĞİL — sadece Excel.Application
/// nesneleri var. Bu yüzden her Application'ın .Workbooks koleksiyonu gezilip
/// FullName ile eşleştirilir. Görünür + yazılabilir kitap tercih edilir.
/// </summary>
internal static class RotHelper
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable prot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    public static object? AcikCalismaKitabiBul(string yol)
    {
        string hedefTam = yol;
        string hedefAd = SafeName(yol);

        IRunningObjectTable? rot = null;
        IEnumMoniker? enumMon = null;
        IBindCtx? ctx = null;
        object? yedek = null;      // yazılabilir ama görünür değilse yedek

        try
        {
            if (GetRunningObjectTable(0, out rot) != 0 || rot is null) return null;
            rot.EnumRunning(out enumMon);
            enumMon.Reset();
            CreateBindCtx(0, out ctx);

            int appSayac = 0, esaday = 0;
            var mons = new IMoniker[1];
            while (enumMon.Next(1, mons, IntPtr.Zero) == 0)
            {
                var mon = mons[0];
                object? obj = null;
                try { rot.GetObject(mon, out obj); } catch { }
                Release(mon);
                if (obj is null) continue;

                // Bu nesne bir Excel.Application mı? .Workbooks koleksiyonunu dene.
                dynamic? wbs = null;
                int cnt;
                try { dynamic app = obj; wbs = app.Workbooks; cnt = (int)wbs.Count; }
                catch { Release(obj); continue; }   // Excel.Application değil → atla

                appSayac++;
                for (int i = 1; i <= cnt; i++)
                {
                    dynamic? wb = null;
                    try { wb = wbs![i]; } catch { continue; }

                    string fn = "";
                    try { fn = (string)wb!.FullName; } catch { fn = ""; }

                    bool eslesme = !string.IsNullOrEmpty(fn) &&
                        (string.Equals(fn, hedefTam, StringComparison.OrdinalIgnoreCase) ||
                         string.Equals(SafeName(fn), hedefAd, StringComparison.OrdinalIgnoreCase));

                    if (!eslesme) { Release(wb); continue; }

                    esaday++;
                    bool yazilabilir = false, gorunur = false;
                    try { yazilabilir = !(bool)wb!.ReadOnly; } catch { }
                    try { gorunur = (bool)wb!.Application.Visible; } catch { }
                    Log.Yaz($"  ROT eslesme: yazilabilir={yazilabilir} gorunur={gorunur} fn={fn}");

                    if (yazilabilir && gorunur)
                    {
                        Log.Yaz("  ROT: gorunur+yazilabilir SECILDI");
                        Release(obj);   // app RCW referansımızı bırak (Excel kapanmaz)
                        return wb!;
                    }
                    if (yazilabilir && yedek is null) { yedek = wb; continue; }
                    Release(wb);
                }
                Release(obj);
            }
            Log.Yaz($"  ROT taramasi: {appSayac} Excel.App, {esaday} eslesme, yedek={(yedek is not null)}");
            if (yedek is not null) return yedek;
        }
        catch (Exception e) { Log.Yaz("  RotHelper hata: " + e.Message); }
        finally
        {
            if (ctx is not null) Release(ctx);
            if (enumMon is not null) Release(enumMon);
            if (rot is not null) Release(rot);
        }
        return null;
    }

    private static string SafeName(string s)
    {
        try { return Path.GetFileName(s); } catch { return s; }
    }

    private static void Release(object? o)
    {
        try { if (o is not null && Marshal.IsComObject(o)) Marshal.ReleaseComObject(o); } catch { }
    }
}
