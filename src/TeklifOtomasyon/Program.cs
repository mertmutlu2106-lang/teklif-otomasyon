namespace TeklifOtomasyon;

static class Program
{
    [STAThread]
    static void Main(string[] args)
    {
        var pdfler = args.Where(a => a.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)).ToList();
        var excel = args.FirstOrDefault(a => a.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase));
        bool yaz = args.Any(a => a.Equals("--yaz", StringComparison.OrdinalIgnoreCase));

        // Argümanla PDF+Excel verilmişse konsol (test) modu:
        if (pdfler.Count > 0 && excel is not null)
        {
            KonsolCalistir(excel, pdfler, yaz);
            return;
        }

        // Aksi halde arayüz:
        ApplicationConfiguration.Initialize();
        Application.Run(new MainForm());
    }

    private static void KonsolCalistir(string excel, List<string> pdfler, bool yaz)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        var ogeler = new List<(TeklifCalismasi t, string mert, string pdfAdi)>();
        foreach (var pdf in pdfler)
        {
            if (!File.Exists(pdf)) { Console.WriteLine($"⚠ PDF yok: {pdf}"); continue; }
            try
            {
                var t = TeklifParser.Parse(pdf);
                // Parser branşı bulamazsa dosya adındaki branş devreye girer.
                if (t.Brans == Brans.Bilinmiyor) t.Brans = DosyaAdi.Coz(pdf).Brans;
                var mert = MertFormatter.Uret(t);
                Console.WriteLine($"{Path.GetFileName(pdf)} [{t.Brans}] {t.Plaka} TC={t.Tc}");
                Console.WriteLine($"  MERT: {mert}");
                ogeler.Add((t, mert, Path.GetFileName(pdf)));
            }
            catch (Exception ex) { Console.WriteLine($"⚠ {Path.GetFileName(pdf)} okunamadı: {ex.Message}"); }
        }
        if (ogeler.Count == 0) { Console.WriteLine("İşlenecek geçerli PDF yok."); return; }
        try
        {
            var res = ExcelYazici.Isle(excel, ogeler, yaz);
            if (res.Count > 0)
                Console.WriteLine($"  [Excel modu: {res[0].KaynakMod}]");
            foreach (var d in res)
            {
                if (d.GecmisTarihi is not null)
                    Console.WriteLine($"  ⚠ {d.PdfAdi} : bu plaka için {d.GecmisTarihi} tarihinde zaten liste yazılmış.");
                Console.WriteLine($"  {d.PdfAdi} → {(d.Bulundu ? d.HedefHucre : "-")} : {d.Mesaj}");
            }

            // Arayüzle aynı davranış: başarılı yazımlar geçmişe kaydedilir.
            if (yaz)
                for (int i = 0; i < res.Count && i < ogeler.Count; i++)
                    if (res[i].Bulundu && !res[i].Atlandi)
                        YazimGecmisi.Kaydet(ogeler[i].t.Plaka, ogeler[i].t.Brans, ogeler[i].mert);
        }
        catch (Exception ex) { Console.WriteLine("  Hata: " + ex.Message); }
    }
}
