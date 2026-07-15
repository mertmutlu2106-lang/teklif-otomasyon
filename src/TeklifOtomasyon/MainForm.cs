using System.Runtime.InteropServices;
using System.Text;

namespace TeklifOtomasyon;

public class MainForm : Form
{
    private readonly Config _config = Config.Yukle();
    private readonly TextBox _excelKutu = new();
    private readonly TextBox _klasorKutu = new();
    private readonly CheckBox _izleKutu = new();
    private readonly Panel _dropPanel = new();
    private readonly RichTextBox _onizleme = new();
    private readonly Button _yazBtn = new();
    private readonly Button _temizleBtn = new();
    private readonly Label _durum = new();
    private List<(TeklifCalismasi t, string mert, string pdfAdi)> _bekleyen = new();
    private KlasorIzleyici? _izleyici;

    public MainForm()
    {
        Text = "Teklif Otomasyonu";
        Width = 780; Height = 620; MinimumSize = new Size(640, 480);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9f);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(12) };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // ── Satır 0: hedef Excel seçici ─────────────────────
        var excelRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 0, 0, 4) };
        excelRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        excelRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        excelRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var lbl = new Label { Text = "Hedef Excel:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 6, 0) };
        _excelKutu.ReadOnly = true; _excelKutu.Dock = DockStyle.Fill;
        _excelKutu.Text = _config.ExcelYolu ?? "(seçilmedi)";
        var secBtn = new Button { Text = "Seç…", AutoSize = true };
        secBtn.Click += (_, _) => ExcelSec();
        excelRow.Controls.Add(lbl, 0, 0);
        excelRow.Controls.Add(_excelKutu, 1, 0);
        excelRow.Controls.Add(secBtn, 2, 0);

        // ── Satır 1: izlenecek klasör ───────────────────────
        var klasorRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 5, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
        klasorRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        klasorRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        klasorRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        klasorRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        klasorRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        var klbl = new Label { Text = "İzlenecek klasör:", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 7, 6, 0) };
        _klasorKutu.ReadOnly = true; _klasorKutu.Dock = DockStyle.Fill;
        _klasorKutu.Text = _config.IzlenecekKlasor ?? "(seçilmedi)";
        var klasorBtn = new Button { Text = "Seç…", AutoSize = true };
        klasorBtn.Click += (_, _) => KlasorSec();
        _izleKutu.Text = "İzle"; _izleKutu.AutoSize = true; _izleKutu.Margin = new Padding(8, 6, 4, 0);
        _izleKutu.Checked = _config.IzlemeAcik;
        _izleKutu.CheckedChanged += (_, _) => IzlemeDegisti();
        var taraBtn = new Button { Text = "Klasörü Tara", AutoSize = true };
        taraBtn.Click += (_, _) => Tara();
        klasorRow.Controls.Add(klbl, 0, 0);
        klasorRow.Controls.Add(_klasorKutu, 1, 0);
        klasorRow.Controls.Add(klasorBtn, 2, 0);
        klasorRow.Controls.Add(_izleKutu, 3, 0);
        klasorRow.Controls.Add(taraBtn, 4, 0);

        // ── Satır 2: sürükle-bırak alanı ────────────────────
        _dropPanel.Dock = DockStyle.Fill;
        _dropPanel.AllowDrop = true;
        _dropPanel.BackColor = Color.FromArgb(245, 247, 250);
        _dropPanel.BorderStyle = BorderStyle.FixedSingle;
        _dropPanel.Margin = new Padding(0, 0, 0, 8);
        var dropLbl = new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            Text = "PDF'leri buraya sürükle-bırak\n(KASKO ve/veya TRAFİK)",
            ForeColor = Color.FromArgb(90, 100, 110),
            AllowDrop = true
        };
        _dropPanel.Controls.Add(dropLbl);
        _dropPanel.DragEnter += Drop_Enter; _dropPanel.DragDrop += Drop_Drop;
        dropLbl.DragEnter += Drop_Enter; dropLbl.DragDrop += Drop_Drop;

        // ── Satır 3: önizleme ───────────────────────────────
        _onizleme.Dock = DockStyle.Fill;
        _onizleme.ReadOnly = true;
        _onizleme.Font = new Font("Consolas", 9f);
        _onizleme.BackColor = Color.White;
        _onizleme.Text = "Nasıl kullanılır:\r\n 1) Hedef Excel'i seç (TECDİT dosyası).\r\n 2) İzlenecek klasörü seç (TEKLIF-GELEN) ve 'İzle'yi işaretle.\r\n 3) Open'da 'Dışa Al' → kaydet. Önizleme buraya kendiliğinden gelir.\r\n 4) Kontrol et, 'Excel'e Yaz'a bas.\r\n\r\n(Dilersen PDF'leri yukarıya sürükleyebilirsin de.)";

        // ── Satır 4: alt bar ────────────────────────────────
        var alt = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 8, 0, 0) };
        alt.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        alt.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        alt.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        _durum.AutoSize = true; _durum.Anchor = AnchorStyles.Left; _durum.Margin = new Padding(0, 8, 0, 0);
        _durum.Text = "Hazır. Önce hedef Excel'i seç, sonra PDF'leri sürükle.";
        _temizleBtn.Text = "Temizle"; _temizleBtn.AutoSize = true; _temizleBtn.Margin = new Padding(0, 0, 6, 0);
        _temizleBtn.Click += (_, _) => Temizle();
        _yazBtn.Text = "Excel'e Yaz"; _yazBtn.AutoSize = true; _yazBtn.Enabled = false;
        _yazBtn.BackColor = Color.FromArgb(40, 167, 69); _yazBtn.ForeColor = Color.White;
        _yazBtn.FlatStyle = FlatStyle.Flat; _yazBtn.Padding = new Padding(12, 4, 12, 4);
        _yazBtn.Click += (_, _) => Yaz();
        alt.Controls.Add(_durum, 0, 0);
        alt.Controls.Add(_temizleBtn, 1, 0);
        alt.Controls.Add(_yazBtn, 2, 0);

        root.Controls.Add(excelRow, 0, 0);
        root.Controls.Add(klasorRow, 0, 1);
        root.Controls.Add(_dropPanel, 0, 2);
        root.Controls.Add(_onizleme, 0, 3);
        root.Controls.Add(alt, 0, 4);
        Controls.Add(root);

        if (_config.IzlemeAcik) IzlemeBaslat();
    }

    private void ExcelSec()
    {
        using var d = new OpenFileDialog { Filter = "Excel dosyası|*.xlsx", Title = "TECDİT Excel'ini seç" };
        if (!string.IsNullOrEmpty(_config.ExcelYolu))
        {
            try { d.InitialDirectory = Path.GetDirectoryName(_config.ExcelYolu); } catch { }
        }
        if (d.ShowDialog(this) == DialogResult.OK)
        {
            _config.ExcelYolu = d.FileName;
            _config.Kaydet();
            _excelKutu.Text = d.FileName;
            _durum.Text = "Excel seçildi. PDF'leri sürükle.";
        }
    }

    private void KlasorSec()
    {
        using var d = new FolderBrowserDialog { Description = "Open'ın PDF'leri kaydettiği klasörü seç (TEKLIF-GELEN)" };
        if (!string.IsNullOrEmpty(_config.IzlenecekKlasor) && Directory.Exists(_config.IzlenecekKlasor))
            d.SelectedPath = _config.IzlenecekKlasor;

        if (d.ShowDialog(this) != DialogResult.OK) return;

        _config.IzlenecekKlasor = d.SelectedPath;
        _config.Kaydet();
        _klasorKutu.Text = d.SelectedPath;
        if (_izleKutu.Checked) IzlemeBaslat();
        else _durum.Text = "Klasör seçildi. İzlemek için 'İzle'yi işaretle.";
    }

    private void IzlemeDegisti()
    {
        _config.IzlemeAcik = _izleKutu.Checked;
        _config.Kaydet();
        if (_izleKutu.Checked) IzlemeBaslat();
        else { _izleyici?.Dur(); _durum.Text = "İzleme kapalı."; }
    }

    private void IzlemeBaslat()
    {
        if (string.IsNullOrWhiteSpace(_config.IzlenecekKlasor) || !Directory.Exists(_config.IzlenecekKlasor))
        {
            _durum.Text = "İzlenecek klasör seçili değil (veya bulunamadı).";
            return;
        }
        _izleyici ??= new KlasorIzleyici(PdfGeldi);
        _izleyici.Basla(_config.IzlenecekKlasor);
        _durum.Text = _izleyici.Calisiyor
            ? $"İzleniyor: {_config.IzlenecekKlasor}  —  Open'dan 'Dışa Al' yapman yeterli."
            : "İzleme başlatılamadı.";
    }

    /// <summary>Arka plan thread'inden gelir → UI'ya marshal et.</summary>
    private void PdfGeldi(string yol)
    {
        if (IsDisposed || !IsHandleCreated) return;
        try
        {
            BeginInvoke(() =>
            {
                Onizle(new List<string> { yol });
                GorevCubuguYanipSon();
            });
        }
        catch (ObjectDisposedException) { /* pencere kapandı */ }
    }

    private void Tara()
    {
        if (string.IsNullOrWhiteSpace(_config.IzlenecekKlasor) || !Directory.Exists(_config.IzlenecekKlasor))
        {
            _durum.Text = "Önce izlenecek klasörü seç.";
            return;
        }
        var pdfler = Directory.GetFiles(_config.IzlenecekKlasor, DosyaAdi.Filtre).ToList();
        if (pdfler.Count == 0) { _durum.Text = "Klasörde işlenecek PDF yok."; return; }
        Onizle(pdfler);
    }

    private void Drop_Enter(object? s, DragEventArgs e)
    {
        if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true)
        {
            var files = (string[])e.Data.GetData(DataFormats.FileDrop)!;
            if (files.Any(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)))
                e.Effect = DragDropEffects.Copy;
        }
    }

    private void Drop_Drop(object? s, DragEventArgs e)
    {
        var files = ((string[])e.Data!.GetData(DataFormats.FileDrop)!)
            .Where(f => f.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (files.Count > 0) Onizle(files);
    }

    private void Onizle(List<string> pdfler)
    {
        var sb = new StringBuilder();

        // Yeni gelen PDF'leri mevcut listeye EKLE (aynı ad varsa güncelle).
        // Böylece KASKO ve TRAFİK'i ayrı ayrı da sürükleyebilirsin, ikisi birikir.
        foreach (var pdf in pdfler)
        {
            string ad = Path.GetFileName(pdf);
            try
            {
                var t = TeklifParser.Parse(pdf);
                // Parser branşı bulamazsa dosya adındaki branş devreye girer.
                if (t.Brans == Brans.Bilinmiyor) t.Brans = DosyaAdi.Coz(pdf).Brans;
                var mert = MertFormatter.Uret(t);
                _bekleyen.RemoveAll(x => string.Equals(x.pdfAdi, ad, StringComparison.OrdinalIgnoreCase));
                _bekleyen.Add((t, mert, ad));
            }
            catch (Exception ex)
            {
                sb.AppendLine($"✗ {ad} okunamadı: {ex.Message}");
            }
        }

        List<YazimSonucu>? dry = null;
        if (_bekleyen.Count > 0 && !string.IsNullOrEmpty(_config.ExcelYolu) && File.Exists(_config.ExcelYolu))
        {
            try { dry = ExcelYazici.Isle(_config.ExcelYolu, _bekleyen, false); }
            catch (Exception ex) { sb.AppendLine($"⚠ Excel okunamadı: {ex.Message}\r\n"); }
        }

        for (int i = 0; i < _bekleyen.Count; i++)
        {
            var (t, mert, ad) = _bekleyen[i];
            sb.AppendLine($"── {ad}   [{t.Brans}]");
            sb.AppendLine($"   Müşteri : {t.MusteriAdi}    Plaka : {t.Plaka}");
            sb.AppendLine($"   MERT     : {mert}");
            if (dry != null && i < dry.Count)
            {
                var dd = dry[i];
                if (dd.GecmisTarihi is not null)
                    sb.AppendLine($"   ⚠ Bu plaka için {dd.GecmisTarihi} tarihinde zaten liste yazılmış.");
                if (dd.Atlandi)
                    sb.AppendLine($"   → {dd.HedefHucre} : ⏭ zaten yazılmış (atlandı)");
                else if (dd.Bulundu)
                    sb.AppendLine($"   → Hedef  : {dd.HedefHucre}   (mevcut not korunur, sonuna eklenir)");
                else
                    sb.AppendLine($"   → {(dd.Coklu || dd.AdUyusmazligi ? "⚠ " : "")}{dd.Mesaj}");
            }
            sb.AppendLine();
        }

        bool yazilabilir = dry != null && dry.Any(x => x.Bulundu && !x.Atlandi);
        _yazBtn.Enabled = yazilabilir;
        _durum.Text = yazilabilir
            ? "Önizleme hazır. Yazmak için 'Excel'e Yaz'a bas."
            : string.IsNullOrEmpty(_config.ExcelYolu) ? "Önce hedef Excel'i seç."
            : (dry != null && dry.Any(x => x.Atlandi)) ? "Bu liste(ler) zaten yazılmış."
            : "Eşleşen satır bulunamadı.";
        _onizleme.Text = sb.ToString();
    }

    private void Temizle()
    {
        _bekleyen.Clear();
        _yazBtn.Enabled = false;
        _onizleme.Text = "Liste temizlendi. PDF'leri sürükle.";
        _durum.Text = "Hazır.";
    }

    private void Yaz()
    {
        if (_bekleyen.Count == 0 || string.IsNullOrEmpty(_config.ExcelYolu)) return;
        if (!GecmisOnayi()) return;

        _yazBtn.Enabled = false;
        _durum.Text = "Yazılıyor…";
        Application.DoEvents();
        try
        {
            var yazilan = _bekleyen.ToList();   // Isle sonuçları girdi sırasıyla döner
            var res = ExcelYazici.Isle(_config.ExcelYolu, yazilan, true);
            var sb = new StringBuilder();
            foreach (var d in res)
                sb.AppendLine(d.Bulundu ? $"✔ {d.PdfAdi} → {d.HedefHucre} : {d.Mesaj}"
                                        : $"✗ {d.PdfAdi} : {d.Mesaj}");
            _onizleme.Text = sb.ToString();

            for (int i = 0; i < res.Count && i < yazilan.Count; i++)
                if (res[i].Bulundu && !res[i].Atlandi)
                    YazimGecmisi.Kaydet(yazilan[i].t.Plaka, yazilan[i].t.Brans, yazilan[i].mert);

            int ok = res.Count(x => x.Bulundu && !x.Atlandi);
            string mod = res.Count > 0 ? res[0].KaynakMod : "";
            _durum.Text = $"Tamamlandı: {ok}/{res.Count} yazıldı.  (mod: {mod})";
            _bekleyen.Clear();
        }
        catch (Exception ex)
        {
            _durum.Text = "Hata: " + ex.Message;
            MessageBox.Show(this, ex.Message, "Yazma hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>Daha önce yazılmış plaka+branş varsa kullanıcıya sorar. Araç hücredeki eski bloğu
    /// değiştiremediği için ikinci yazım hücrede iki liste demek — kararı kullanıcı verir.</summary>
    private bool GecmisOnayi()
    {
        var tekrar = _bekleyen
            .Select(x => (x, g: YazimGecmisi.Bul(x.t.Plaka, x.t.Brans)))
            .Where(p => p.g is not null)
            .ToList();
        if (tekrar.Count == 0) return true;

        var sb = new StringBuilder("Bu liste(ler) daha önce yazılmış:\r\n\r\n");
        foreach (var (x, g) in tekrar)
            sb.AppendLine($"    {x.t.Plaka}  [{x.t.Brans}]  →  {g!.Tarih}");
        sb.AppendLine("\r\nAraç eski listeyi değiştiremez, yenisini hücrenin sonuna EKLER.");
        sb.AppendLine("Yani hücrede iki liste olur. Yine de eklensin mi?");

        return MessageBox.Show(this, sb.ToString(), "Daha önce yazılmış",
            MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) == DialogResult.Yes;
    }

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        _izleyici?.Dispose();
        base.OnFormClosing(e);
    }

    // ── Görev çubuğu bildirimi ──────────────────────────────
    // Odak ÇALINMAZ: kullanıcı iki export arasında Open'da çalışıyor, pencereyi öne getirmek araya girer.

    [StructLayout(LayoutKind.Sequential)]
    private struct FLASHWINFO
    {
        public uint cbSize;
        public IntPtr hwnd;
        public uint dwFlags;
        public uint uCount;
        public uint dwTimeout;
    }

    private const uint FLASHW_ALL = 3;
    private const uint FLASHW_TIMERNOFG = 12;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool FlashWindowEx(ref FLASHWINFO pwfi);

    private void GorevCubuguYanipSon()
    {
        try
        {
            var fi = new FLASHWINFO
            {
                cbSize = (uint)Marshal.SizeOf<FLASHWINFO>(),
                hwnd = Handle,
                dwFlags = FLASHW_ALL | FLASHW_TIMERNOFG,
                uCount = 3,
                dwTimeout = 0
            };
            FlashWindowEx(ref fi);
        }
        catch { /* bildirim kritik değil */ }
    }
}
