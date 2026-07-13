using System.Text;

namespace TeklifOtomasyon;

public class MainForm : Form
{
    private readonly Config _config = Config.Yukle();
    private readonly TextBox _excelKutu = new();
    private readonly Panel _dropPanel = new();
    private readonly RichTextBox _onizleme = new();
    private readonly Button _yazBtn = new();
    private readonly Button _temizleBtn = new();
    private readonly Label _durum = new();
    private List<(TeklifCalismasi t, string mert, string pdfAdi)> _bekleyen = new();

    public MainForm()
    {
        Text = "Teklif Otomasyonu";
        Width = 780; Height = 620; MinimumSize = new Size(640, 480);
        StartPosition = FormStartPosition.CenterScreen;
        Font = new Font("Segoe UI", 9f);

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(12) };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 100));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // ── Satır 0: hedef Excel seçici ─────────────────────
        var excelRow = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 3, AutoSize = true, Margin = new Padding(0, 0, 0, 8) };
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

        // ── Satır 1: sürükle-bırak alanı ────────────────────
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

        // ── Satır 2: önizleme ───────────────────────────────
        _onizleme.Dock = DockStyle.Fill;
        _onizleme.ReadOnly = true;
        _onizleme.Font = new Font("Consolas", 9f);
        _onizleme.BackColor = Color.White;
        _onizleme.Text = "Nasıl kullanılır:\r\n 1) Hedef Excel'i seç (TECDİT dosyası).\r\n 2) Open'dan çıkan KASKO/TRAFİK PDF'lerini yukarıya sürükle.\r\n 3) Önizlemeyi kontrol et, 'Excel'e Yaz'a bas.";

        // ── Satır 3: alt bar ────────────────────────────────
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
        root.Controls.Add(_dropPanel, 0, 1);
        root.Controls.Add(_onizleme, 0, 2);
        root.Controls.Add(alt, 0, 3);
        Controls.Add(root);
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
                if (dd.Atlandi)
                    sb.AppendLine($"   → {dd.HedefHucre} : ⏭ zaten yazılmış (atlandı)");
                else if (dd.Bulundu)
                    sb.AppendLine($"   → Hedef  : {dd.HedefHucre}   (mevcut not korunur, sonuna eklenir)");
                else
                    sb.AppendLine($"   → {(dd.Coklu ? "⚠ " : "")}{dd.Mesaj}");
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
        _yazBtn.Enabled = false;
        _durum.Text = "Yazılıyor…";
        Application.DoEvents();
        try
        {
            var res = ExcelYazici.Isle(_config.ExcelYolu, _bekleyen, true);
            var sb = new StringBuilder();
            foreach (var d in res)
                sb.AppendLine(d.Bulundu ? $"✔ {d.PdfAdi} → {d.HedefHucre} : {d.Mesaj}"
                                        : $"✗ {d.PdfAdi} : {d.Mesaj}");
            _onizleme.Text = sb.ToString();
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
}
