namespace TeklifOtomasyon;

public enum Brans { Bilinmiyor, Kasko, Trafik }

/// <summary>Bir sigorta şirketinin tek bir teklifi.</summary>
public record SirketTeklif(string HamAd, decimal Fiyat);

/// <summary>Bir teklif PDF'inden çıkarılan tüm bilgiler.</summary>
public class TeklifCalismasi
{
    public Brans Brans { get; set; } = Brans.Bilinmiyor;
    public string? MusteriAdi { get; set; }
    public string? Plaka { get; set; }
    public string? Tc { get; set; }
    public string? BelgeSeri { get; set; }
    public string? Sasi { get; set; }
    public string? Motor { get; set; }
    public string? MarkaTip { get; set; }   // örn: "TOYOTA YARIS 1.33 STYLE"
    public string? ModelYili { get; set; }
    public string? KaskoDegeri { get; set; }
    public string? Hazirlayan { get; set; }
    public string? Tarih { get; set; }
    public List<SirketTeklif> Teklifler { get; } = new();
}
