namespace SzarotkaWZ.Models
{

    internal class Wz
    {
        public string Cukiernia { get; set; } = "";
        public string Nabywca { get; set; } = "";
        public string Odbiorca { get; set; } = "";
        public string DataWystawienia { get; set; } = "";
        public WzZamowienie WzDodatkowe { get; set; } = new();
        public List<Produkt> Products { get; set; } = [];
        public WzPodsumowanie WzPodsumowanie { get; set; } = new();
        public static (string, string) SplitBold(ReadOnlySpan<char> value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] == ':')
                {
                    var first = value.Slice(0, i + 1).ToString();
                    var second = value.Slice(i + 2, value.Length - i - 2).ToString();
                    return (first, second);
                }
            }
            return ("", value.ToString());
        }
        public static (string, string) SplitStamp(ReadOnlySpan<char> value)
        {
            for (int i = 0; i < value.Length; i++)
            {
                if (value[i] == '(')
                {
                    var first = value.Slice(0, i);
                    var second = value.Slice(i, value.Length - i);
                    return (first.ToString(), second.ToString());
                }
            }
            return ("", value.ToString());
        }

    }

    internal class WzZamowienie
    {
        public string SrodekTransportu { get; set; } = "";
        public string Zamowienie { get; set; } = "";
        public string Przeznaczenie { get; set; } = "";
        public string DataWysylki { get; set; } = "";
        public string NumerIDataFaktury { get; set; } = "";
    }

    internal class WzPodsumowanie
    {
        public string Wystawil { get; set; } = "";
        public string Zatwierdzil { get; set; } = "";
        public string Wydal { get; set; } = "";
        public string Data { get; set; } = "";
        public string Odebral { get; set; } = "";
        public string Ewidencja { get; set; } = "";
    }

    internal class Produkt
    {
        public string KodTowaru { get; set; } = "";
        public string Nazwa { get; set; } = "";
        public string Zadysponowana { get; set; } = "";
        public string Jm { get; set; } = "";
        public string Wydana { get; set; } = "";
        public string Cena { get; set; } = "";
        public string Wartosc { get; set; } = "";
    }
}
