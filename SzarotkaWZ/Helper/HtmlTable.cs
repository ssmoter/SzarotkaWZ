using SzarotkaWZ.Models;

namespace SzarotkaWZ.Helper
{
    internal class HtmlTable
    {
        static string style = @"
<style>
    .centerText {
        text-align: center;
        font-family: ""Times New Roman"", Times, serif;
    }

    .widthTable {
        width: 99%;
        align-content: center;
        margin: auto;
    }

    .backLightGreen {
        background-color: rgb(234,241,221);
    }

    .backDarkGreen {
        background-color: rgb(214,227,188);
    }

    table, th {
        border-left: 1px solid black;
        border-bottom: 1px solid black;
        border-right: 1px solid black;
    }

    td {
        border-right: 1px solid black;
    }

    .borderBottom {
        border-bottom: 1px solid black;
    }

    .rightBorderNone {
        border-right: 0px solid white;
    }

</style>


";
        public static string StartHtml = $@"
<html>
{style}
<body>

";

        public static string EndHtml = @$"
</body>
</html>";

        static string ListOfProducts(List<Models.Produkt> produkts)
        {
            int empty = 4;
            string list = "";
            for (int i = 0; i < produkts.Count; i++)
            {
                if (produkts[i].Wartosc == "0" || string.IsNullOrWhiteSpace(produkts[i].Wartosc))
                    continue;
                list += @$"
            <tr class=""centerText"">
                <td class=""borderBottom centerText"">{produkts[i].KodTowaru}</td>
                <td class=""borderBottom centerText"">{produkts[i].Nazwa}</td>
                <td class=""borderBottom centerText"">{produkts[i].Zadysponowana}</td>
                <td class=""borderBottom centerText"">{produkts[i].Jm}</td>
                <td class=""borderBottom centerText"">{produkts[i].Wydana}</td>
                <td class=""borderBottom centerText"">{produkts[i].Cena}</td>
                <td class=""borderBottom centerText"" style=""border-right:none"">{produkts[i].Wartosc}</td>
            </tr>
";
                empty--;
            }
            if (empty > 0)
                for (int i = 0; i < empty; i++)
                {
                    list += @"
            <tr class=""centerText"">
                <td class=""borderBottom centerText"">&nbsp;</td>
                <td class=""borderBottom centerText""></td>
                <td class=""borderBottom centerText""></td>
                <td class=""borderBottom centerText""></td>
                <td class=""borderBottom centerText""></td>
                <td class=""borderBottom centerText""></td>
                <td class=""borderBottom centerText"" style=""border-right:none""></td>
            </tr>
";
                }


            return list;
        }

        public static string Created(Models.Wz wz)
        {
            var nabywca = Wz.SplitBold(wz.Nabywca);
            var odbiorca = Wz.SplitBold(wz.Odbiorca);

            string html = $@"
    <! –– 1 szereg ––>

    <table class=""widthTable "" style=""border-top:1px solid black;"">
        <tr>
            <td class=""centerText"">
                <span class=""centerText"">
{wz.Cukiernia}
                </span>
            </td>
            <td class=""borderRigth "">
                <p>
                    <b><span class=""centerText"">{nabywca.Item1}</span> </b>
                    <span class=""centerText"">{nabywca.Item2}</span>
                </p>
                <p>
                    <b><span class=""centerText"">{odbiorca.Item1}</span></b>
                    <span class=""centerText"">{odbiorca.Item2}</span>
                </p>

            </td>
            <td class=""backDarkGreen"">
                <div class=""centerText"">
                    <p>
                        <b><span class=""centerText"">WZ</span></b>
                    </p>
                    <div><span class=""centerText"">wydanie zewnętrzne</span></div>
                </div>
            </td>
            <td class=""backLightGreen"" style=""border-right:none"">
                <span class=""backLightGreen"">
                    <p><span style=""font-family: "" Times New Roman"", Times, serif;"">Data wystawienia:</span></p>
                    <p class=""centerText"">{wz.DataWystawienia}</p>
                </span>
            </td>
        </tr>
    </table>

    <! –– 2 szereg ––>
    <table class=""widthTable "">
        <tr class=""centerText backLightGreen borderBottom"">
            <td class=""centerText"" style=""border-bottom:1px solid black"">Środek transportu</td>
            <td class=""centerText"" style=""border-bottom:1px solid black"">Zamówienie</td>
            <td class=""centerText"" style=""border-bottom:1px solid black"">Przeznaczenie</td>
            <td class=""centerText"" style=""border-bottom:1px solid black"">Data wysyłki</td>
            <td class=""centerText"" style=""border-right:none; border-bottom:1px solid black"">Numer i data faktury - specyfikacji</td>
        </tr>
        <tr class=""centerText"">
            <td class=""centerText"">{(string.IsNullOrWhiteSpace(wz.WzDodatkowe.SrodekTransportu) ? "&nbsp;" : wz.WzDodatkowe.SrodekTransportu)}</td>
            <td class=""centerText"">{wz.WzDodatkowe.Zamowienie}</td>
            <td class=""centerText"">{wz.WzDodatkowe.Przeznaczenie}</td>
            <td class=""centerText"">{wz.WzDodatkowe.Przeznaczenie}</td>
            <td class=""centerText"" style=""border-right:none"">{wz.WzDodatkowe.NumerIDataFaktury}</td>
        </tr>
    </table>

    <! –– 3 szereg ––>
    <table class=""widthTable centerText tableBorder"">

        <tr class=""backDarkGreen"">
            <td rowspan=""2"" class=""borderBottom centerText"">Kod towaru/ materiału</td>
            <td rowspan=""2"" class=""borderBottom centerText"">Nazwa towaru/materiału/opakowania</td>

            <td colspan=""3"" class=""borderBottom centerText"">Ilość</td>

            <td rowspan=""2"" class=""borderBottom centerText"">Cena</td>
            <td rowspan=""2"" style=""border-right:none"" class=""borderBottom centerText"">Wartość</td>

        </tr>
        <tr class=""backDarkGreen"">
            <td class=""borderBottom centerText"">Zadysponowana</td>
            <td class=""borderBottom centerText"">j.m.</td>
            <td class=""borderBottom centerText"">Wydana</td>
        </tr>
        <! –– pętla ––>

{ListOfProducts(wz.Products)}

            <tr class=""centerText centerText"">
                <td class=""centerText"">&nbsp;</td>
                <td class=""centerText""></td>
                <td class=""centerText""></td>
                <td class=""centerText""></td>
                <td class=""centerText""></td>
                <td class=""centerText""></td>
                <td class=""centerText"" style=""border-right:none""></td>
            </tr>
    </table>

    <! –– 4 szereg ––>
    <table class=""widthTable centerText tableBorder"">

        <tr class=""backDarkGreen"">
            <td class=""borderBottom centerText"" rowspan=""2"">Wystawił</td>
            <td class=""borderBottom centerText "" rowspan=""2"">Zatwierdził</td>

            <td colspan=""4"" class=""borderBottom centerText"" style=""border-right:none"">Wymienione ilości</td>

        </tr>
        <tr class=""backLightGreen"">
            <td class=""borderBottom""><span class=""centerText"">Wydał</span></td>
            <td class=""borderBottom""><span class=""centerText"">Data</span></td>
            <td class=""borderBottom""><span class=""centerText"">Odebrał</span></td>
            <td class=""borderBottom"" style=""border-right:none""><span class=""centerText"">Ewidencja ilościowo - wartościowa</span></td>
        </tr>
        <tr>
            <td><span class=""centerText"">{wz.WzPodsumowanie.Wystawil}</span></td>
            <td><span class=""centerText"">{wz.WzPodsumowanie.Zatwierdzil}</span></td>
            <td><span class=""centerText"">{wz.WzPodsumowanie.Wydal}</span></td>
            <td><span class=""centerText"">{wz.WzPodsumowanie.Data}</span></td>
            <td><span class=""centerText"">{wz.WzPodsumowanie.Odebral}</span></td>
            <td style=""border-right:none""><b><span class=""centerText"">{wz.WzPodsumowanie.Ewidencja}</span></b></td>

        </tr>
    </table>


";


            return html;
        }
    }
}
