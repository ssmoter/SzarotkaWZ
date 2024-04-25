using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

using SzarotkaWZ.Models;

namespace SzarotkaWZ
{
    internal class OpenXmlCreatedDocx
    {

        public static void CreateDocxWithTables(string path, List<Wz> wzs)
        {
            using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
            // Add a main document part. 
            MainDocumentPart mainPart = doc.AddMainDocumentPart();

            // Create the document structure and add some text.
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            var section = new SectionProperties
            {
                InnerXml = @"<w:pgMar w:top=""720"" w:right=""720"" w:bottom=""720"" w:left=""720"" w:header=""708"" w:footer=""708"" w:gutter=""0""/>"
            };
            body.AppendChild(section);

            for (int i = 0; i < wzs.Count; i++)
            {
                FillTable(body, wzs[i]);
                if (i == wzs.Count - 1)
                {
                    break;
                }
                body.AppendChild(new Paragraph(new Text("")));
                body.AppendChild(new Paragraph(new Text("")));
                body.AppendChild(new Paragraph(new Text("")));

            }
        }

        static Table CreatedTable()
        {
            Table table = new();

            TableProperties props = new(
                new TableBorders(
                new TopBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 6
                },
                new BottomBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 6
                },
                new LeftBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 6
                },
                new RightBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 6
                },
                new InsideHorizontalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 6
                },
                new InsideVerticalBorder
                {
                    Val = new EnumValue<BorderValues>(BorderValues.Single),
                    Size = 6
                }))
            {
                TableWidth = new TableWidth() { Width = "100", Type = TableWidthUnitValues.Auto }
            };
            table.AppendChild<TableProperties>(props);
            return table;
        }

        static void FillTable(Body body, Wz wz)
        {
            var table = CreatedTable();
            var tr = new TableRow();

            #region 1-Row

            var tc = new TableCell();
            var cukiernia = Wz.SplitStamp(wz.Cukiernia);
            tc.Append(new Paragraph(new Run(new Text(cukiernia.Item1)) { RunProperties = new RunProperties { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new Paragraph(new Run(new Text(cukiernia.Item2)) { RunProperties = new RunProperties { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 2 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            tc = new TableCell();
            var nabywca = Wz.SplitBold(wz.Nabywca);
            var odbiorca = Wz.SplitBold(wz.Odbiorca);
            tc.Append(new Paragraph(new Run(new Text(nabywca.Item1)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman, Bold = new Bold() { Val = new OnOffValue(true) } } }, new Run(new Text(nabywca.Item2)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { });
            tc.Append(new Paragraph(new Run(new Text(odbiorca.Item1)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman, Bold = new Bold() { Val = new OnOffValue(true) } } }, new Run(new Text(odbiorca.Item2)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 3 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("WZ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(22), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new Paragraph(new Run(new Text("wydanie zewnętrzne")) { RunProperties = new RunProperties() { FontSize = SetFontSize(16), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2.99 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);



            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Data wystawienia:")) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { });
            tc.Append(new Paragraph(new Run(new Text(wz.DataWystawienia)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{3.82 * 3600}" }));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            table.Append(tr);
            #endregion

            #region 2-Row
            tr = new TableRow();

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Środek transportu")) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Zamówienie")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2.7 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Przeznaczenie")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2.5 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Data wysyłki")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{1.9 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Numer i data faktury - specyfikacji")) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 3 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{9.2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            table.Append(tr);


            tr = new TableRow();

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzDodatkowe.SrodekTransportu)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzDodatkowe.Zamowienie)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2.7 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzDodatkowe.Przeznaczenie)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2.5 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzDodatkowe.DataWysylki)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{1.9 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzDodatkowe.NumerIDataFaktury)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 3 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{9.2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            table.Append(tr);


            #endregion

            #region 3-row

            tr = new TableRow();

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Kod towaru/ materiału")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Nazwa towaru/materiału/opakowania")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 1 }, new VerticalMerge() { Val = MergedCellValues.Restart },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Ilość")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 3 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Cena")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Wartość")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            table.Append(tr);

            #region 3.5-row
            tr = new TableRow();

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Zadysponowana")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("j.m.")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Wydana")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);
            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            table.Append(tr);

            #endregion

            #endregion

            #region Produkty
            int empty = 5;

            for (int i = 0; i < wz.Products.Count; i++)
            {
                if (wz.Products[i].Wartosc == "0" 
                    || wz.Products[i].Wartosc == "" 
                    || wz.Products[i].Wartosc == " "
                    || wz.Products[i].Wartosc == "0 zł")
                    continue;

                empty--;
                tr = new TableRow
                {
                    InnerXml = @"<w:trHeight w:hRule=""exact"" w:val=""255""/>"
                };

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].KodTowaru)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].Nazwa)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].Zadysponowana)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].Jm)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].Wydana)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].Cena)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                tc = new TableCell();
                tc.Append(new Paragraph(new Run(new Text(wz.Products[i].Wartosc)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                tc.Append(new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                tr.Append(tc);

                table.Append(tr);
            }
            if (empty > 0)
                for (int i = 0; i < empty; i++)
                {
                    tr = new TableRow
                    {
                        InnerXml = @"<w:trHeight w:hRule=""exact"" w:val=""255""/>"
                    };

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    tc = new TableCell();
                    tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
                    tc.Append(new TableCellProperties(
                        new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{2 * 3600}" }));
                    tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
                    tr.Append(tc);

                    table.Append(tr);
                }


            #endregion

            #region 5-row podsumowanie

            tr = new TableRow
            {
                InnerXml = @"<w:trHeight w:hRule=""exact"" w:val=""255""/>"
            };

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Wystawił")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Zatwierdził")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Wymienione ilości")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 5 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = Green, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            table.Append(tr);


            #region 5.3 naglowki

            tr = new()
            {
                InnerXml = @"<w:trHeight w:hRule=""exact"" w:val=""255""/>"
            };

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzPodsumowanie.Wystawil)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzPodsumowanie.Zatwierdzil)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Restart },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Wydał")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Data")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Odebrał")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text("Ewidencja ilościowo - wartościowa")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 2 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tc.Append(new TableCellProperties(new Shading() { Color = "auto", Fill = LightGreen, Val = ShadingPatternValues.Clear }));
            tr.Append(tc);


            table.Append(tr);

            #endregion

            #region 5.6 Dane

            tr = new()
            {
                InnerXml = @"<w:trHeight w:hRule=""exact"" w:val=""850""/>"
            };

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(" ")) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new VerticalMerge() { Val = MergedCellValues.Continue },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzPodsumowanie.Wydal)) { RunProperties = new RunProperties() { FontSize = SetFontSize(20), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzPodsumowanie.Data)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzPodsumowanie.Odebral)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);

            tc = new TableCell();
            tc.Append(new Paragraph(new Run(new Text(wz.WzPodsumowanie.Ewidencja)) { RunProperties = new RunProperties() { FontSize = SetFontSize(18), RunFonts = RunFontsRoman, Bold = new Bold() { Val = new OnOffValue(true) } } }) { ParagraphProperties = ParagraphPropertiesCenter });
            tc.Append(new TableCellProperties(new GridSpan() { Val = 2 },
                new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = $"{5.49 * 3600}" }));
            tc.Append(new TableCellProperties(TableCellVerticalAlignmentCenter));
            tr.Append(tc);


            table.Append(tr);


            #endregion

            #endregion

            body.AppendChild(table);

        }

        static Justification JustificationCenter => new() { Val = JustificationValues.Center };
        static TableCellVerticalAlignment TableCellVerticalAlignmentCenter => new() { Val = TableVerticalAlignmentValues.Center };
        static ParagraphProperties ParagraphPropertiesCenter => new() { Justification = JustificationCenter };
        static FontSize SetFontSize(int size)
        {
            return new FontSize() { Val = size.ToString() };
        }
        static RunFonts RunFontsRoman => new() { Ascii = "Times New Roman" };

        static string Green => "D6E3BC";
        static string LightGreen => "EAF1DD";
    }
}