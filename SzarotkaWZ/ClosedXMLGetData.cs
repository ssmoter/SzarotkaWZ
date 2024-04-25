
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace SzarotkaWZ
{
    internal class ClosedXMLGetData
    {
        public static Models.Wz GetModel(int numberOfSheet, string path)
        {
            Models.Wz wz = new();

            // Otwórz plik Excela
            try
            {
                using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false);
                if (spreadsheetDocument.WorkbookPart is null)
                {
                    throw new ArgumentNullException("spreadsheetDocument.WorkbookPart Was null");
                }

                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.ElementAt(numberOfSheet); // Zakładając, że arkusz 2 to drugi arkusz
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Odczytaj dane z komórki
                wz.Cukiernia = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 37));
                wz.Nabywca = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 38));
                wz.Odbiorca = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 39));
                wz.DataWystawienia = PartseToDate(GetCellValue(workbookPart, GetCell(worksheetPart, "K", 40)));
                wz.WzDodatkowe.SrodekTransportu = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 41));
                wz.WzDodatkowe.Zamowienie = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 42));
                wz.WzDodatkowe.Przeznaczenie = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 43));
                wz.WzDodatkowe.DataWysylki = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 44));
                wz.WzDodatkowe.NumerIDataFaktury = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 45));
                wz.WzPodsumowanie.Wystawil = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 46));
                wz.WzPodsumowanie.Zatwierdzil = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 47));
                wz.WzPodsumowanie.Wydal = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 48));
                wz.WzPodsumowanie.Data = PartseToDate(GetCellValue(workbookPart, GetCell(worksheetPart, "K", 49)));
                wz.WzPodsumowanie.Odebral = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 50));
                wz.WzPodsumowanie.Ewidencja = GetCellValue(workbookPart, GetCell(worksheetPart, "K", 51));


                bool readProduct = true;
                uint row = 38;
                while (readProduct)
                {
                    wz.Products.Add(GetProdukt(workbookPart, worksheetPart, row));

                    var lastProduct = wz.Products.LastOrDefault();

                    if (lastProduct is not null)
                        if (string.IsNullOrWhiteSpace(lastProduct.Nazwa))
                            readProduct = false;

                    row++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error wiadomość{0}{1}{2}Error Stack Trace{3}{4}"
                    , Environment.NewLine
                    , ex.Message
                    , Environment.NewLine
                    , Environment.NewLine
                    , ex.StackTrace);
            }
            return wz;
        }

        static Models.Produkt GetProdukt(WorkbookPart workbookPart, WorksheetPart worksheetPart, uint row)
        {
            Models.Produkt produkt = new();
            try
            {
                produkt.KodTowaru = GetCellValue(workbookPart, GetCell(worksheetPart, "C", row));
                produkt.Nazwa = GetCellValue(workbookPart, GetCell(worksheetPart, "D", row));
                produkt.Cena = GetCellValue(workbookPart, GetCell(worksheetPart, "E", row));
                produkt.Jm = GetCellValue(workbookPart, GetCell(worksheetPart, "F", row));
                produkt.Zadysponowana = GetCellValue(workbookPart, GetCell(worksheetPart, "G", row));
                produkt.Wydana = GetCellValue(workbookPart, GetCell(worksheetPart, "H", row));
                produkt.Wartosc = GetCellValue(workbookPart, GetCell(worksheetPart, "I", row));
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error wiadomość{0}{1}{2}Error Stack Trace{3}{4}"
                    , Environment.NewLine
                    , ex.Message
                    , Environment.NewLine
                    , Environment.NewLine
                    , ex.StackTrace);
            }
            return produkt;
        }








        private static string ParseAsString(string value)
        {
            if (value == "0")
                return "";
            return value;
        }


        private static string PartseToDate(string value)
        {
            if (double.TryParse(value, out double dd))
            {
                var date = DateTime.FromOADate(dd);
                return date.ToString("dd.MM.yyyy");
            }
            return "";
        }


        // Metoda pomocnicza do pobrania konkretnej komórki
        private static Cell GetCell(WorksheetPart worksheetPart, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheetPart, rowIndex);
            if (row == null)
                return null;
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == columnName + rowIndex);
        }

        // Metoda pomocnicza do pobrania konkretnej wiersza
        private static Row GetRow(WorksheetPart worksheetPart, uint rowIndex)
        {
            return worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        // Metoda pomocnicza do odczytu wartości komórki (uwzględniająca stringi)
        private static string GetCellValue(WorkbookPart workbookPart, Cell cell)
        {
            if (cell is null)
            {
                return "";
            }

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int id = int.Parse(cell.InnerText);
                SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                return item.Text?.Text ?? item.InnerText ?? item.InnerXml;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(cell?.CellValue?.Text))
                    return "";

                return ParseAsString(cell?.CellValue?.Text);
            }
        }
    }
}
