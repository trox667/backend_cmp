using database;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace backend_cs.Service;

public class ExcelReaderOpenXML
{
    private static List<string> GetSharedString(SharedStringTablePart? sharedStringTablePart)
    {
        List<string> sharedStrings = new List<string>();
        if (sharedStringTablePart != null)
        {
            foreach (SharedStringItem item in sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>())
            {
                sharedStrings.Add(item.InnerText);
            }
        }
        return sharedStrings;
    }
    
    public static List<Entry> ReadExcel(string filename)
    {
        // Open a SpreadsheetDocument based on a file path.
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filename, false))
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            if (workbookPart == null)
            {
                return [];
            }

            var worksheetPart = workbookPart.WorksheetParts.First();
            var worksheet = worksheetPart.Worksheet.Elements<SheetData>().First();
            List<string> sharedStrings = GetSharedString(workbookPart.SharedStringTablePart);



            var getValue = (Cell cell) =>
            {
                var value = cell.CellValue?.InnerText ?? string.Empty;
                if (cell.DataType != null)
                {
                    if (cell.DataType == CellValues.SharedString)
                    {
                        value = sharedStrings.ElementAt(Convert.ToInt32(value));
                    }
                    else if (cell.DataType == CellValues.Boolean)
                    {
                        switch (value)
                        {
                            case "0":
                                value = "FALSE";
                                break;
                            default:
                                value = "TRUE";
                                break;
                        }
                    }
                }

                return value;
            };

            var entries = new List<Entry>(worksheet.Elements<Row>().Count());
            foreach (var row in worksheet.Elements<Row>().Skip(2))
            {
                var entry = new Entry();
                var cells = row.Elements<Cell>().ToList();
                entry.Id = Convert.ToInt32(getValue(cells[0]));
                entry.Title = getValue(cells[1]);
                entry.Category = getValue(cells[2]);
                entry.Income = Convert.ToDouble(getValue(cells[3]));
                entry.Stock = Convert.ToInt32(getValue(cells[4]));
                entries.Add(entry);
            }

            return entries;
        }

        return [];
    }
}