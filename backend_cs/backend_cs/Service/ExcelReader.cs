using ClosedXML.Excel;
using database;

namespace backend_cs.Service;

public class ExcelReader
{
    public static List<Entry> ReadExcel(string filename)
    {
        var workbook = new XLWorkbook(filename);
        var worksheet = workbook.Worksheet(1);
        var entries = new List<Entry>(worksheet.RowCount());

        var getText = (XLCellValue cell) =>
        {
            if (cell.IsBlank) return "";
            if (cell.IsText) return cell.GetText();
            return "";
        };

        var getDouble = (XLCellValue cell) =>
        {
            if (cell.IsBlank) return 0.0;
            if (cell.IsNumber) return cell.GetNumber();
            if (cell.IsText) return Double.Parse(cell.GetText());
            return 0.0;
        };

        var getInt = (XLCellValue cell) =>
        {
            if (cell.IsBlank) return 0;
            if (cell.IsNumber) return (int)cell.GetNumber();
            if (cell.IsText) return Int32.Parse(cell.GetText());
            return 0;
        };
        
        for (var r = 3; r < worksheet.RowCount(); r++)
        {
            var entry = new Entry();
            var row = worksheet.Row(r);
            entry.Id = getInt(row.Cell(1).Value);
            entry.Title = getText(row.Cell(2).Value);
            entry.Category = getText(row.Cell(3).Value);
            entry.Income = getDouble(row.Cell(4).Value);
            entry.Stock = getInt(row.Cell(5).Value);
            entries.Add(entry);
        }

        return entries;
    }
}