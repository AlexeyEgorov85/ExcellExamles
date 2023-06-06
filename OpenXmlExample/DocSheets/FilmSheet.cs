using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExamples.Domain;

namespace ExcelExamples.OpenXmlExample.DocSheets;

internal class FilmSheet
{
    public void Fill(Worksheet worksheet, IEnumerable<Film> films)
    {
        var sheetData = worksheet.GetFirstChild<SheetData>()!;

        OpenXmlExample.AddColumns(worksheet, new double[] { 30, 100, 10});

        uint rowIndex = 1;

        var row = OpenXmlExample.AddRow(sheetData, rowIndex++);
        OpenXmlExample.InsertCell(row, 0, "Name", ExcelStyles.CenterBorder);
        OpenXmlExample.InsertCell(row, 1, "Description", ExcelStyles.CenterBorder);
        OpenXmlExample.InsertCell(row, 2, "Rate", ExcelStyles.CenterBorder);

        foreach (var film in films)
        {
            row = OpenXmlExample.AddRow(sheetData, rowIndex++);
            OpenXmlExample.InsertCell(row, 0, film.Name, ExcelStyles.CenterBorder);
            OpenXmlExample.InsertCell(row, 1, film.Description, ExcelStyles.CenterWrapBorder);
            var rateStyle = film.rate < 8 
                ? ExcelStyles.CenterBorderRedNumber0_00 
                : ExcelStyles.CenterBorderNumber0_00;
            OpenXmlExample.InsertCell(row, 2, film.rate, rateStyle);
        }
    }
}