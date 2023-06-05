using ExcelExamples.OpenXML.Format;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExamples.Domain;
using ExcelExamples.OpenXML.DocSheets;
using NLog;
using System.Globalization;

namespace ExcelExamples.OpenXML;
public class OpenXML : IExcel
{
    private readonly Logger _logger = LogManager.GetCurrentClassLogger();

    public Task RunAsync(Film[] films)
    {
        _logger.Info("OpenXML run");
        var filePath = Path.Combine(AppContext.BaseDirectory, "Films.xlsx");
        using var document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

        worksheetPart.Worksheet = new Worksheet(new SheetData());
        WorkbookStylesPart wbsp = workbookPart.AddNewPart<WorkbookStylesPart>();


        // // Add styles
        wbsp.Stylesheet = new(
            FormatNumbering.Get(),
            FormatFont.Get(),
            FormatFill.Get(),
            FormatBorder.Get(),
            FormatCell.Get()
        );

        wbsp.Stylesheet.Save();

        var sheets = workbookPart.Workbook.AppendChild(new Sheets());
        uint sheetId = 1;
        var worksheet = AddWorksheet(document, sheets, "Films", sheetId++);

        var fs = new FilmSheet();
        fs.Fill(worksheet, films);

        workbookPart.Workbook.Save();

        return Task.CompletedTask;
    }

    private Worksheet AddWorksheet(SpreadsheetDocument? document, Sheets sheets, string name, uint sheetId)
    {
        var workbookPart = document!.WorkbookPart;

        WorksheetPart worksheetPart = workbookPart!.AddNewPart<WorksheetPart>();
        SheetData sheetData = new();
        var worksheet = new Worksheet(sheetData);
        worksheetPart.Worksheet = worksheet;

        var sheet = new Sheet
        {
            Id = document.WorkbookPart!.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name
        };
        sheets.Append(sheet);

        return worksheet;
    }

    public static void InsertCell(Row row, int colNo, string val, ExcelStyles style)
    {
        var newCell = new Cell
        {
            CellReference = GetCellName(colNo, (int)row.RowIndex!.Value),
            StyleIndex = (uint)style,
            CellValue = new CellValue(val),
            DataType = new EnumValue<CellValues>(CellValues.String)
        };

        row.AppendChild(newCell);
    }

    public static void InsertCell(Row row, int colNo, decimal val, ExcelStyles style)
    {
        var newCell = new Cell
        {
            CellReference = GetCellName(colNo, (int)row.RowIndex!.Value),
            StyleIndex = (uint)style,
            CellValue = new CellValue(val),
            DataType = new EnumValue<CellValues>(CellValues.Number),
        };

        row.AppendChild(newCell);
    }

    public static void InsertCell(Row row, int colNo, double val, ExcelStyles style)
    {
        var newCell = new Cell
        {
            CellReference = GetCellName(colNo, (int)row.RowIndex!.Value),
            StyleIndex = (uint)style,
            CellValue = new CellValue(val),
            DataType = new EnumValue<CellValues>(CellValues.Number),
        };

        row.AppendChild(newCell);
    }

    private static void UpdateCell(Row row, int colNo, string val, CellValues type)
    {
        var reference = GetCellName(colNo, (int)row.RowIndex!.Value);
        var cell = row
            .Elements<Cell>()
            .FirstOrDefault(c => string.Compare(c.CellReference!.Value!, reference, true, CultureInfo.InvariantCulture) == 0);

        if (cell != null)
        {
            cell.CellValue = new CellValue(val);
            cell.DataType = new EnumValue<CellValues>(type);
        }
    }

    public static void AddColumns(Worksheet worksheet, int columnsNumber, double width)
    {
        var columns = worksheet.GetFirstChild<Columns>() ?? new Columns();

        for (uint i = 1; i <= columnsNumber; i++)
        {
            columns.Append(new Column() { Min = i, Max = i, Width = new DoubleValue(width), CustomWidth = true });
        }

        worksheet.InsertAt(columns, 0);
    }

    public static void AddColumns(Worksheet worksheet, double[] width)
    {
        var columns = worksheet.GetFirstChild<Columns>() ?? new Columns();

        for (uint i = 0; i < width.Length; i++)
        {
            columns.Append(new Column() { Min = i + 1, Max = i + 1, Width = new DoubleValue(width[i]), CustomWidth = true });
        }

        worksheet.InsertAt(columns, 0);
    }

    public static Row AddRow(SheetData sheetData, uint rowIndex, double height = 15)
    {
        var row = new Row { RowIndex = rowIndex, Height = height };
        sheetData.Append(row);
        return row;
    }

    public static void MergeCells(Worksheet worksheet, int startColNo, int endColNo, int rowNo)
    {
        var totalCellName = GetCellName(startColNo, rowNo);
        var nextCellName = GetCellName(endColNo, rowNo);
        var mergeCellNames = $"{totalCellName}:{nextCellName}";

        var mergeCell = new MergeCell() { Reference = new StringValue(mergeCellNames) };

        var mergeCells = worksheet.GetFirstChild<MergeCells>();
        if (mergeCells == null)
        {
            mergeCells = new MergeCells();
            worksheet.InsertAfter(mergeCells, GetSheetData(worksheet));
        }
        mergeCells.Append(mergeCell);
    }

    public static SheetData GetSheetData(Worksheet worksheet)
    {
        return worksheet.Elements<SheetData>().First();
    }

    public static string GetCellName(int colNo, int rowNo)
    {
        int div = colNo + 1;
        string colLetter = String.Empty;
        int mod = 0;

        while (div > 0)
        {
            mod = (div - 1) % 26;
            colLetter = (char)(65 + mod) + colLetter;
            div = (int)((div - mod) / 26);
        }

        return $"{colLetter}{rowNo}";
    }
}