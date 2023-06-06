using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.OpenXmlExample.Format;

internal static class FormatBorder
{
    public static Borders Get() =>
    new Borders(
        // 0
        new Border(
            new LeftBorder(),
            new RightBorder(),
            new TopBorder(),
            new BottomBorder(),
            new DiagonalBorder()),
        // 1
        new Border(
            new LeftBorder(new Color() { Auto = true })
            { Style = BorderStyleValues.Thin },
            new RightBorder(new Color() { Indexed = (UInt32Value)64U })
            { Style = BorderStyleValues.Thin },
            new TopBorder(new Color() { Auto = true })
            { Style = BorderStyleValues.Thin },
            new BottomBorder(new Color() { Indexed = (UInt32Value)64U })
            { Style = BorderStyleValues.Thin },
            new DiagonalBorder()),
        // 2
        new Border(
            new LeftBorder(new Color() { Auto = true })
            { Style = BorderStyleValues.Thin },
            new RightBorder(new Color() { Indexed = (UInt32Value)64U })
            { Style = BorderStyleValues.Thin },
            new TopBorder(),
            new BottomBorder(),
            new DiagonalBorder())

    );
}