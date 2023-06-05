using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.OpenXML.Format;

internal static class FormatFont
{
    private const int FontSize = 12;
    private const string FontName = "Times New Roman";

    public static Fonts Get() =>
        new Fonts(
            // 0 
            new Font( 
                new FontSize() { Val = FontSize },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = FontName }),
            // 1 - bold
            new Font(
                new Bold(),
                new FontSize() { Val = FontSize },
                new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                new FontName() { Val = FontName }),
            // 2 - italic
            new Font(
            new Italic(),
            new FontSize() { Val = FontSize },
            new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
            new FontName() { Val = FontName }),
            // 3 - bold Red
            new Font(
                new Bold(),
                new FontSize() { Val = FontSize },
                new Color() { Rgb = new HexBinaryValue() { Value = "EE0000" } },
                new FontName() { Val = FontName })
        );
}