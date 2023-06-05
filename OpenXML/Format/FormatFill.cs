using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.OpenXML.Format;

internal static class FormatFill
{
        public static Fills Get() =>
        new Fills(
            // 0 no Fill
            new Fill(new PatternFill() { PatternType = PatternValues.None }),
            // 1 Grey Fill
            new Fill(new PatternFill() { PatternType = PatternValues.Gray125 }),
            // 2 Red
            new Fill(new PatternFill(
                new ForegroundColor { Rgb = new HexBinaryValue() { Value = "F897C1" } })
            {
                PatternType = PatternValues.Solid
            }),
            // 3 Blue
            new Fill(new PatternFill(
            new ForegroundColor { Rgb = new HexBinaryValue() { Value = "D4D9FD" } })
            {
                PatternType = PatternValues.Solid
            })
        );
}