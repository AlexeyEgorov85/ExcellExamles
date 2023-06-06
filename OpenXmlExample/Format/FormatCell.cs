using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.OpenXmlExample.Format;

internal static class FormatCell
{
    private static Alignment AlignmentLeft =>
        new()
        {
            Horizontal = HorizontalAlignmentValues.Left,
            Vertical = VerticalAlignmentValues.Center
        };

    private static Alignment AlignmentLeftWrap =>
        new()
        {
            Horizontal = HorizontalAlignmentValues.Left,
            Vertical = VerticalAlignmentValues.Center,
            WrapText = true
        };

    private static Alignment AlignmentRight =>
        new()
        {
            Horizontal = HorizontalAlignmentValues.Right,
            Vertical = VerticalAlignmentValues.Center
        };

    private static Alignment AlignmentCenter =>
        new()
        {
            Horizontal = HorizontalAlignmentValues.Center,
            Vertical = VerticalAlignmentValues.Center
        };

    private static Alignment AlignmentCenterWrap =>
        new()
        {
            Horizontal = HorizontalAlignmentValues.Center,
            Vertical = VerticalAlignmentValues.Center,
            WrapText = true
        };

    public static CellFormats Get() =>
        new(
            // ExcelStyles.Left
            Create(AlignmentLeft),
            // ExcelStyles.LeftBold
            Create(AlignmentLeft, fontId: 1),
            // ExcelStyles.LeftItalic
            Create(AlignmentLeft, fontId: 2),
            // ExcelStyles.LeftBorder
            Create(AlignmentLeft, borderId: 1),
            // ExcelStyles.LeftBorderWrap
            Create(AlignmentLeftWrap, borderId: 1),
            // ExcelStyles.LeftBoldBorder
            Create(AlignmentLeft, fontId: 1, borderId: 1),
            // ExcelStyles.CenterBorder
            Create(AlignmentCenter, borderId: 1),
            // ExcelStyles.CenterWrapBorder
            Create(AlignmentCenterWrap, borderId: 1),
            // ExcelStyles.CenterBoldBorder
            Create(AlignmentCenter, fontId: 1, borderId: 1),
            // ExcelStyles.LeftBoldBorderCenterRed
            Create(AlignmentCenter, fontId: 1, fillId: 2, borderId: 1),
            // ExcelStyles.RightBorder
            Create(AlignmentRight, borderId: 1),
            // ExcelStyles.RightBorderRed
            Create(AlignmentRight, fillId: 2, borderId: 1),
            // ExcelStyles.CenterBorderTextRed
            Create(AlignmentRight, fontId: 3, borderId: 1),
            // ExcelStyles.RightBorderTable
            Create(AlignmentRight, borderId: 2),
            // ExcelStyles.CenterBoldBorderBlue
            Create(AlignmentCenter, fontId: 1, fillId: 3, borderId: 1),
            // ExcelStyles.CenterBorderNumber0_00
            Create(AlignmentCenter, borderId: 1, numberFormatId: 165),
            // ExcelStyles.CenterBorderRedNumber0_00
            Create(AlignmentCenter, fillId: 2, borderId: 1, numberFormatId: 165)
        );

    private static CellFormat Create(Alignment alignment, uint fontId = 0, uint fillId = 0, uint borderId = 0, uint numberFormatId = 164) =>
        new CellFormat()
        {
            Alignment = alignment,
            FontId = fontId,
            ApplyFont = true,
            FillId = fillId,
            ApplyFill = true,
            BorderId = borderId,
            ApplyBorder = true,
            NumberFormatId = numberFormatId,
            ApplyNumberFormat = true
        };
}