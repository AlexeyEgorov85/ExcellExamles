using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.OpenXmlExample.Format;

internal static class FormatNumbering
{
    public static NumberingFormats Get()
    {
        uint index = FormatConstants.NumberingFormatsDefault;
        var formats = new NumberingFormats();
        // 164
        formats.AppendChild(new NumberingFormat()
            {
                FormatCode = "General",
                NumberFormatId = index++
            });
        // 165
        formats.AppendChild(new NumberingFormat()
        {
            FormatCode = "0.00",
            NumberFormatId = index++
        });

        formats.Count = (uint)formats.ChildElements.Count;

        return formats;
    }
}