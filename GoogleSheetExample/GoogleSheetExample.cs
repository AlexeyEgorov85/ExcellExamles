using ExcelExamples.Domain;
using NLog;

namespace ExcelExamples.GoogleSheetExample;

public class GoogleSheetExample : IExcel
{
    private readonly Logger _logger = LogManager.GetCurrentClassLogger();

    public Task RunAsync(Film[] films)
    {
        _logger.Info("OpenXML run");
        return Task.CompletedTask;
    }
}