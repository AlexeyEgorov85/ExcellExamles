using ExcelExamples.Domain;
using NLog;

namespace ExcelExamples.ExcelDataReader;

public class ExcelDataReader : IExcel
{
    private readonly Logger _logger = LogManager.GetCurrentClassLogger();

    public Task RunAsync()
    {
        _logger.Info("OpenXML run");
        return Task.CompletedTask;
    }
}