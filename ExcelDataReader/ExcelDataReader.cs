using ExcelExamples.Domain;

namespace ExcelExamples.ExcelDataReader;

public class ExcelDataReader : IExcel
{
    public Task RunAsync()
    {
        Console.WriteLine("ExcelDataReader run");
        return Task.CompletedTask;
    }
}