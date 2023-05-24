
namespace ExcelExamples;

public class ExcelDataReader : IExcel
{
    public Task RunAsync()
    {
        Console.WriteLine("ExcelDataReader run");
        return Task.CompletedTask;
    }
}