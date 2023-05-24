
namespace ExcelExamples;

public class OpenXML: IExcel
{
    public Task RunAsync()
    {
        Console.WriteLine("OpenXML run");
        return Task.CompletedTask;
    }
}