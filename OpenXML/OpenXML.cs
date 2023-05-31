using ExcelExamples.Domain;

namespace ExcelExamples.OpenXML;
public class OpenXML: IExcel
{
    public Task RunAsync()
    {
        Console.WriteLine("OpenXML run");
        return Task.CompletedTask;
    }
}