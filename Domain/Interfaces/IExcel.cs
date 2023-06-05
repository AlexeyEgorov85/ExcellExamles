
namespace ExcelExamples.Domain;

public interface IExcel
{
    Task RunAsync(Film[] films);
}