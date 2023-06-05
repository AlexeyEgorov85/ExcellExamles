using ExcelExamples;
using ExcelExamples.Domain;
using NLog;
using System.Reflection;

Logger logger = LogManager.GetCurrentClassLogger();

args = new string[] {"OpenXML.dll"};

if (!args.Any())
{
    logger.Fatal("Enter excel dll name");
    return;
}

var dllPath = Path.Combine(AppContext.BaseDirectory, args[0]);
if (!File.Exists(dllPath))
{
    logger.Fatal($"File '{dllPath}' does not exist");
    return;
}
Assembly? dll;
try
{
    dll = Assembly.LoadFrom(dllPath);
}
catch (Exception ex)
{
    logger.Fatal($"Failed to load dll '{dllPath}'.Error: {ex.Message}");
    return;
}

var types = dll.GetExportedTypes()
    .Where(p => typeof(IExcel).IsAssignableFrom(p) && !p.IsAbstract)
    .ToArray();

var excelExample = types.FirstOrDefault();
if (excelExample == null)
{
    logger.Fatal($"Failed to find excel class in dll '{dllPath}'");
    return;
}

try
{
    logger.Info($"------------Start {excelExample.Name}------------");
    var p = Activator.CreateInstance(excelExample);
    var p1 = p as IExcel;
    var data = FilmSeeder.Get();
    await p1!.RunAsync(data);
}
catch (Exception ex)
{
    logger.Fatal(ex.Message);
    logger.Debug(ex);
}

logger.Info("----------------End----------------");


