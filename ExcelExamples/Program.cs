using ExcelExamples;
using ExcelExamples.Domain;
using Microsoft.Extensions.Configuration;
using NLog;
using System.Reflection;

Logger logger = LogManager.GetCurrentClassLogger();

try
{
    Config appConfig = GetConfig(logger);
    Type excelExample = GetDllType(appConfig, logger);
    logger.Info($"------------Start {excelExample.Name}------------");
    var p = Activator.CreateInstance(excelExample);
    var p1 = p as IExcel;
    var data = FilmSeeder.Get();
    await p1!.RunAsync(data);
    logger.Info("----------------End----------------");
}
catch (Exception ex)
{
    logger.Fatal(ex.Message);
    logger.Debug(ex);
}


static Config GetConfig(ILogger logger)
{
    Config? appConfig = null;

    try
    {
        appConfig = new ConfigurationBuilder()
        .SetBasePath(AppContext.BaseDirectory)
        .AddJsonFile("config.json")
        .Build()
        .GetSection("Configuration")
        .Get<Config>();
    }
    catch (Exception ex)
    {
        throw new Exception($"Failed to get config. Error: {ex.Message}");
    }

    return appConfig;
}

static Type GetDllType(Config? appConfig, ILogger logger)
{
    if (string.IsNullOrEmpty(appConfig?.DllName))
    {
        throw new ArgumentException("Enter excel dll name");
    }

    var dllPath = Path.Combine(AppContext.BaseDirectory, appConfig.DllName);
    if (!File.Exists(dllPath))
    {
        throw new IOException($"File '{dllPath}' does not exist");
    }
    Assembly? dll;
    try
    {
        dll = Assembly.LoadFrom(dllPath);
    }
    catch (Exception ex)
    {
        throw new IOException($"Failed to load dll '{dllPath}'.Error: {ex.Message}");
    }

    var types = dll.GetExportedTypes()
        .Where(p => typeof(IExcel).IsAssignableFrom(p) && !p.IsAbstract)
        .ToArray();

    var excelExample = types.FirstOrDefault();
    if (excelExample == null)
    {
        throw new IOException($"Failed to find excel class in dll '{dllPath}'");
    }

    return excelExample;
}
