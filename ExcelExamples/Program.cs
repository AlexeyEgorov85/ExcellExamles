using ExcelExamples;

var types = AppDomain.CurrentDomain.GetAssemblies()
    .SelectMany(s => s.GetTypes())
    .Where(p => typeof(IExcel).IsAssignableFrom(p) && !p.IsAbstract)
    .ToArray();

if (!args.Any())
{
    Console.WriteLine($"Enter name excel example to run. Allowed values: '{string.Join(',', types.Select(t => t.Name))}'");
    return;
}

var excelExample = types.FirstOrDefault(t => t.Name == args[0]);
if (excelExample == null)
{
    Console.WriteLine($"Incorrect argument value. Allowed values: '{string.Join(',', types.Select(t => t.Name))}'");
    return;
}

var p = Activator.CreateInstance(excelExample);
var p1 = p as IExcel;
await p1!.RunAsync();

