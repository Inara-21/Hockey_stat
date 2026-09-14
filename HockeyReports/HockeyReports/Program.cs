using System.Text;
using HockeyReports;
using HockeyReports.Models;
using QuestPDF.Infrastructure;

// Протоколы хоккейных игровых дней: Excel -> PDF (по одному файлу на месяц).
//
// Запуск:
//   dotnet run -- [путь к xlsx] [папка для отчётов]
// По умолчанию:
//   dotnet run -- data/games.xlsx reports

Console.OutputEncoding = Encoding.UTF8;

// QuestPDF: бесплатная лицензия Community.
QuestPDF.Settings.License = LicenseType.Community;

var excelPath = args.Length > 0 ? args[0] : Path.Combine("data", "games.xlsx");
var outputDir = args.Length > 1 ? args[1] : "reports";

if (!File.Exists(excelPath))
{
    Console.Error.WriteLine($"Не найден файл с данными: {Path.GetFullPath(excelPath)}");
    return 2;
}

// На Linux/macOS шрифта Segoe UI нет — берём то, что есть с кириллицей.
if (!OperatingSystem.IsWindows())
    PdfReportBuilder.FontName = "DejaVu Sans";

Directory.CreateDirectory(outputDir);

Console.WriteLine($"Читаю: {Path.GetFullPath(excelPath)}");
var (games, labelWarnings) = ExcelReader.Read(excelPath);
Console.WriteLine($"Прочитано игр: {games.Count}");

// Игровой день = лист Excel (месяц) + дата. Порядок игр внутри дня сохраняем как в таблице.
var days = games
    .GroupBy(g => (g.Sheet, g.Date.Date))
    .OrderBy(g => g.Key.Date)
    .Select(g => (IReadOnlyList<Game>)g.ToList())
    .ToList();

// Месяц = имя листа Excel ("июль", "август"): по одному PDF на месяц.
var months = games.Select(g => g.Sheet).Distinct().ToList();

foreach (var month in months)
{
    var monthDays = days.Where(d => d[0].Sheet == month).ToList();
    var path = Path.Combine(outputDir, $"Протоколы_{month}.pdf");

    PdfReportBuilder.CreateDocument(monthDays).GeneratePdf(path);

    var sheets = monthDays.Sum(d => PdfReportBuilder.SheetsForDay(d.Count));
    var count = monthDays.Sum(d => d.Count);
    Console.WriteLine($"  {month}: игр {count}, дней {monthDays.Count}, листов {sheets} -> {Path.GetFileName(path)}");
}

var result = Validation.Check(games, labelWarnings, days);
var reportPath = Path.Combine(outputDir, "otchet_o_sverke.txt");
File.WriteAllText(reportPath, result.Report, Encoding.UTF8);

Console.WriteLine();
Console.Write(result.Report);
Console.WriteLine($"Отчёт о сверке: {Path.GetFullPath(reportPath)}");

return result.Ok ? 0 : 1;
