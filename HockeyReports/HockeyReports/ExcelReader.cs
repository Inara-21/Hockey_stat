using ClosedXML.Excel;
using HockeyReports.Models;

namespace HockeyReports;

/// <summary>Расхождение подписи в левом столбце — только предупреждение.</summary>
public sealed record LabelWarning(string Sheet, int Game, int Row, string Expected, string Actual);

/// <summary>
/// Читает игры из исходной таблицы.
///
/// Каждая игра — жёсткий блок из 15 строк. Данные берутся ПО ПОЗИЦИИ внутри блока,
/// а подписи в левом столбце только сверяются: в исходнике встречаются опечатки
/// ("Навание соревнования"), и разбор не должен от них зависеть.
/// </summary>
public static class ExcelReader
{
    public const int RowsPerGame = 15;
    public const int MaxPlayers = 7;
    private const int FirstValueColumn = 2;   // столбец B — первое значение
    private const int LastValueColumn = 8;    // столбец H — последнее значение

    // Строки ЗНАЧЕНИЙ внутри блока (смещение от первой строки блока).
    private const int OffCompetition = 0;
    private const int OffDate = 1;
    private const int OffTime = 2;
    private const int OffTeam1 = 3;
    private const int OffTeam2 = 4;
    private const int OffRoster1 = 6;
    private const int OffRoster2 = 8;
    private const int OffSecretary = 10;
    private const int OffJudge = 12;
    private const int OffFinalScore = 13;
    private const int OffPeriods = 14;

    // Строки ПОДПИСЕЙ. У составов, секретарей и судей подпись стоит на строку
    // ВЫШЕ своих значений — поэтому смещения здесь отличаются от списка выше.
    private static readonly (int Offset, string Text)[] ExpectedLabels =
    {
        (0, "Название соревнования"), (1, "Дата"), (2, "Время начала матча"),
        (3, "Команда 1"), (4, "Команда2"),
        (5, "Состав команды 1"), (7, "Состав команды 2"),
        (9, "Секретари"), (11, "Судьи"),
        (13, "Итоговый счёт матча"), (14, "Счёт по периодам"),
    };

    public static (List<Game> Games, List<LabelWarning> Warnings) Read(string path)
    {
        var games = new List<Game>();
        var warnings = new List<LabelWarning>();

        using var workbook = new XLWorkbook(path);
        foreach (var sheet in workbook.Worksheets)
        {
            var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
            if (lastRow == 0) continue;

            if (lastRow % RowsPerGame != 0)
                throw new InvalidDataException(
                    $"Лист '{sheet.Name}': {lastRow} строк не делится на {RowsPerGame}. " +
                    "Структура блоков нарушена — разбор остановлен, чтобы не перенести данные неверно.");

            var gameCount = lastRow / RowsPerGame;
            for (var i = 0; i < gameCount; i++)
            {
                var baseRow = i * RowsPerGame + 1;

                foreach (var (offset, expected) in ExpectedLabels)
                {
                    var actual = Text(sheet, baseRow + offset, 1);
                    if (!string.Equals(actual, expected, StringComparison.Ordinal))
                        warnings.Add(new LabelWarning(sheet.Name, i + 1, baseRow + offset, expected, actual));
                }

                var finalRaw = Text(sheet, baseRow + OffFinalScore, FirstValueColumn);
                var (g1, g2) = ScoreFormat.Parse(finalRaw);

                games.Add(new Game
                {
                    Sheet = sheet.Name,
                    Index = i + 1,
                    Row = baseRow,
                    Competition = Text(sheet, baseRow + OffCompetition, FirstValueColumn),
                    Date = Date(sheet, baseRow + OffDate, FirstValueColumn),
                    Time = Time(sheet, baseRow + OffTime, FirstValueColumn),
                    Team1 = Text(sheet, baseRow + OffTeam1, FirstValueColumn),
                    Team2 = Text(sheet, baseRow + OffTeam2, FirstValueColumn),
                    Roster1 = Roster(sheet, baseRow + OffRoster1),
                    Roster2 = Roster(sheet, baseRow + OffRoster2),
                    Secretary = Text(sheet, baseRow + OffSecretary, FirstValueColumn),
                    Judge = Text(sheet, baseRow + OffJudge, FirstValueColumn),
                    FinalScoreRaw = finalRaw,
                    Goals1 = g1,
                    Goals2 = g2,
                    PeriodsRaw = RowValues(sheet, baseRow + OffPeriods),
                });
            }
        }

        return (games, warnings);
    }

    /// <summary>Непустые значения строки слева направо (столбцы B..H).</summary>
    private static List<string> RowValues(IXLWorksheet sheet, int row)
    {
        var values = new List<string>();
        for (var col = FirstValueColumn; col <= LastValueColumn; col++)
        {
            var value = Text(sheet, row, col);
            if (value.Length > 0) values.Add(value);
        }
        return values;
    }

    private static List<Player> Roster(IXLWorksheet sheet, int row) =>
        RowValues(sheet, row).Select(ParsePlayer).ToList();

    /// <summary>
    /// "Мажоров Владислав Владимирович №1" -> номер "1", ФИО "Мажоров Владислав Владимирович".
    /// Терпит слитное написание "Лаврентьев Сергей Игоревич№15".
    /// Если знака № нет, всё значение считается ФИО.
    /// </summary>
    public static Player ParsePlayer(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return new Player("", "");
        var idx = raw.LastIndexOf('№');
        if (idx < 0) return new Player("", raw.Trim());
        return new Player(raw[(idx + 1)..].Trim(), raw[..idx].Trim());
    }

    private static string Text(IXLWorksheet sheet, int row, int col) =>
        sheet.Cell(row, col).GetString().Trim();

    private static DateTime Date(IXLWorksheet sheet, int row, int col)
    {
        var cell = sheet.Cell(row, col);
        if (cell.Value.IsDateTime) return cell.Value.GetDateTime();
        return DateTime.TryParse(cell.GetString(), out var parsed) ? parsed : default;
    }

    private static TimeSpan Time(IXLWorksheet sheet, int row, int col)
    {
        var cell = sheet.Cell(row, col);
        if (cell.Value.IsTimeSpan) return cell.Value.GetTimeSpan();
        if (cell.Value.IsDateTime) return cell.Value.GetDateTime().TimeOfDay;
        return TimeSpan.TryParse(cell.GetString(), out var parsed) ? parsed : default;
    }
}
