using System.Text;
using HockeyReports.Models;

namespace HockeyReports;

/// <summary>
/// Сверка переноса. Ничего не исправляет — только проверяет и сообщает.
///
/// Две разные вещи:
///  * ошибки переноса (их быть не должно) — расхождение числа игр и карточек;
///  * аномалии самих исходных данных — переносятся как есть, но выносятся в список.
/// </summary>
public static class Validation
{
    public sealed record Result(bool Ok, string Report);

    public static Result Check(
        IReadOnlyList<Game> games,
        IReadOnlyList<LabelWarning> labelWarnings,
        IReadOnlyList<IReadOnlyList<Game>> days)
    {
        var text = new StringBuilder();
        var ok = true;

        void Say(string line = "") => text.AppendLine(line);

        Say(new string('=', 72));
        Say("ОТЧЁТ О СВЕРКЕ: Excel -> PDF");
        Say(new string('=', 72));

        Say();
        Say("1. ИСХОДНЫЕ ДАННЫЕ");
        Say($"   Игр прочитано: {games.Count}");
        foreach (var group in games.GroupBy(g => g.Sheet))
            Say($"     лист '{group.Key}': {group.Count()}");

        Say();
        Say("2. РАСКЛАДКА ПО ЛИСТАМ");
        var plannedCards = days.Sum(d => d.Count);
        var plannedSheets = days.Sum(d => PdfReportBuilder.SheetsForDay(d.Count));
        Say($"   Игровых дней: {days.Count}");
        Say($"   Листов А4: {plannedSheets}");
        Say($"   Карточек размещено: {plannedCards}");
        if (plannedCards == games.Count)
        {
            Say($"   OK: {plannedCards} = {games.Count} — ничего не потеряно и не задвоено");
        }
        else
        {
            ok = false;
            Say($"   ОШИБКА: размещено {plannedCards}, а игр {games.Count}");
        }

        Say();
        Say("3. КОНТРОЛЬНЫЕ СУММЫ ГОЛОВ (по итоговому счёту)");
        var goals1 = games.Sum(g => g.Goals1 ?? 0);
        var goals2 = games.Sum(g => g.Goals2 ?? 0);
        var unparsed = games.Count(g => g.Goals1 is null || g.Goals2 is null);
        Say($"   Команда А: {goals1}   Команда В: {goals2}   Всего: {goals1 + goals2}");
        if (unparsed == 0)
        {
            Say("   OK: все итоговые счета разобраны");
        }
        else
        {
            ok = false;
            Say($"   ВНИМАНИЕ: не удалось разобрать счёт в {unparsed} играх (перенесены как есть)");
            foreach (var g in games.Where(g => g.Goals1 is null || g.Goals2 is null))
                Say($"     лист '{g.Sheet}' игра {g.Index} (стр.{g.Row}): '{g.FinalScoreRaw}'");
        }

        Say();
        Say("4. АНОМАЛИИ В ИСХОДНЫХ ДАННЫХ (перенесены как есть, на проверку)");

        Say();
        Say("   4.1 Сумма по периодам != итоговый счёт:");
        var mismatches = 0;
        foreach (var g in games)
        {
            if (g.Goals1 is not int f1 || g.Goals2 is not int f2) continue;

            var parsed = g.PeriodsRaw.Select(ScoreFormat.Parse).ToList();
            if (parsed.Count == 0 || parsed.Any(p => p.Left is null || p.Right is null)) continue;

            var sum1 = parsed.Sum(p => p.Left!.Value);
            var sum2 = parsed.Sum(p => p.Right!.Value);
            if (sum1 == f1 && sum2 == f2) continue;

            mismatches++;
            Say($"     лист '{g.Sheet}' игра {g.Index} (стр.{g.Row}), {g.DateText} {g.TimeText}, " +
                $"{g.Team1} — {g.Team2}: итог {g.FinalScoreRaw}, " +
                $"периоды {string.Join(" ", g.PeriodsRaw)} (сумма {sum1}*{sum2})");
        }
        Say($"     Всего: {mismatches}");

        Say();
        Say("   4.2 Опечатки в подписях левого столбца (на разбор не влияют):");
        foreach (var w in labelWarnings)
            Say($"     лист '{w.Sheet}' игра {w.Game}, строка {w.Row}: " +
                $"'{w.Actual}' вместо '{w.Expected}'");
        Say($"     Всего: {labelWarnings.Count}");

        Say();
        Say("   4.3 Игры вничью (в поле ПОБЕДИТЕЛЬ стоит 'ничья'):");
        var draws = games.Where(g => g.Winner == "ничья").ToList();
        foreach (var g in draws)
            Say($"     лист '{g.Sheet}' игра {g.Index}, {g.DateText} {g.TimeText}, " +
                $"{g.Team1} — {g.Team2}: {g.FinalScoreRaw}");
        Say($"     Всего: {draws.Count}");

        Say();
        Say(new string('=', 72));
        Say(ok ? "ИТОГ: СВЕРКА ПРОЙДЕНА — данные перенесены точно"
               : "ИТОГ: ЕСТЬ ОШИБКИ ПЕРЕНОСА — см. выше");
        Say(new string('=', 72));

        return new Result(ok, text.ToString());
    }
}
