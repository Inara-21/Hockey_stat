namespace HockeyReports.Models;

/// <summary>Один игрок состава: игровой номер и ФИО.</summary>
public sealed record Player(string Number, string Name);

/// <summary>
/// Одна игра — ровно то, что записано в исходной таблице.
/// Значения не пересчитываются и не исправляются.
/// </summary>
public sealed class Game
{
    /// <summary>Имя листа Excel, откуда взята игра (например, "июль").</summary>
    public string Sheet { get; init; } = "";

    /// <summary>Порядковый номер игры на листе, начиная с 1.</summary>
    public int Index { get; init; }

    /// <summary>Номер первой строки блока игры в Excel — для сверки с исходником.</summary>
    public int Row { get; init; }

    public string Competition { get; init; } = "";
    public DateTime Date { get; init; }
    public TimeSpan Time { get; init; }
    public string Team1 { get; init; } = "";
    public string Team2 { get; init; } = "";
    public IReadOnlyList<Player> Roster1 { get; init; } = Array.Empty<Player>();
    public IReadOnlyList<Player> Roster2 { get; init; } = Array.Empty<Player>();
    public string Secretary { get; init; } = "";
    public string Judge { get; init; } = "";

    /// <summary>Итоговый счёт ровно как в Excel, например "3*5".</summary>
    public string FinalScoreRaw { get; init; } = "";

    /// <summary>Голы команды 1 из итогового счёта; null, если разобрать не удалось.</summary>
    public int? Goals1 { get; init; }

    /// <summary>Голы команды 2 из итогового счёта; null, если разобрать не удалось.</summary>
    public int? Goals2 { get; init; }

    /// <summary>Счёт по периодам ровно как в Excel, например ["2*1", "1*2", "0*2"].</summary>
    public IReadOnlyList<string> PeriodsRaw { get; init; } = Array.Empty<string>();

    /// <summary>Дата в формате протокола.</summary>
    public string DateText => Date.ToString("dd.MM.yyyy");

    /// <summary>Время начала в формате протокола.</summary>
    public string TimeText => Time.ToString(@"hh\:mm");

    /// <summary>Итоговый счёт для печати: "3 : 5" (числа те же, другой разделитель).</summary>
    public string FinalScoreText => ScoreFormat.ToDisplay(FinalScoreRaw);

    /// <summary>Счёт за первые три периода, через запятую.</summary>
    public string RegulationText =>
        string.Join(", ", PeriodsRaw.Take(3).Select(ScoreFormat.ToDisplay));

    /// <summary>Счёт за периоды после третьего (овертаймы), через запятую.</summary>
    public string OvertimeText =>
        string.Join(", ", PeriodsRaw.Skip(3).Select(ScoreFormat.ToDisplay));

    /// <summary>
    /// Победитель. В исходной таблице такого поля нет — определяется по итоговому счёту.
    /// При равном счёте возвращается "ничья".
    /// </summary>
    public string Winner
    {
        get
        {
            if (Goals1 is not int g1 || Goals2 is not int g2) return "";
            if (g1 > g2) return Team1;
            if (g2 > g1) return Team2;
            return "ничья";
        }
    }
}
