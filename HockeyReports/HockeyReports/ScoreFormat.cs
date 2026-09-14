namespace HockeyReports;

/// <summary>Разбор и печать счёта вида "3*5".</summary>
public static class ScoreFormat
{
    private static readonly char[] Separators = { '*', ':', '-' };

    /// <summary>
    /// Разбирает "3*5" в пару чисел. Если разобрать не удалось, возвращает (null, null)
    /// — значение при этом не теряется, оно печатается как есть.
    /// </summary>
    public static (int? Left, int? Right) Parse(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return (null, null);
        var s = raw.Trim();
        var idx = s.IndexOfAny(Separators);
        if (idx <= 0) return (null, null);

        var left = s[..idx].Trim();
        var right = s[(idx + 1)..].Trim();
        if (int.TryParse(left, out var l) && int.TryParse(right, out var r))
            return (l, r);
        return (null, null);
    }

    /// <summary>
    /// Готовит счёт к печати: "3*5" -> "3 : 5". Числа не меняются, меняется
    /// только разделитель. Неразобранное значение печатается без изменений.
    /// </summary>
    public static string ToDisplay(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        var (l, r) = Parse(raw);
        return l is null || r is null ? raw.Trim() : $"{l} : {r}";
    }
}
