using HockeyReports.Models;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;

namespace HockeyReports;

/// <summary>
/// Строит "Протокол игрового дня": карточки игр по 4 на лист А4,
/// судьи и секретари — внизу последнего листа дня.
/// </summary>
public static class PdfReportBuilder
{
    /// <summary>Карточек игр на одном листе А4 (сетка 2 x 2).</summary>
    public const int CardsPerSheet = 4;

    /// <summary>
    /// Шрифт протокола. Нужен шрифт с кириллицей: на Windows подходит "Segoe UI"
    /// или "Arial". Можно подключить свой файл через <see cref="RegisterFont"/>.
    /// </summary>
    public static string FontName { get; set; } = "Segoe UI";

    private const string BorderColor = "#3A3F6B";
    private const string GridColor = "#B9BDD6";
    private const string TextColor = "#1A1F36";
    private const string LabelColor = "#2B2F4C";
    private const string HeaderFill = "#F2F3F9";
    private const string LogoColor = "#2B2F77";

    private const float BaseFontSize = 7.1f;
    private const float LabelFontSize = 6f;
    private const float CellFontSize = 6.2f;
    private const float CellHeaderFontSize = 5.6f;

    /// <summary>Подключает свой TTF-шрифт (например, если нет Segoe UI).</summary>
    public static void RegisterFont(string fontFilePath, string fontName)
    {
        using var stream = File.OpenRead(fontFilePath);
        FontManager.RegisterFont(stream);
        FontName = fontName;
    }

    /// <summary>Сколько листов займёт игровой день.</summary>
    public static int SheetsForDay(int gamesInDay) =>
        (int)Math.Ceiling(gamesInDay / (double)CardsPerSheet);

    /// <summary>
    /// Готовит документ. Дни идут по порядку, каждый день режется на листы
    /// по <see cref="CardsPerSheet"/> карточек.
    /// </summary>
    public static Document CreateDocument(IReadOnlyList<IReadOnlyList<Game>> days)
    {
        return Document.Create(container =>
        {
            foreach (var day in days)
            {
                if (day.Count == 0) continue;

                var sheets = Chunk(day, CardsPerSheet);
                for (var i = 0; i < sheets.Count; i++)
                {
                    var games = sheets[i];
                    var isLastSheet = i == sheets.Count - 1;
                    var caption = sheets.Count > 1
                        ? $"{day[0].DateText} · лист {i + 1} из {sheets.Count}"
                        : day[0].DateText;

                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        // Поля: сверху больше, чем по краям.
                        page.MarginTop(21, Unit.Millimetre);
                        page.MarginBottom(14, Unit.Millimetre);
                        page.MarginLeft(14, Unit.Millimetre);
                        page.MarginRight(14, Unit.Millimetre);
                        page.DefaultTextStyle(style => style
                            .FontFamily(FontName)
                            .FontSize(BaseFontSize)
                            .FontColor(TextColor));

                        page.Header().Element(header => ComposeHeader(header, caption));
                        page.Content().Element(content =>
                            ComposeContent(content, games, isLastSheet ? day[0] : null));
                    });
                }
            }
        });
    }

    private static List<IReadOnlyList<Game>> Chunk(IReadOnlyList<Game> games, int size)
    {
        var chunks = new List<IReadOnlyList<Game>>();
        for (var i = 0; i < games.Count; i += size)
            chunks.Add(games.Skip(i).Take(size).ToList());
        return chunks;
    }

    private static void ComposeHeader(IContainer container, string caption)
    {
        container.PaddingBottom(8, Unit.Millimetre).Row(row =>
        {
            row.RelativeItem(1f).AlignMiddle()
               .Text("ХОККЕЙ").Bold().FontSize(12f).FontColor(LogoColor);

            row.RelativeItem(2.2f).AlignMiddle().AlignCenter()
               .Text("ПРОТОКОЛ ИГРОВОГО ДНЯ").Bold().FontSize(12.75f);

            row.RelativeItem(1.8f).AlignMiddle().AlignRight()
               .Text(caption).Bold().FontSize(10f);
        });
    }

    /// <param name="officialsFrom">
    /// Игра, из которой берутся судья и секретарь для блока внизу листа.
    /// null — блок не печатается (лист не последний в дне).
    /// </param>
    private static void ComposeContent(IContainer container, IReadOnlyList<Game> games, Game? officialsFrom)
    {
        container.Column(column =>
        {
            column.Spacing(7, Unit.Millimetre);

            for (var i = 0; i < games.Count; i += 2)
            {
                var left = games[i];
                var right = i + 1 < games.Count ? games[i + 1] : null;

                column.Item().Row(row =>
                {
                    row.Spacing(7, Unit.Millimetre);
                    row.RelativeItem().Element(cell => ComposeCard(cell, left));
                    if (right is null)
                        row.RelativeItem();          // пустая половина строки
                    else
                        row.RelativeItem().Element(cell => ComposeCard(cell, right));
                });
            }

            if (officialsFrom is not null)
                column.Item().Element(cell => ComposeOfficials(cell, officialsFrom));
        });
    }

    private static void ComposeCard(IContainer container, Game game)
    {
        container
            .Border(1.3f).BorderColor(BorderColor)
            .PaddingVertical(4.5f, Unit.Millimetre)
            .PaddingHorizontal(5, Unit.Millimetre)
            .Column(column =>
            {
                column.Item().PaddingBottom(3, Unit.Millimetre)
                      .Element(cell => Field(cell, "СОРЕВНОВАНИЕ", game.Competition));

                column.Item().PaddingBottom(2.2f, Unit.Millimetre).Row(row =>
                {
                    row.Spacing(5, Unit.Millimetre);
                    row.RelativeItem().Element(cell => Field(cell, "ДАТА", game.DateText));
                    row.RelativeItem().Element(cell => Field(cell, "ВРЕМЯ", game.TimeText));
                });

                column.Item().Row(row =>
                {
                    row.Spacing(5, Unit.Millimetre);
                    row.RelativeItem().Element(cell => Field(cell, "КОМАНДА А", game.Team1));
                    row.RelativeItem().Element(cell => Field(cell, "КОМАНДА В", game.Team2));
                });

                column.Item().PaddingVertical(2.6f, Unit.Millimetre).Row(row =>
                {
                    row.Spacing(5, Unit.Millimetre);
                    row.RelativeItem().Element(cell => RosterTable(cell, game.Roster1));
                    row.RelativeItem().Element(cell => RosterTable(cell, game.Roster2));
                });

                column.Item().Row(row =>
                {
                    row.Spacing(5, Unit.Millimetre);
                    row.RelativeItem().Element(cell => Field(cell, "ИТОГОВЫЙ СЧЁТ", game.FinalScoreText));
                    row.RelativeItem().Element(cell => Field(cell, "ПОБЕДИТЕЛЬ", game.Winner));
                });

                column.Item().PaddingTop(2.2f, Unit.Millimetre)
                      .Element(cell => Field(cell, "СЧЁТ ПО ПЕРИОДАМ", game.RegulationText));

                column.Item().PaddingTop(2.2f, Unit.Millimetre)
                      .Element(cell => Field(cell, "ОВЕРТАЙМЫ", game.OvertimeText));
            });
    }

    /// <summary>Подпись слева, значение курсивом справа, снизу линия — как в бланке.</summary>
    private static void Field(IContainer container, string label, string value)
    {
        container
            .BorderBottom(1).BorderColor(GridColor)
            .PaddingVertical(2)
            .MinHeight(11)
            .Row(row =>
            {
                row.Spacing(4);
                row.AutoItem().AlignBottom()
                   .Text(label).Bold().FontSize(LabelFontSize).FontColor(LabelColor);
                row.RelativeItem().AlignBottom()
                   .Text(value).Italic();
            });
    }

    private static void RosterTable(IContainer container, IReadOnlyList<Player> roster)
    {
        container.Table(table =>
        {
            table.ColumnsDefinition(columns =>
            {
                columns.ConstantColumn(16);
                columns.RelativeColumn();
            });

            table.Header(header =>
            {
                header.Cell().Element(HeaderCell).AlignCenter().Text("№");
                header.Cell().Element(HeaderCell).AlignCenter().Text("ИГРОК");
            });

            // Всегда печатаем фиксированное число строк, чтобы карточки были одной высоты.
            for (var i = 0; i < ExcelReader.MaxPlayers; i++)
            {
                var player = i < roster.Count ? roster[i] : new Player("", "");
                table.Cell().Element(BodyCell).AlignCenter().Text(player.Number);
                table.Cell().Element(BodyCell).Text(player.Name).Italic();
            }
        });
    }

    private static IContainer HeaderCell(IContainer container) =>
        container
            .Border(0.7f).BorderColor(GridColor)
            .Background(HeaderFill)
            .PaddingVertical(1.4f).PaddingHorizontal(3)
            .DefaultTextStyle(style => style.Bold().FontSize(CellHeaderFontSize));

    private static IContainer BodyCell(IContainer container) =>
        container
            .Border(0.7f).BorderColor(GridColor)
            .PaddingVertical(1.4f).PaddingHorizontal(3)
            .DefaultTextStyle(style => style.FontSize(CellFontSize));

    private static void ComposeOfficials(IContainer container, Game game)
    {
        container
            .PaddingTop(6, Unit.Millimetre)
            .Border(1.3f).BorderColor(BorderColor)
            .PaddingVertical(5, Unit.Millimetre)
            .PaddingHorizontal(6, Unit.Millimetre)
            .Row(row =>
            {
                row.Spacing(10, Unit.Millimetre);
                row.RelativeItem().Element(cell => OfficialsColumn(cell, "СУДЬИ", game.Judge));
                row.RelativeItem().Element(cell => OfficialsColumn(cell, "СЕКРЕТАРИ", game.Secretary));
            });
    }

    private static void OfficialsColumn(IContainer container, string title, string name)
    {
        container.Column(column =>
        {
            column.Item().PaddingBottom(3.5f, Unit.Millimetre)
                  .Text(title).Bold().FontSize(6.75f);
            column.Item().BorderBottom(1).BorderColor(GridColor).PaddingBottom(3)
                  .Text(name).Italic();
        });
    }
}
