using System.Globalization;

namespace BrimOutlook.Outlook;

/// <summary>Kurze deutsche Zeitangaben für die Notch.</summary>
public static class Format
{
    // Fest statt über CultureInfo: Windows (NLS) und ICU kürzen Wochentage verschieden ab („Sa“ / „Sa.“).
    private static readonly string[] Weekdays = ["So", "Mo", "Di", "Mi", "Do", "Fr", "Sa"];

    /// <summary>„in 12 min“, „in 1 h 5 min“, „jetzt“, „seit 3 min“.</summary>
    public static string Countdown(DateTime start, DateTime now)
    {
        var minutes = (int)Math.Ceiling((start - now).TotalMinutes);
        if (minutes > 0)
        {
            return minutes < 60 ? $"in {minutes} min" : $"in {minutes / 60} h{(minutes % 60 == 0 ? "" : $" {minutes % 60} min")}";
        }
        var since = (int)Math.Floor((now - start).TotalMinutes);
        return since < 1 ? "jetzt" : $"seit {since} min";
    }

    /// <summary>„09:30–10:00“ oder „ganztägig“.</summary>
    public static string TimeRange(MeetingSummary meeting) =>
        meeting.AllDay ? "ganztägig" : $"{meeting.Start:HH:mm}–{meeting.End:HH:mm}";

    /// <summary>Empfangszeit: heute „14:05“, gestern „gestern“, sonst „Mo 20.10.“.</summary>
    public static string Received(DateTime received, DateTime now) =>
        received.Date == now.Date ? received.ToString("HH:mm", CultureInfo.InvariantCulture)
        : received.Date == now.Date.AddDays(-1) ? "gestern"
        : Day(received);

    /// <summary>Fälligkeit einer Aufgabe: „überfällig seit 2 Tagen“, „heute fällig“, „morgen fällig“, „fällig Fr 24.10.“.</summary>
    public static string Due(DateTime? due, DateTime now)
    {
        if (due is not { } date)
        {
            return "ohne Termin";
        }
        var days = (date.Date - now.Date).Days;
        return days switch
        {
            < -1 => $"überfällig seit {-days} Tagen",
            -1 => "seit gestern überfällig",
            0 => "heute fällig",
            1 => "morgen fällig",
            _ => $"fällig {Day(date)}",
        };
    }

    /// <summary>„Fr 24.10.“</summary>
    public static string Day(DateTime date) => $"{Weekdays[(int)date.DayOfWeek]} {date.ToString("dd.MM.", CultureInfo.InvariantCulture)}";

    /// <summary>„1 ungelesene E-Mail“ bzw. „3 ungelesene E-Mails“.</summary>
    public static string Unread(int count) => count == 1 ? "1 ungelesene E-Mail" : $"{count} ungelesene E-Mails";

    /// <summary>Kürzt auf <paramref name="max"/> Zeichen mit „…“; Brim nimmt höchstens 2000 Zeichen pro Text.</summary>
    public static string Clip(string? text, int max = 200)
    {
        var value = (text ?? "").ReplaceLineEndings(" ").Trim();
        return value.Length <= max ? value : value[..(max - 1)] + "…";
    }
}
