namespace BrimOutlook.Outlook;

/// <summary>Verweis auf ein Outlook-Element. Bei Serienterminen gehört der Beginn dazu, sonst wäre nur die Serie gemeint.</summary>
public sealed record ItemRef(string EntryId, DateTime? OccurrenceStart = null);

/// <summary>Eine ungelesene E-Mail im Posteingang.</summary>
public sealed record MailSummary(ItemRef Ref, string Sender, string Subject, DateTime Received, bool Important);

/// <summary>Ein Termin aus dem Standardkalender (Zeiten lokal).</summary>
public sealed record MeetingSummary(ItemRef Ref, string Subject, DateTime Start, DateTime End, string? Location, bool AllDay, Uri? JoinLink)
{
    public bool IsRunning(DateTime now) => Start <= now && now < End;
}

/// <summary>Eine offene Aufgabe (lokale Zeit; <see cref="Due"/> ist nur ein Datum).</summary>
public sealed record TaskSummary(ItemRef Ref, string Subject, DateTime? Due, bool Important)
{
    public bool IsOverdue(DateTime now) => Due is { } due && due.Date < now.Date;

    public bool IsDueToday(DateTime now) => Due is { } due && due.Date == now.Date;
}

/// <summary>Stand von Posteingang, Kalender und Aufgaben zu einem Zeitpunkt.</summary>
public sealed record OutlookSnapshot(
    int UnreadCount,
    IReadOnlyList<MailSummary> Mails,
    IReadOnlyList<MeetingSummary> Meetings,
    IReadOnlyList<TaskSummary> Tasks);

/// <summary>Was „Neu“ bzw. „Anzeigen“ öffnet.</summary>
public enum OutlookKind
{
    Mail,
    Meeting,
    Task,
}

/// <summary>Outlook läuft nicht (oder ist gerade nicht erreichbar).</summary>
public sealed class OutlookUnavailableException(string message, Exception? inner = null) : Exception(message, inner);

/// <summary>Zugriff auf Outlook. Die Umsetzung über COM steht in <c>Com/ComOutlookSource</c>.</summary>
public interface IOutlookSource : IDisposable
{
    /// <param name="maxMails">So viele ungelesene E-Mails mitliefern (gezählt werden alle).</param>
    /// <param name="until">Termine, die vor diesem Zeitpunkt beginnen und nach <paramref name="now"/> enden.</param>
    Task<OutlookSnapshot> ReadAsync(DateTime now, DateTime until, int maxMails, CancellationToken cancellationToken = default);

    Task OpenAsync(ItemRef item);

    Task MarkReadAsync(ItemRef item);

    Task CompleteTaskAsync(ItemRef item);

    Task CreateAsync(OutlookKind kind);

    Task ShowFolderAsync(OutlookKind kind);

    /// <summary>Startet Outlook, wenn es nicht läuft.</summary>
    Task LaunchAsync();
}
