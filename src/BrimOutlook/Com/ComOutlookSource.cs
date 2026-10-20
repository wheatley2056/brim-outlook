using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using BrimOutlook.Outlook;
using Microsoft.CSharp.RuntimeBinder;

namespace BrimOutlook.Com;

/// <summary>
/// Outlook über das Objektmodell (COM, spätes Binden): braucht das klassische Outlook für Windows, das
/// läuft. Das Plugin hängt sich nur an eine laufende Instanz und startet Outlook nie von selbst.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class ComOutlookSource : IOutlookSource
{
    private const string ProgId = "Outlook.Application";

    // OlDefaultFolders, OlItemType und weitere Konstanten aus dem Outlook-Objektmodell
    private const int FolderInbox = 6;
    private const int FolderCalendar = 9;
    private const int FolderTasks = 13;
    private const int ItemMail = 0;
    private const int ItemAppointment = 1;
    private const int ItemTask = 3;
    private const int ImportanceHigh = 2;
    private const int MeetingCanceled = 5;
    private const int MeetingReceivedAndCanceled = 7;
    private const int ResponseDeclined = 4;
    private const int WindowMinimized = 1;
    private const int WindowNormal = 2;

    /// <summary>So viele Termine bzw. Aufgaben liest das Plugin höchstens.</summary>
    private const int MaxMeetings = 40;
    private const int MaxTasks = 100;

    /// <summary>Teams trägt den Link zum Beitreten in diese benannte Eigenschaft ein.</summary>
    private const string TeamsUrlProperty = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/SkypeTeamsMeetingUrl";

    /// <summary>Outlook nimmt 4501-01-01 für „kein Datum“.</summary>
    private static readonly DateTime NoDate = new(4501, 1, 1);

    private readonly StaDispatcher _sta = new("Outlook COM");

    public Task<OutlookSnapshot> ReadAsync(DateTime now, DateTime until, int maxMails, CancellationToken cancellationToken = default) =>
        _sta.InvokeAsync(() => Use((_, session) =>
        {
            object mapi = session;
            var (unread, mails) = ReadMail(mapi, maxMails);
            return new OutlookSnapshot(unread, mails, ReadCalendar(mapi, now, until), ReadTasks(mapi));
        })).WaitAsync(cancellationToken);

    public Task OpenAsync(ItemRef item) => _sta.InvokeAsync(() => Use((_, session) =>
    {
        object found = session.GetItemFromID(item.EntryId);
        try
        {
            dynamic target = found;
            if (item.OccurrenceStart is { } start && Get(() => target.IsRecurring, false))
            {
                // Bei Serien ist die EntryID die der Serie; geöffnet werden soll aber dieser eine Termin.
                object pattern = target.GetRecurrencePattern();
                try
                {
                    object occurrence = ((dynamic)pattern).GetOccurrence(start);
                    ((dynamic)occurrence).Display(false);
                    Interop.Release(occurrence);
                }
                finally
                {
                    Interop.Release(pattern);
                }
            }
            else
            {
                target.Display(false);
            }
        }
        finally
        {
            Interop.Release(found);
        }
        return true;
    }));

    public Task MarkReadAsync(ItemRef item) => Change(item, target =>
    {
        target.UnRead = false;
        target.Save();
    });

    public Task CompleteTaskAsync(ItemRef item) => Change(item, target =>
    {
        target.MarkComplete();
        target.Save();
    });

    public Task CreateAsync(OutlookKind kind) => _sta.InvokeAsync(() => Use((app, _) =>
    {
        object created = app.CreateItem(kind switch
        {
            OutlookKind.Mail => ItemMail,
            OutlookKind.Meeting => ItemAppointment,
            _ => ItemTask,
        });
        try
        {
            ((dynamic)created).Display(false);
        }
        finally
        {
            Interop.Release(created);
        }
        return true;
    }));

    public Task ShowFolderAsync(OutlookKind kind) => _sta.InvokeAsync(() => Use((app, session) =>
    {
        object folder = session.GetDefaultFolder(DefaultFolder(kind));
        object? explorer = app.ActiveExplorer();
        try
        {
            if (explorer is null)
            {
                ((dynamic)folder).Display();
            }
            else
            {
                dynamic window = explorer;
                window.CurrentFolder = folder;
                if (Get(() => window.WindowState, WindowNormal) == WindowMinimized)
                {
                    window.WindowState = WindowNormal;
                }
                window.Activate();
            }
        }
        finally
        {
            Interop.Release(explorer);
            Interop.Release(folder);
        }
        return true;
    }));

    public Task LaunchAsync()
    {
        // „outlook.exe“ findet die Shell über App Paths, auch wenn Office nicht im PATH liegt.
        Process.Start(new ProcessStartInfo("outlook.exe") { UseShellExecute = true })?.Dispose();
        return Task.CompletedTask;
    }

    public void Dispose() => _sta.Dispose();

    // ───────── Lesen ─────────

    private static (int Unread, IReadOnlyList<MailSummary> Mails) ReadMail(object session, int max)
    {
        object folder = ((dynamic)session).GetDefaultFolder(FolderInbox);
        object items = ((dynamic)folder).Items;
        object unread = ((dynamic)items).Restrict("[UnRead] = True");
        try
        {
            dynamic list = unread;
            int count = list.Count;
            list.Sort("[ReceivedTime]", true);
            var mails = new List<MailSummary>();
            for (var i = 1; i <= count && mails.Count < max; i++)
            {
                object entry = list.Item(i);
                try
                {
                    dynamic item = entry;
                    // SenderName gehört zu den geschützten Eigenschaften; ohne aktuellen Virenschutz fragt Outlook nach.
                    mails.Add(new MailSummary(
                        new ItemRef((string)item.EntryID),
                        Get(() => item.SenderName, ""),
                        Get(() => item.Subject, ""),
                        Get(() => item.ReceivedTime, Get(() => item.CreationTime, DateTime.MinValue)),
                        Get(() => item.Importance, 1) == ImportanceHigh));
                }
                catch (Exception ex) when (IsItemError(ex))
                {
                    // Element ohne die erwarteten Eigenschaften (z. B. ein Bericht): überspringen.
                }
                finally
                {
                    Interop.Release(entry);
                }
            }
            return (count, mails);
        }
        finally
        {
            Interop.Release(unread);
            Interop.Release(items);
            Interop.Release(folder);
        }
    }

    private static IReadOnlyList<MeetingSummary> ReadCalendar(object session, DateTime now, DateTime until)
    {
        object folder = ((dynamic)session).GetDefaultFolder(FolderCalendar);
        object items = ((dynamic)folder).Items;
        object? found = null;
        try
        {
            // Serientermine einzeln: erst nach Beginn sortieren, dann IncludeRecurrences, dann einschränken.
            dynamic all = items;
            all.Sort("[Start]");
            all.IncludeRecurrences = true;
            found = all.Restrict($"[Start] < '{Jet(until)}' AND [End] > '{Jet(now)}'");
            dynamic list = found!;
            var meetings = new List<MeetingSummary>();
            object? entry = list.GetFirst();
            while (entry is not null && meetings.Count < MaxMeetings)
            {
                try
                {
                    if (ReadMeeting(entry) is { } meeting)
                    {
                        meetings.Add(meeting);
                    }
                }
                catch (Exception ex) when (IsItemError(ex))
                {
                }
                finally
                {
                    Interop.Release(entry);
                }
                entry = list.GetNext();
            }
            return meetings;
        }
        finally
        {
            Interop.Release(found);
            Interop.Release(items);
            Interop.Release(folder);
        }
    }

    private static MeetingSummary? ReadMeeting(object entry)
    {
        dynamic item = entry;
        var status = Get(() => item.MeetingStatus, 0);
        if (status is MeetingCanceled or MeetingReceivedAndCanceled || Get(() => item.ResponseStatus, 0) == ResponseDeclined)
        {
            return null;
        }
        DateTime start = item.Start;
        string entryId = item.EntryID;
        var recurring = Get(() => item.IsRecurring, false);
        string? location = Get<string?>(() => item.Location, null);
        var link = MeetingLinks.Find(TeamsUrl(entry), location);
        if (link is null)
        {
            // Body ist geschützt und kann groß sein; nur lesen, wenn sonst kein Link da ist.
            string? body = Get<string?>(() => item.Body, null);
            link = MeetingLinks.Find(body);
        }
        return new MeetingSummary(
            new ItemRef(entryId, recurring ? start : null),
            Get(() => item.Subject, ""),
            start,
            Get(() => item.End, start),
            location,
            Get(() => item.AllDayEvent, false),
            link);
    }

    private static string? TeamsUrl(object item)
    {
        object? accessor = null;
        try
        {
            accessor = ((dynamic)item).PropertyAccessor;
            return ((dynamic)accessor!).GetProperty(TeamsUrlProperty) as string;
        }
        catch (Exception ex) when (IsItemError(ex))
        {
            return null; // Eigenschaft fehlt: kein Teams-Termin
        }
        finally
        {
            Interop.Release(accessor);
        }
    }

    private static IReadOnlyList<TaskSummary> ReadTasks(object session)
    {
        object folder = ((dynamic)session).GetDefaultFolder(FolderTasks);
        object items = ((dynamic)folder).Items;
        object open = ((dynamic)items).Restrict("[Complete] = False");
        try
        {
            dynamic list = open;
            int count = list.Count;
            var tasks = new List<TaskSummary>();
            for (var i = 1; i <= count && tasks.Count < MaxTasks; i++)
            {
                object entry = list.Item(i);
                try
                {
                    dynamic item = entry;
                    var due = Get(() => item.DueDate, NoDate);
                    tasks.Add(new TaskSummary(
                        new ItemRef((string)item.EntryID),
                        Get(() => item.Subject, ""),
                        due >= NoDate ? null : due,
                        Get(() => item.Importance, 1) == ImportanceHigh));
                }
                catch (Exception ex) when (IsItemError(ex))
                {
                }
                finally
                {
                    Interop.Release(entry);
                }
            }
            // Überfällige zuerst, Aufgaben ohne Termin zuletzt; bei gleichem Datum die wichtigen.
            return [.. tasks.OrderBy(t => t.Due ?? DateTime.MaxValue).ThenByDescending(t => t.Important)];
        }
        finally
        {
            Interop.Release(open);
            Interop.Release(items);
            Interop.Release(folder);
        }
    }

    // ───────── Hilfen ─────────

    private Task Change(ItemRef item, Action<dynamic> change) => _sta.InvokeAsync(() => Use((_, session) =>
    {
        object found = session.GetItemFromID(item.EntryId);
        try
        {
            change(found);
        }
        finally
        {
            Interop.Release(found);
        }
        return true;
    }));

    /// <summary>Hängt sich an das laufende Outlook und gibt Anwendung und MAPI-Sitzung danach wieder frei.</summary>
    private static T Use<T>(Func<dynamic, dynamic, T> work)
    {
        var app = Interop.GetActive(ProgId, out var installed) ?? throw new OutlookUnavailableException(installed
            ? "Outlook läuft nicht. Starte Outlook, dann erscheinen hier E-Mails, Termine und Aufgaben."
            : "Das klassische Outlook für Windows ist nicht installiert. Das neue Outlook bietet keine Schnittstelle für Plugins.");
        object? session = null;
        try
        {
            session = ((dynamic)app).GetNamespace("MAPI");
            return work(app, session!);
        }
        catch (COMException ex) when (ex.HResult is Interop.RpcDisconnected or Interop.RpcServerUnavailable or Interop.RpcCallFailed)
        {
            throw new OutlookUnavailableException("Outlook wurde beendet.", ex);
        }
        catch (COMException ex) when (ex.HResult is Interop.RpcCallRejected or Interop.RpcRetryLater)
        {
            throw new OutlookUnavailableException("Outlook ist gerade beschäftigt (offener Dialog?).", ex);
        }
        finally
        {
            Interop.Release(session);
            Interop.Release(app);
        }
    }

    private static int DefaultFolder(OutlookKind kind) => kind switch
    {
        OutlookKind.Mail => FolderInbox,
        OutlookKind.Meeting => FolderCalendar,
        _ => FolderTasks,
    };

    /// <summary>Datum für Restrict: Outlook erwartet das Format der Ländereinstellungen, ohne Sekunden.</summary>
    private static string Jet(DateTime value) => value.ToString("g", CultureInfo.CurrentCulture);

    /// <summary>Liest eine Eigenschaft; fehlt sie oder verweigert Outlook den Zugriff, gilt <paramref name="fallback"/>.</summary>
    private static T Get<T>(Func<object?> read, T fallback)
    {
        try
        {
            return read() is T value ? value : fallback;
        }
        catch (Exception ex) when (IsItemError(ex))
        {
            return fallback;
        }
    }

    private static bool IsItemError(Exception ex) =>
        ex is RuntimeBinderException or InvalidCastException
        || (ex is COMException com && com.HResult is not (Interop.RpcDisconnected or Interop.RpcServerUnavailable or Interop.RpcCallFailed));
}
