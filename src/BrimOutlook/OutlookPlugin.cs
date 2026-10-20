using System.Text.Json;
using Brim.Core.Plugins;
using Brim.Core.Widgets;
using BrimOutlook.Outlook;

namespace BrimOutlook;

/// <summary>
/// Liest Outlook in festen Abständen, zeigt Karten im Tab „Plugins“, Termine mit Countdown in der Notch und
/// kurze Meldungen bei neuen E-Mails. Schickt nur, was sich geändert hat.
/// </summary>
public sealed class OutlookPlugin
{
    /// <summary>So oft wird neu gezeichnet (Countdown); gelesen wird nach <see cref="PluginSettings.Refresh"/>.</summary>
    public static readonly TimeSpan Tick = TimeSpan.FromSeconds(15);

    public const string MailView = "mail";
    public const string CalendarView = "calendar";
    public const string TasksView = "tasks";
    public const string UnavailableView = "outlook";

    private static readonly string[] ActivityIds = ["meeting", "mail", "tasks"];

    private readonly INotchSink _sink;
    private readonly IOutlookSource _source;
    private readonly TimeProvider _time;
    private readonly Func<Uri, Task> _openUrl;
    private readonly SemaphoreSlim _gate = new(1, 1);
    private readonly ActionKeys _keys = new();
    private readonly Dictionary<string, string> _views = [];
    private readonly Dictionary<string, string> _activities = [];
    private readonly HashSet<ItemRef> _announced = [];

    private OutlookSnapshot? _snapshot;
    private string? _unavailable;
    private DateTime? _lastAttempt;
    private DateTime _lastRead;
    private HashSet<string>? _seenMail;
    private bool _notificationsSeen;

    public OutlookPlugin(INotchSink sink, IOutlookSource source, TimeProvider time, Func<Uri, Task> openUrl, PluginSettings settings)
    {
        _sink = sink;
        _source = source;
        _time = time;
        _openUrl = openUrl;
        Settings = settings;
    }

    public PluginSettings Settings { get; private set; }

    private DateTime Now => _time.GetLocalNow().DateTime;

    /// <summary>Liest und zeichnet, bis <paramref name="cancellationToken"/> abbricht.</summary>
    public async Task RunAsync(CancellationToken cancellationToken)
    {
        while (!cancellationToken.IsCancellationRequested)
        {
            if (_lastAttempt is not { } last || Now - last >= Settings.Refresh)
            {
                await RefreshAsync();
            }
            else
            {
                await RenderAsync();
            }
            try
            {
                await Task.Delay(Tick, _time, cancellationToken);
            }
            catch (OperationCanceledException)
            {
                break;
            }
        }
    }

    /// <summary>Liest Outlook sofort neu.</summary>
    public async Task RefreshAsync()
    {
        await _gate.WaitAsync();
        try
        {
            var now = Now;
            _lastAttempt = now;
            try
            {
                var snapshot = await _source.ReadAsync(now, now.Date.AddDays(2), Settings.MaxItems);
                await AnnounceNewMailAsync(snapshot, now);
                _snapshot = snapshot;
                _unavailable = null;
            }
            catch (OutlookUnavailableException ex)
            {
                _snapshot = null;
                _unavailable = ex.Message;
            }
            catch (Exception ex)
            {
                await LogAsync($"Outlook lesen fehlgeschlagen: {ex.Message}", "warn");
                return;
            }
            await RenderLockedAsync(now);
        }
        finally
        {
            _gate.Release();
        }
    }

    /// <summary>Im Tab „Plugins“ nicht mit Daten von vor einer halben Minute stehen.</summary>
    public Task RefreshIfStaleAsync() =>
        _lastAttempt is { } last && Now - last < TimeSpan.FromSeconds(10) ? Task.CompletedTask : RefreshAsync();

    public Task ApplySettingsAsync(PluginSettings settings)
    {
        Settings = settings;
        return RefreshAsync();
    }

    /// <summary>Eine Windows-Benachrichtigung von Outlook: als HUD zeigen und gleich neu lesen.</summary>
    public async Task OnNotificationAsync(NotificationParams notification)
    {
        // Kommen Outlooks eigene Meldungen an, braucht es keine zweite für dieselbe E-Mail.
        _notificationsSeen = true;
        var text = string.IsNullOrWhiteSpace(notification.Body) ? notification.Title : $"{notification.Title}: {notification.Body}";
        await ShowHudAsync(Views.Icons.Mail, Format.Clip(text, 120));
        await RefreshAsync();
    }

    /// <summary>Knopf in einer Karte oder Aktivität, z. B. <c>open:i3</c>.</summary>
    public async Task HandleActionAsync(string actionId)
    {
        var colon = actionId.IndexOf(':');
        var verb = colon < 0 ? actionId : actionId[..colon];
        var argument = colon < 0 ? "" : actionId[(colon + 1)..];
        try
        {
            switch (verb)
            {
                case "refresh":
                    await RefreshAsync();
                    break;
                case "launch":
                    await _source.LaunchAsync();
                    await ShowHudAsync(Views.Icons.Mail, "Outlook startet …");
                    // Outlook meldet sich erst nach ein paar Sekunden beim System an.
                    await Task.Delay(TimeSpan.FromSeconds(8), _time);
                    await RefreshAsync();
                    break;
                case "new" when Enum.TryParse<OutlookKind>(argument, ignoreCase: true, out var kind):
                    await _source.CreateAsync(kind);
                    break;
                case "show" when Enum.TryParse<OutlookKind>(argument, ignoreCase: true, out var kind):
                    await _source.ShowFolderAsync(kind);
                    break;
                case "open" when _keys.Resolve(argument) is { } item:
                    await _source.OpenAsync(item);
                    break;
                case "read" when _keys.Resolve(argument) is { } item:
                    await _source.MarkReadAsync(item);
                    await RefreshAsync();
                    break;
                case "done" when _keys.Resolve(argument) is { } item:
                    await _source.CompleteTaskAsync(item);
                    var subject = _snapshot?.Tasks.FirstOrDefault(t => t.Ref == item)?.Subject;
                    await ShowHudAsync(Glyphs.CheckMark, subject is { Length: > 0 } ? $"Erledigt: {Format.Clip(subject, 80)}" : "Aufgabe erledigt");
                    await RefreshAsync();
                    break;
                case "join" when _keys.Resolve(argument) is { } item
                    && _snapshot?.Meetings.FirstOrDefault(m => m.Ref == item)?.JoinLink is { } link:
                    await _openUrl(link);
                    break;
                default:
                    await LogAsync($"Unbekannte oder veraltete Aktion „{actionId}“", "warn");
                    break;
            }
        }
        catch (OutlookUnavailableException ex)
        {
            await ShowHudAsync(Views.Icons.Warning, ex.Message);
            await RefreshAsync();
        }
        catch (Exception ex)
        {
            await LogAsync($"Aktion {verb} fehlgeschlagen: {ex.Message}", "error");
        }
    }

    /// <summary>Zeichnet mit dem letzten Stand neu (Countdown, Termin beginnt).</summary>
    public async Task RenderAsync()
    {
        await _gate.WaitAsync();
        try
        {
            await RenderLockedAsync(Now);
        }
        finally
        {
            _gate.Release();
        }
    }

    private async Task RenderLockedAsync(DateTime now)
    {
        if (_unavailable is not null)
        {
            await SetViewAsync(UnavailableView, Views.UnavailableCard(_unavailable), 20);
            foreach (var id in new[] { CalendarView, MailView, TasksView })
            {
                await ClearViewAsync(id);
            }
            foreach (var id in ActivityIds)
            {
                await EndActivityAsync(id);
            }
            return;
        }
        if (_snapshot is not { } snapshot)
        {
            return;
        }

        await ClearViewAsync(UnavailableView);
        await SetViewAsync(CalendarView, Views.CalendarCard(snapshot, now, Settings.MaxItems, _keys), 20);
        await SetViewAsync(MailView, Views.MailCard(snapshot, now, Settings.MaxItems, _keys), 21);
        await SetViewAsync(TasksView, Views.TasksCard(snapshot, now, Settings.MaxItems, _keys), 22);

        var meeting = Settings.MeetingActivity ? Views.UpcomingMeeting(snapshot, now, Settings.MeetingLead) : null;
        if (meeting is null)
        {
            await EndActivityAsync("meeting");
        }
        else
        {
            await SetActivityAsync(Views.MeetingActivity(meeting, now, _keys));
            if (meeting.Start <= now && now - meeting.Start < TimeSpan.FromMinutes(1) && _announced.Add(meeting.Ref))
            {
                await ShowHudAsync(Glyphs.Calendar, $"{Format.Clip(meeting.Subject, 80)} beginnt jetzt");
            }
        }
        if (_announced.Count > 100)
        {
            _announced.Clear();
        }

        if (snapshot.UnreadCount > 0)
        {
            await SetActivityAsync(Views.MailActivity(snapshot, now, Settings.ShowUnreadInNotch));
        }
        else
        {
            await EndActivityAsync("mail");
        }

        if (Views.TasksActivity(snapshot, now) is { } tasks)
        {
            await SetActivityAsync(tasks);
        }
        else
        {
            await EndActivityAsync("tasks");
        }

        _keys.Trim(snapshot.Mails.Select(m => m.Ref).Concat(snapshot.Meetings.Select(m => m.Ref)).Concat(snapshot.Tasks.Select(t => t.Ref)));
    }

    private async Task AnnounceNewMailAsync(OutlookSnapshot snapshot, DateTime now)
    {
        var known = _seenMail;
        var lastRead = _lastRead;
        _seenMail = snapshot.Mails.Select(m => m.Ref.EntryId).ToHashSet();
        _lastRead = now;
        if (known is null || !Settings.MailHud || _notificationsSeen)
        {
            return; // Beim ersten Lesen ist nichts neu.
        }
        // Nur wirklich neu Eingegangenes; eine alte, wieder auf ungelesen gesetzte E-Mail zählt nicht.
        var since = lastRead - TimeSpan.FromMinutes(2);
        var fresh = snapshot.Mails.Where(m => !known.Contains(m.Ref.EntryId) && m.Received >= since).ToList();
        if (fresh.Count == 1)
        {
            var mail = fresh[0];
            var sender = mail.Sender.Length > 0 ? mail.Sender : "Neue E-Mail";
            await ShowHudAsync(Views.Icons.Mail, Format.Clip($"{sender}: {mail.Subject}", 120));
        }
        else if (fresh.Count > 1)
        {
            await ShowHudAsync(Views.Icons.Mail, $"{fresh.Count} neue E-Mails");
        }
    }

    // ───────── An Brim, nur bei Änderungen ─────────

    private Task SetViewAsync(string id, Widget content, int order)
    {
        var json = JsonSerializer.Serialize(content, PluginJson.Options) + order;
        if (_views.TryGetValue(id, out var sent) && sent == json)
        {
            return Task.CompletedTask;
        }
        _views[id] = json;
        return SendAsync(() => _sink.SetViewAsync(id, content, order));
    }

    private Task ClearViewAsync(string id) =>
        _views.Remove(id) ? SendAsync(() => _sink.ClearViewAsync(id)) : Task.CompletedTask;

    private Task SetActivityAsync(ActivityParams activity)
    {
        var json = JsonSerializer.Serialize(activity, PluginJson.Options);
        if (_activities.TryGetValue(activity.Id, out var sent) && sent == json)
        {
            return Task.CompletedTask;
        }
        _activities[activity.Id] = json;
        return SendAsync(() => _sink.SetActivityAsync(activity));
    }

    private Task EndActivityAsync(string id) =>
        _activities.Remove(id) ? SendAsync(() => _sink.EndActivityAsync(id)) : Task.CompletedTask;

    private Task ShowHudAsync(string glyph, string text) =>
        SendAsync(() => _sink.ShowHudAsync(new HudParams { Glyph = glyph, Text = text, Announcement = text }));

    private Task LogAsync(string message, string level) => SendAsync(() => _sink.LogAsync(message, level));

    private static async Task SendAsync(Func<Task> send)
    {
        try
        {
            await send();
        }
        catch (Exception ex)
        {
            // Brim hat abgelehnt (ungültige Parameter, zu viele Nachrichten) oder ist weg: im Plugin-Log festhalten.
            Console.Error.WriteLine($"Brim: {ex.Message}");
        }
    }
}
