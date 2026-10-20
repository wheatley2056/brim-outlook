using Brim.Core.Plugins;
using BrimOutlook.Outlook;
using Microsoft.Extensions.Time.Testing;

namespace BrimOutlook.Tests;

public class OutlookPluginTests
{
    private static readonly DateTime Start = new(2026, 10, 20, 9, 0, 0);

    private readonly FakeTimeProvider _time = new(new DateTimeOffset(Start, TimeSpan.Zero));
    private readonly FakeOutlook _outlook = new();
    private readonly RecordingSink _sink = new();
    private readonly List<Uri> _opened = [];
    private readonly OutlookPlugin _plugin;

    public OutlookPluginTests()
    {
        _time.SetLocalTimeZone(TimeZoneInfo.Utc);
        _plugin = new OutlookPlugin(_sink, _outlook, _time, uri =>
        {
            _opened.Add(uri);
            return Task.CompletedTask;
        }, new PluginSettings());
    }

    private DateTime Now => _time.GetLocalNow().DateTime;

    private static MailSummary Mail(string id, DateTime received, string sender = "Anna", string subject = "Angebot") =>
        new(new ItemRef(id), sender, subject, received, false);

    private static MeetingSummary Meeting(string id, DateTime start, Uri? link = null) =>
        new(new ItemRef(id), "Teamrunde", start, start.AddMinutes(30), "Raum 1", false, link);

    [Fact]
    public async Task ShowsCardsForMailCalendarAndTasks()
    {
        _outlook.Snapshot = new OutlookSnapshot(
            7,
            [Mail("m1", Start.AddMinutes(-5))],
            [Meeting("c1", Start.AddHours(2))],
            [new TaskSummary(new ItemRef("t1"), "Bericht schreiben", Start.Date.AddDays(-1), true)]);

        await _plugin.RefreshAsync();

        Assert.Equal(["calendar", "mail", "tasks"], _sink.Views.Keys.Order());
        Assert.Contains("Angebot", WidgetTree.Texts(_sink.Views["mail"]));
        Assert.Contains("und 6 weitere", WidgetTree.Texts(_sink.Views["mail"]));
        Assert.Contains("11:00–11:30", WidgetTree.Texts(_sink.Views["calendar"]));
        Assert.Contains("seit gestern überfällig", WidgetTree.Texts(_sink.Views["tasks"]));

        // Ungelesene nur im Tab „Jetzt“, fällige Aufgaben ebenso; der Termin ist noch zu weit weg.
        Assert.True(_sink.Activities["mail"].Background);
        Assert.True(_sink.Activities["tasks"].Background);
        Assert.False(_sink.Activities.ContainsKey("meeting"));
    }

    [Fact]
    public async Task MeetingCountsDownAndCanBeJoined()
    {
        var link = new Uri("https://teams.microsoft.com/l/meetup-join/19%3aabc/0");
        _outlook.Snapshot = new OutlookSnapshot(0, [], [Meeting("c1", Start.AddMinutes(20), link)], []);

        await _plugin.RefreshAsync();
        Assert.False(_sink.Activities.ContainsKey("meeting"));

        _time.Advance(TimeSpan.FromMinutes(8));
        await _plugin.RenderAsync();
        var activity = _sink.Activities["meeting"];
        Assert.Equal("in 12 min", ((Brim.Core.Widgets.TextWidget)activity.Trailing!).Text);

        var join = WidgetTree.Actions(activity.Expanded!).Single(a => a.StartsWith("join:"));
        await _plugin.HandleActionAsync(join);
        Assert.Equal([link], _opened);

        _time.Advance(TimeSpan.FromMinutes(12));
        await _plugin.RenderAsync();
        Assert.Contains("Teamrunde beginnt jetzt", _sink.Huds);

        // Fünf Minuten nach Beginn verschwindet er aus der Notch.
        _time.Advance(TimeSpan.FromMinutes(6));
        await _plugin.RenderAsync();
        Assert.False(_sink.Activities.ContainsKey("meeting"));
    }

    [Fact]
    public async Task AnnouncesOnlyNewMail()
    {
        _outlook.Snapshot = new OutlookSnapshot(1, [Mail("old", Start.AddHours(-3))], [], []);
        await _plugin.RefreshAsync();
        Assert.Empty(_sink.Huds);

        _time.Advance(TimeSpan.FromSeconds(30));
        _outlook.Snapshot = new OutlookSnapshot(2, [Mail("new", Now, "Ben", "Rechnung"), Mail("old", Start.AddHours(-3))], [], []);
        await _plugin.RefreshAsync();
        Assert.Equal(["Ben: Rechnung"], _sink.Huds);

        // Eine alte E-Mail, die wieder auf ungelesen gesetzt wurde, ist nicht neu.
        _time.Advance(TimeSpan.FromSeconds(30));
        _outlook.Snapshot = new OutlookSnapshot(3, [Mail("new", Now), Mail("old", Start.AddHours(-3)), Mail("older", Start.AddDays(-2))], [], []);
        await _plugin.RefreshAsync();
        Assert.Single(_sink.Huds);
    }

    [Fact]
    public async Task OutlookNotificationsReplaceOwnMailHud()
    {
        _outlook.Snapshot = new OutlookSnapshot(0, [], [], []);
        await _plugin.RefreshAsync();
        await _plugin.OnNotificationAsync(new NotificationParams("Outlook", "Ben", "Rechnung"));
        Assert.Equal(["Ben: Rechnung"], _sink.Huds);

        _time.Advance(TimeSpan.FromSeconds(30));
        _outlook.Snapshot = new OutlookSnapshot(1, [Mail("new", Now, "Ben", "Rechnung")], [], []);
        await _plugin.RefreshAsync();
        Assert.Single(_sink.Huds);
    }

    [Fact]
    public async Task ButtonsReachOutlook()
    {
        _outlook.Snapshot = new OutlookSnapshot(1, [Mail("m1", Start)], [], [new TaskSummary(new ItemRef("t1"), "Bericht", null, false)]);
        await _plugin.RefreshAsync();

        foreach (var action in WidgetTree.Actions(_sink.Views["mail"]).Concat(WidgetTree.Actions(_sink.Views["tasks"])))
        {
            await _plugin.HandleActionAsync(action);
        }

        Assert.Equal(["show Mail", "new Mail", "read m1", "open m1", "show Task", "new Task", "done t1", "open t1"], _outlook.Calls);
        Assert.Contains("Erledigt: Bericht", _sink.Huds);
    }

    [Fact]
    public async Task ShowsHintWhileOutlookIsClosed()
    {
        _outlook.Snapshot = new OutlookSnapshot(2, [Mail("m1", Start)], [], []);
        await _plugin.RefreshAsync();

        _outlook.Unavailable = "Outlook läuft nicht.";
        await _plugin.RefreshAsync();
        Assert.Equal(["outlook"], _sink.Views.Keys);
        Assert.Empty(_sink.Activities);
        Assert.Contains("launch", WidgetTree.Actions(_sink.Views["outlook"]));

        _outlook.Unavailable = null;
        await _plugin.RefreshAsync();
        Assert.Equal(["calendar", "mail", "tasks"], _sink.Views.Keys.Order());
    }

    [Fact]
    public async Task SendsOnlyChanges()
    {
        _outlook.Snapshot = new OutlookSnapshot(1, [Mail("m1", Start)], [Meeting("c1", Start.AddHours(3))], []);
        await _plugin.RefreshAsync();
        var sets = _sink.ViewSets;

        await _plugin.RenderAsync();
        await _plugin.RefreshAsync();
        Assert.Equal(sets, _sink.ViewSets);
    }
}
