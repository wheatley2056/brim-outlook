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

    private ActivityParams Outlook => _sink.Activities["outlook"];

    /// <summary>Abschnitte der Liste im Tab „Jetzt“.</summary>
    private IReadOnlyList<Brim.Core.Widgets.Widget> Sections => ((Brim.Core.Widgets.ListWidget)Outlook.Expanded!).Items;

    private IEnumerable<string> Texts => WidgetTree.Texts(Outlook.Expanded!);

    private static MailSummary Mail(string id, DateTime received, string sender = "Anna", string subject = "Angebot") =>
        new(new ItemRef(id), sender, subject, received, false);

    private static MeetingSummary Meeting(string id, DateTime start, Uri? link = null) =>
        new(new ItemRef(id), "Teamrunde", start, start.AddMinutes(30), "Raum 1", false, link);

    [Fact]
    public async Task ShowsEverythingScrollableInNow()
    {
        _outlook.Snapshot = new OutlookSnapshot(
            7,
            [Mail("m1", Start.AddMinutes(-5))],
            [Meeting("c1", Start.AddHours(2))],
            [new TaskSummary(new ItemRef("t1"), "Bericht schreiben", Start.Date.AddDays(-1), true)]);

        await _plugin.RefreshAsync();

        // Eine Aktivität mit höchstem Plugin-Rang; im Tab „Jetzt“ eine scrollbare Liste, nichts im Tab „Plugins“.
        Assert.Equal(["outlook"], _sink.Activities.Keys);
        Assert.Empty(_sink.Views);
        Assert.Equal(9, Outlook.Rank);
        Assert.True(((Brim.Core.Widgets.ListWidget)Outlook.Expanded!).Scrollable);
        Assert.Equal(3, Sections.Count);
        Assert.Contains("11:00–11:30", WidgetTree.Texts(Sections[0]));
        Assert.Contains("Angebot", WidgetTree.Texts(Sections[1]));
        Assert.Contains("und 6 weitere", WidgetTree.Texts(Sections[1]));
        Assert.Contains("seit gestern überfällig", WidgetTree.Texts(Sections[2]));

        // Der Termin ist noch zu weit weg: kompakt bleibt die Notch frei.
        Assert.True(Outlook.Background);
    }

    [Fact]
    public async Task MeetingCountsDownAndCanBeJoined()
    {
        var link = new Uri("https://teams.microsoft.com/l/meetup-join/19%3aabc/0");
        _outlook.Snapshot = new OutlookSnapshot(0, [], [Meeting("c1", Start.AddMinutes(20), link)], []);

        await _plugin.RefreshAsync();
        Assert.True(Outlook.Background);

        _time.Advance(TimeSpan.FromMinutes(8));
        await _plugin.RenderAsync();
        Assert.False(Outlook.Background);
        Assert.Equal("in 12 min", ((Brim.Core.Widgets.TextWidget)Outlook.Trailing!).Text);
        Assert.Equal(4, Sections.Count); // Termin groß obenauf

        var join = WidgetTree.Actions(Sections[0]).Single(a => a.StartsWith("join:"));
        await _plugin.HandleActionAsync(join);
        Assert.Equal([link], _opened);

        _time.Advance(TimeSpan.FromMinutes(12));
        await _plugin.RenderAsync();
        Assert.Contains("Teamrunde beginnt jetzt", _sink.Huds);

        // Fünf Minuten nach Beginn verschwindet er aus der Notch.
        _time.Advance(TimeSpan.FromMinutes(6));
        await _plugin.RenderAsync();
        Assert.True(Outlook.Background);
        Assert.Equal(3, Sections.Count);
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

        foreach (var action in WidgetTree.Actions(Sections[1]).Concat(WidgetTree.Actions(Sections[2])))
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
        Assert.True(Outlook.Background);
        Assert.Contains("Outlook läuft nicht.", Texts);
        Assert.Contains("launch", WidgetTree.Actions(Outlook.Expanded!));

        _outlook.Unavailable = null;
        await _plugin.RefreshAsync();
        Assert.Equal(3, Sections.Count);
    }

    [Fact]
    public async Task UnreadInCompactNotchOnRequest()
    {
        _outlook.Snapshot = new OutlookSnapshot(4, [Mail("m1", Start)], [], []);
        await _plugin.ApplySettingsAsync(new PluginSettings { ShowUnreadInNotch = true });
        Assert.False(Outlook.Background);
        Assert.Equal("4", ((Brim.Core.Widgets.TextWidget)Outlook.Trailing!).Text);
    }

    [Fact]
    public async Task SendsOnlyChanges()
    {
        _outlook.Snapshot = new OutlookSnapshot(1, [Mail("m1", Start)], [Meeting("c1", Start.AddHours(3))], []);
        await _plugin.RefreshAsync();
        var sets = _sink.ActivitySets;

        await _plugin.RenderAsync();
        await _plugin.RefreshAsync();
        Assert.Equal(sets, _sink.ActivitySets);
    }
}
