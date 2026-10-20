using System.Text.Json;
using Brim.Core.Plugins;
using Brim.Core.Widgets;
using BrimOutlook.Outlook;

namespace BrimOutlook.Tests;

public class ViewsTests
{
    private static readonly DateTime Now = new(2026, 10, 20, 9, 0, 0);

    private static OutlookSnapshot Full()
    {
        var link = new Uri("https://teams.microsoft.com/l/meetup-join/19%3aabc/0");
        var long_ = new string('x', 5000);
        return new OutlookSnapshot(
            42,
            [.. Enumerable.Range(0, 10).Select(i => new MailSummary(new ItemRef($"m{i}"), i == 0 ? "" : $"Absender {i}", long_, Now.AddHours(-i * 10), i % 2 == 0))],
            [.. Enumerable.Range(0, 10).Select(i => new MeetingSummary(new ItemRef($"c{i}", Now.AddHours(i)), long_, Now.AddMinutes(i * 50 + 10), Now.AddMinutes(i * 50 + 40), long_, i == 3, link))],
            [.. Enumerable.Range(0, 10).Select(i => new TaskSummary(new ItemRef($"t{i}"), long_, i == 9 ? null : Now.AddDays(i - 3), i % 3 == 0))]);
    }

    /// <summary>Alles, was das Plugin schickt, muss Brims Prüfung bestehen und als JSON hin und zurück gehen.</summary>
    [Fact]
    public void EverythingPassesBrimsValidator()
    {
        var snapshot = Full();
        var keys = new ActionKeys();
        var meeting = Views.UpcomingMeeting(snapshot, Now, TimeSpan.FromMinutes(15))!;
        var activities = new[] { Views.MeetingActivity(meeting, Now, keys), Views.MailActivity(snapshot, Now, true), Views.TasksActivity(snapshot, Now)! };
        var widgets = new List<Widget>
        {
            Views.MailCard(snapshot, Now, 10, keys),
            Views.CalendarCard(snapshot, Now, 10, keys),
            Views.TasksCard(snapshot, Now, 10, keys),
            Views.UnavailableCard("Outlook läuft nicht."),
        };
        widgets.AddRange(activities.SelectMany(a => new[] { a.Leading, a.Trailing, a.Bead, a.Expanded }).OfType<Widget>());

        foreach (var widget in widgets)
        {
            Assert.Null(WidgetValidator.Validate(widget));
            var json = JsonSerializer.Serialize(widget, PluginJson.Options);
            Assert.Null(WidgetValidator.Validate(JsonSerializer.Deserialize<Widget>(json, PluginJson.Options)!));
        }
        Assert.All(WidgetTree.Actions(widgets[0]).Concat(WidgetTree.Actions(widgets[1])), id => Assert.True(id.Length <= 64));
    }

    [Fact]
    public void CalendarShowsTomorrowWhenTodayIsDone()
    {
        var tomorrow = new MeetingSummary(new ItemRef("c1"), "Planung", Now.Date.AddDays(1).AddHours(10), Now.Date.AddDays(1).AddHours(11), null, false, null);
        var card = Views.CalendarCard(new OutlookSnapshot(0, [], [tomorrow], []), Now, 5, new ActionKeys());
        Assert.Contains("Morgen", WidgetTree.Texts(card));
        Assert.Contains("Planung", WidgetTree.Texts(card));
    }

    [Fact]
    public void AllDayEventsStayOutOfTheNotch()
    {
        var allDay = new MeetingSummary(new ItemRef("c1"), "Urlaub", Now.Date, Now.Date.AddDays(1), null, true, null);
        Assert.Null(Views.UpcomingMeeting(new OutlookSnapshot(0, [], [allDay], []), Now, TimeSpan.FromMinutes(15)));
    }
}
