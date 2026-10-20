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
            [.. Enumerable.Range(0, 20).Select(i => new MailSummary(new ItemRef($"m{i}"), i == 0 ? "" : $"Absender {i}", long_, Now.AddHours(-i * 10), true))],
            [.. Enumerable.Range(0, 20).Select(i => new MeetingSummary(new ItemRef($"c{i}", Now.AddHours(i)), long_, Now.AddMinutes(i * 20 + 10), Now.AddMinutes(i * 20 + 40), long_, i == 3, link))],
            [.. Enumerable.Range(0, 20).Select(i => new TaskSummary(new ItemRef($"t{i}"), long_, i == 9 ? null : Now.AddDays(i - 3), true))]);
    }

    /// <summary>Alles, was das Plugin schickt, muss Brims Prüfung bestehen und als JSON hin und zurück gehen, auch mit den meisten Einträgen.</summary>
    [Fact]
    public void EverythingPassesBrimsValidator()
    {
        var snapshot = Full();
        var settings = PluginSettings.From(new Dictionary<string, JsonElement> { ["maxItems"] = JsonSerializer.SerializeToElement(99) });
        var keys = new ActionKeys();
        var activities = new[]
        {
            Views.OutlookActivity(snapshot, Now, settings, keys),
            Views.OutlookActivity(snapshot, Now.AddHours(-3), settings, keys),
            Views.UnavailableActivity("Outlook läuft nicht."),
        };
        Assert.Equal(4, ((ListWidget)activities[0].Expanded!).Items.Count);

        foreach (var widget in activities.SelectMany(a => new[] { a.Leading, a.Trailing, a.Bead, a.Expanded }).OfType<Widget>())
        {
            Assert.Null(WidgetValidator.Validate(widget));
            var json = JsonSerializer.Serialize(widget, PluginJson.Options);
            Assert.Null(WidgetValidator.Validate(JsonSerializer.Deserialize<Widget>(json, PluginJson.Options)!));
        }
        Assert.All(WidgetTree.Actions(activities[0].Expanded!), id => Assert.True(id.Length <= 64));
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
