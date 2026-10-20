using System.Text.Json;

namespace BrimOutlook.Tests;

public class PluginSettingsTests
{
    [Fact]
    public void ReadsNumbersAndFlagsAlsoAsText()
    {
        var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(
            """{ "refreshSeconds": "45", "meetingLeadMinutes": 10, "maxItems": 99, "showUnreadInNotch": "true", "mailHud": false }""")!;
        var settings = PluginSettings.From(values);
        Assert.Equal(TimeSpan.FromSeconds(45), settings.Refresh);
        Assert.Equal(TimeSpan.FromMinutes(10), settings.MeetingLead);
        Assert.Equal(12, settings.MaxItems); // begrenzt
        Assert.True(settings.ShowUnreadInNotch);
        Assert.False(settings.MailHud);
        Assert.True(settings.MeetingActivity);
    }

    [Fact]
    public void FallsBackOnGarbage()
    {
        var values = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>("""{ "refreshSeconds": "oft", "mailHud": [] }""")!;
        Assert.Equal(new PluginSettings(), PluginSettings.From(values));
    }
}
