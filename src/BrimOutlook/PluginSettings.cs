using System.Globalization;
using System.Text.Json;

namespace BrimOutlook;

/// <summary>Einstellungen aus <c>plugin.json</c> mit den Werten des Nutzers. Zahlen und Schalter dürfen auch als Text kommen.</summary>
public sealed record PluginSettings
{
    public TimeSpan Refresh { get; init; } = TimeSpan.FromSeconds(30);
    public TimeSpan MeetingLead { get; init; } = TimeSpan.FromMinutes(15);
    public int MaxItems { get; init; } = 5;
    public bool ShowUnreadInNotch { get; init; }
    public bool MailHud { get; init; } = true;
    public bool MeetingActivity { get; init; } = true;

    public static PluginSettings From(IReadOnlyDictionary<string, JsonElement> values)
    {
        var defaults = new PluginSettings();
        return new PluginSettings
        {
            Refresh = TimeSpan.FromSeconds(Number(values, "refreshSeconds", defaults.Refresh.TotalSeconds, 10, 600)),
            MeetingLead = TimeSpan.FromMinutes(Number(values, "meetingLeadMinutes", defaults.MeetingLead.TotalMinutes, 1, 120)),
            MaxItems = (int)Number(values, "maxItems", defaults.MaxItems, 1, 10),
            ShowUnreadInNotch = Flag(values, "showUnreadInNotch", defaults.ShowUnreadInNotch),
            MailHud = Flag(values, "mailHud", defaults.MailHud),
            MeetingActivity = Flag(values, "meetingActivity", defaults.MeetingActivity),
        };
    }

    private static double Number(IReadOnlyDictionary<string, JsonElement> values, string key, double fallback, double min, double max)
    {
        if (!values.TryGetValue(key, out var value))
        {
            return fallback;
        }
        double? number = value.ValueKind switch
        {
            JsonValueKind.Number => value.GetDouble(),
            JsonValueKind.String when double.TryParse(value.GetString(), NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) => parsed,
            _ => null,
        };
        return number is { } n && double.IsFinite(n) ? Math.Clamp(n, min, max) : fallback;
    }

    private static bool Flag(IReadOnlyDictionary<string, JsonElement> values, string key, bool fallback)
    {
        if (!values.TryGetValue(key, out var value))
        {
            return fallback;
        }
        return value.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(value.GetString(), out var parsed) => parsed,
            JsonValueKind.Number => value.GetDouble() != 0,
            _ => fallback,
        };
    }
}
