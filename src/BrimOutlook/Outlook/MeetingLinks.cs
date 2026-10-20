using System.Text.RegularExpressions;
using System.Web;

namespace BrimOutlook.Outlook;

/// <summary>Findet den Link zum Beitreten einer Online-Besprechung (Teams, Zoom, Google Meet, Webex) in Ort oder Text eines Termins.</summary>
public static partial class MeetingLinks
{
    [GeneratedRegex(@"https?://[^\s<>""'\]]+", RegexOptions.IgnoreCase)]
    private static partial Regex UrlPattern();

    /// <summary>Erster Besprechungslink, zuerst im Ort, dann im Text; <c>null</c>, wenn keiner da ist.</summary>
    public static Uri? Find(params ReadOnlySpan<string?> texts)
    {
        foreach (var text in texts)
        {
            if (string.IsNullOrEmpty(text))
            {
                continue;
            }
            foreach (Match match in UrlPattern().Matches(text))
            {
                if (Normalize(match.Value) is { } uri && IsMeeting(uri))
                {
                    return uri;
                }
            }
        }
        return null;
    }

    /// <summary>Nur https-Links bekannter Dienste; Safe-Links von Microsoft werden ausgepackt.</summary>
    public static bool IsMeeting(Uri uri)
    {
        if (uri.Scheme != Uri.UriSchemeHttps)
        {
            return false;
        }
        var host = uri.Host.ToLowerInvariant();
        var path = uri.AbsolutePath.ToLowerInvariant();
        return host switch
        {
            "teams.microsoft.com" or "teams.cloud.microsoft" => path.StartsWith("/l/meetup-join/") || path.StartsWith("/meet/"),
            "teams.live.com" => path.StartsWith("/meet/"),
            "meet.google.com" => path.Length > 1,
            _ when host == "zoom.us" || host.EndsWith(".zoom.us") => path.StartsWith("/j/") || path.StartsWith("/my/") || path.StartsWith("/w/"),
            _ when host.EndsWith(".webex.com") => path.Contains("/j.php") || path.StartsWith("/meet/") || path.Contains("/join/"),
            _ => false,
        };
    }

    private static Uri? Normalize(string raw)
    {
        var trimmed = raw.TrimEnd('.', ',', ';', ':', ')', '>', '!', '?');
        if (!Uri.TryCreate(trimmed, UriKind.Absolute, out var uri))
        {
            return null;
        }
        // Safe Links (…safelinks.protection.outlook.com/?url=…) tragen das eigentliche Ziel im Parameter „url“.
        if (uri.Host.EndsWith(".safelinks.protection.outlook.com", StringComparison.OrdinalIgnoreCase)
            && HttpUtility.ParseQueryString(uri.Query)["url"] is { } target
            && Uri.TryCreate(target, UriKind.Absolute, out var unwrapped))
        {
            return unwrapped;
        }
        return uri;
    }
}
