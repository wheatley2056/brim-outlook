using BrimOutlook.Outlook;

namespace BrimOutlook.Tests;

public class MeetingLinksTests
{
    [Theory]
    [InlineData("https://teams.microsoft.com/l/meetup-join/19%3ameeting_abc%40thread.v2/0?context=%7b%7d")]
    [InlineData("https://teams.microsoft.com/meet/1234567890?p=abc")]
    [InlineData("https://teams.live.com/meet/9876543210")]
    [InlineData("https://us02web.zoom.us/j/1234567890?pwd=xyz")]
    [InlineData("https://meet.google.com/abc-defg-hij")]
    [InlineData("https://firma.webex.com/firma/j.php?MTID=m123")]
    public void FindsKnownServices(string link)
    {
        Assert.Equal(new Uri(link), MeetingLinks.Find($"Hier beitreten: {link}\r\nBis gleich"));
    }

    [Fact]
    public void PrefersEarlierTextsAndIgnoresOtherLinks()
    {
        var found = MeetingLinks.Find(
            null,
            "Raum 3.14, https://intranet.example.com/raeume",
            "Agenda: https://example.com <https://teams.microsoft.com/l/meetup-join/19%3aabc/0>.");
        Assert.Equal("https://teams.microsoft.com/l/meetup-join/19%3aabc/0", found?.AbsoluteUri);
    }

    [Fact]
    public void UnwrapsSafeLinks()
    {
        var target = Uri.EscapeDataString("https://teams.microsoft.com/l/meetup-join/19%3aabc/0");
        var found = MeetingLinks.Find($"https://eur01.safelinks.protection.outlook.com/?url={target}&data=05");
        Assert.Equal("teams.microsoft.com", found?.Host);
    }

    [Theory]
    [InlineData("http://teams.microsoft.com/l/meetup-join/abc")]
    [InlineData("https://teams.microsoft.com/l/chat/abc")]
    [InlineData("https://zoom.us.evil.example/j/123")]
    [InlineData("kein Link")]
    public void RejectsEverythingElse(string text)
    {
        Assert.Null(MeetingLinks.Find(text));
    }
}
