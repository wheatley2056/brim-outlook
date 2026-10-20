using BrimOutlook.Outlook;

namespace BrimOutlook.Tests;

public class FormatTests
{
    private static readonly DateTime Now = new(2026, 10, 20, 9, 30, 0);

    [Theory]
    [InlineData(12, "in 12 min")]
    [InlineData(0.5, "in 1 min")]
    [InlineData(60, "in 1 h")]
    [InlineData(75, "in 1 h 15 min")]
    [InlineData(0, "jetzt")]
    [InlineData(-3, "seit 3 min")]
    public void Countdown(double minutes, string expected)
    {
        Assert.Equal(expected, Format.Countdown(Now.AddMinutes(minutes), Now));
    }

    [Theory]
    [InlineData(-3, "überfällig seit 3 Tagen")]
    [InlineData(-1, "seit gestern überfällig")]
    [InlineData(0, "heute fällig")]
    [InlineData(1, "morgen fällig")]
    [InlineData(4, "fällig Sa 24.10.")]
    public void Due(int days, string expected)
    {
        Assert.Equal(expected, Format.Due(Now.Date.AddDays(days), Now));
    }

    [Fact]
    public void ReceivedAndClip()
    {
        Assert.Equal("08:05", Format.Received(Now.Date.AddHours(8).AddMinutes(5), Now));
        Assert.Equal("gestern", Format.Received(Now.AddDays(-1), Now));
        Assert.Equal("So 18.10.", Format.Received(Now.AddDays(-2), Now));
        Assert.Equal("ohne Termin", Format.Due(null, Now));
        Assert.Equal("abc…", Format.Clip("abcdef", 4));
        Assert.Equal("a b", Format.Clip(" a\r\nb "));
    }
}
