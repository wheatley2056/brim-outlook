using Brim.Core;
using Brim.Core.Widgets;
using BrimOutlook.Outlook;

namespace BrimOutlook;

/// <summary>
/// Widgets für Brims Seite „Jetzt“: nächster Termin, Posteingang und Aufgaben in den Größen aus dem Manifest.
/// Reine Funktionen des Outlook-Stands wie <see cref="Views"/>; Aktionen sind dieselben wie in den Karten.
/// </summary>
public static class Widgets
{
    public const string Meeting = "meeting", Inbox = "inbox", Tasks = "tasks";

    public static Widget? Build(string id, TileSize size, OutlookSnapshot? snapshot, string? unavailable, DateTime now, ActionKeys keys)
    {
        if (snapshot is null)
        {
            return Unavailable(id, size, unavailable ?? "Outlook wird gelesen …");
        }
        return id switch
        {
            Meeting => MeetingWidget(snapshot, size, now, keys),
            Inbox => InboxWidget(snapshot, size, now, keys),
            Tasks => TasksWidget(snapshot, size, now, keys),
            _ => null,
        };
    }

    public static Widget MeetingWidget(OutlookSnapshot snapshot, TileSize size, DateTime now, ActionKeys keys)
    {
        var upcoming = snapshot.Meetings.Where(m => !m.AllDay && m.End > now).OrderBy(m => m.Start).ToList();
        if (upcoming.Count == 0)
        {
            return Empty(Glyphs.Calendar, size == TileSize.Small ? "Frei" : "Keine Termine mehr");
        }
        var next = upcoming[0];
        var countdown = next.IsRunning(now) ? "läuft" : Format.Countdown(next.Start, now);
        var soon = next.Start - now <= TimeSpan.FromMinutes(5);
        var color = soon ? BrimColors.Warning : BrimColors.Slate;
        if (size == TileSize.Small)
        {
            return Column(
                new TextWidget(next.Start.ToString("HH:mm")) { Style = TextStyle.Display, FontSize = 22 },
                new TextWidget(countdown) { Style = TextStyle.MonoCaption, Color = color });
        }
        var key = keys.For(next.Ref);
        var head = new List<Widget>
        {
            Lines(Format.Clip(next.Subject, 120), Format.TimeRange(next) + (next.Location is { Length: > 0 } l ? " · " + Format.Clip(l, 80) : "")),
        };
        if (next.JoinLink is not null)
        {
            head.Add(new ButtonWidget($"join:{key}") { Kind = ButtonKind.Secondary, Glyph = Views.Icons.Video, Label = "Beitreten" });
        }
        else
        {
            head.Add(new TextWidget(countdown) { Style = TextStyle.Mono, Color = color });
        }
        if (size != TileSize.Large)
        {
            return new StackWidget(head) { Spacing = 10 };
        }
        var rows = new List<Widget> { new StackWidget(head) { Spacing = 10 } };
        foreach (var meeting in upcoming.Skip(1).Take(3))
        {
            rows.Add(new StackWidget(
                new TextWidget(meeting.Start.ToString("HH:mm")) { Style = TextStyle.MonoCaption, FixedWidth = 40 },
                new TextWidget(Format.Clip(meeting.Subject, 120)) { Style = TextStyle.Secondary, FontSize = 12, Grow = 1 })
            {
                Spacing = 8,
            });
        }
        return Column(rows, CrossAlign.Stretch);
    }

    public static Widget InboxWidget(OutlookSnapshot snapshot, TileSize size, DateTime now, ActionKeys keys)
    {
        if (size == TileSize.Small)
        {
            return Column(
                new IconWidget(Views.Icons.Mail) { Size = 18, Color = snapshot.UnreadCount > 0 ? Views.Accent : BrimColors.Slate },
                new TextWidget(snapshot.UnreadCount.ToString()) { Style = TextStyle.Display, FontSize = 22, AccessibleName = Format.Unread(snapshot.UnreadCount) });
        }
        if (snapshot.Mails.Count == 0)
        {
            return Empty(Views.Icons.Mail, "Keine ungelesenen E-Mails");
        }
        var rows = new List<Widget>();
        foreach (var mail in snapshot.Mails.Take(size == TileSize.Large ? 4 : 1))
        {
            var key = keys.For(mail.Ref);
            rows.Add(new StackWidget(
                Lines(mail.Sender.Length > 0 ? mail.Sender : "(unbekannt)", mail.Subject.Length > 0 ? mail.Subject : "(kein Betreff)"),
                new TextWidget(Format.Received(mail.Received, now)) { Style = TextStyle.MonoCaption },
                new ButtonWidget($"open:{key}") { Glyph = Glyphs.OpenInNew, Size = 28, GlyphSize = 13, AccessibleName = "In Outlook öffnen" })
            {
                Spacing = 8,
            });
        }
        if (size == TileSize.Medium && snapshot.UnreadCount > 1)
        {
            rows.Add(new TextWidget($"und {snapshot.UnreadCount - 1} weitere") { Style = TextStyle.Caption });
            return Column(rows, CrossAlign.Stretch, spacing: 2);
        }
        return Column(rows, CrossAlign.Stretch);
    }

    public static Widget TasksWidget(OutlookSnapshot snapshot, TileSize size, DateTime now, ActionKeys keys)
    {
        if (snapshot.Tasks.Count == 0)
        {
            return Empty(Views.Icons.Tasks, "Keine offenen Aufgaben");
        }
        var rows = new List<Widget>();
        foreach (var task in snapshot.Tasks.Take(size == TileSize.Large ? 4 : 1))
        {
            var key = keys.For(task.Ref);
            var color = task.IsOverdue(now) ? BrimColors.Danger : task.IsDueToday(now) ? BrimColors.Warning : BrimColors.Slate;
            rows.Add(new StackWidget(
                new ButtonWidget($"done:{key}") { Glyph = Views.Icons.Checkbox, Size = 28, GlyphSize = 13, AccessibleName = $"{Format.Clip(task.Subject, 120)} erledigt" },
                Lines(task.Subject.Length > 0 ? task.Subject : "(kein Betreff)", Format.Due(task.Due, now), color))
            {
                Spacing = 6,
            });
        }
        if (size == TileSize.Medium && snapshot.Tasks.Count > 1)
        {
            rows.Add(new TextWidget($"und {snapshot.Tasks.Count - 1} weitere") { Style = TextStyle.Caption });
            return Column(rows, CrossAlign.Stretch, spacing: 2);
        }
        return Column(rows, CrossAlign.Stretch);
    }

    private static Widget Unavailable(string id, TileSize size, string reason)
    {
        var glyph = id switch { Meeting => Glyphs.Calendar, Tasks => Views.Icons.Tasks, _ => Views.Icons.Mail };
        if (size == TileSize.Small)
        {
            return Column(new IconWidget(glyph) { Size = 18, Color = BrimColors.Slate }, new TextWidget("Outlook aus") { Style = TextStyle.Caption });
        }
        return new StackWidget(
            new IconWidget(glyph) { Size = 18, Color = BrimColors.Slate },
            new TextWidget(reason) { Style = TextStyle.Secondary, FontSize = 12, Grow = 1 },
            new ButtonWidget("launch") { Kind = ButtonKind.Secondary, Label = "Starten" })
        {
            Spacing = 10,
        };
    }

    private static Widget Empty(string glyph, string text) => new StackWidget(
        new IconWidget(glyph) { Size = 18, Color = BrimColors.Slate },
        new TextWidget(text) { Style = TextStyle.Secondary, Grow = 1 })
    {
        Spacing = 10,
    };

    private static StackWidget Lines(string first, string second, string? secondColor = null) => new(
        new TextWidget(Format.Clip(first, 120)) { Style = TextStyle.Title, FontSize = 14 },
        new TextWidget(Format.Clip(second, 160)) { Style = TextStyle.Secondary, FontSize = 12, Color = secondColor })
    {
        Orientation = StackOrientation.Vertical,
        Align = CrossAlign.Stretch,
        Spacing = 1,
        Grow = 1,
    };

    private static StackWidget Column(params Widget[] children) => Column(children, CrossAlign.Center);

    private static StackWidget Column(IReadOnlyList<Widget> children, CrossAlign align, double spacing = 8) => new(children)
    {
        Orientation = StackOrientation.Vertical,
        Align = align,
        Justify = Justify.Center,
        Spacing = spacing,
        Grow = 1,
    };
}
