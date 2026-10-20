using Brim.Core;
using Brim.Core.Plugins;
using Brim.Core.Widgets;
using BrimOutlook.Outlook;

namespace BrimOutlook;

/// <summary>Karten für den Tab „Plugins“ und Aktivitäten für die Notch. Reine Funktionen des Outlook-Stands.</summary>
public static class Views
{
    /// <summary>Outlook-Blau für Schimmer und Symbole.</summary>
    public const string Accent = "#3A96DD";

    /// <summary>So lange nach Beginn bleibt ein Termin in der Notch (zum Beitreten).</summary>
    public static readonly TimeSpan Linger = TimeSpan.FromMinutes(5);

    public static class Icons
    {
        public static readonly string Mail = G(0xE715);
        public static readonly string Read = G(0xE8C3);
        public static readonly string Tasks = G(0xE9D5);
        public static readonly string Checkbox = G(0xE739);
        public static readonly string Important = G(0xE8C9);
        public static readonly string Add = G(0xE710);
        public static readonly string Video = G(0xE8AA);
        public static readonly string Refresh = G(0xE72C);
        public static readonly string Warning = G(0xE7BA);

        private static string G(int codePoint) => char.ConvertFromUtf32(codePoint);
    }

    // ───────── Karten ─────────

    public static Widget MailCard(OutlookSnapshot snapshot, DateTime now, int maxItems, ActionKeys keys)
    {
        var rows = new List<Widget>
        {
            Header(Icons.Mail, "Posteingang", snapshot.UnreadCount == 0 ? null : snapshot.UnreadCount.ToString(), OutlookKind.Mail, "Neue E-Mail"),
        };
        if (snapshot.UnreadCount == 0)
        {
            rows.Add(Empty("Keine ungelesenen E-Mails"));
        }
        foreach (var mail in snapshot.Mails.Take(maxItems))
        {
            var key = keys.For(mail.Ref);
            var children = new List<Widget>();
            if (mail.Important)
            {
                children.Add(new IconWidget(Icons.Important) { Size = 12, Color = BrimColors.Danger, AccessibleName = "Wichtig" });
            }
            children.Add(TwoLines(mail.Sender.Length > 0 ? mail.Sender : "(unbekannt)", mail.Subject.Length > 0 ? mail.Subject : "(kein Betreff)"));
            children.Add(new TextWidget(Format.Received(mail.Received, now)) { Style = TextStyle.MonoCaption });
            children.Add(SmallButton($"read:{key}", Icons.Read, "Als gelesen markieren"));
            children.Add(SmallButton($"open:{key}", Glyphs.OpenInNew, "In Outlook öffnen"));
            rows.Add(Row(children));
        }
        AddMore(rows, snapshot.UnreadCount - Math.Min(maxItems, snapshot.Mails.Count));
        return Card(rows);
    }

    public static Widget CalendarCard(OutlookSnapshot snapshot, DateTime now, int maxItems, ActionKeys keys)
    {
        var today = snapshot.Meetings.Where(m => m.End > now && m.Start < now.Date.AddDays(1)).ToList();
        var tomorrow = snapshot.Meetings.Where(m => m.Start >= now.Date.AddDays(1) && m.Start < now.Date.AddDays(2)).ToList();
        var (title, shown) = today.Count > 0 || tomorrow.Count == 0 ? ("Heute", today) : ("Morgen", tomorrow);

        var rows = new List<Widget> { Header(Glyphs.Calendar, title, null, OutlookKind.Meeting, "Neuer Termin") };
        if (shown.Count == 0)
        {
            rows.Add(Empty("Keine Termine mehr heute"));
        }
        foreach (var meeting in shown.Take(maxItems))
        {
            var key = keys.For(meeting.Ref);
            var running = meeting.IsRunning(now);
            var children = new List<Widget>
            {
                new IconWidget(Glyphs.Dot) { Size = 8, Color = running ? BrimColors.Charge : Accent },
                new TextWidget(Format.TimeRange(meeting)) { Style = TextStyle.MonoCaption, FixedWidth = 84 },
                TwoLines(Format.Clip(meeting.Subject, 120), meeting.Location is { Length: > 0 } location ? Format.Clip(location, 120) : null),
            };
            if (meeting.JoinLink is not null)
            {
                children.Add(new ButtonWidget($"join:{key}") { Kind = ButtonKind.Secondary, Label = "Beitreten", AccessibleName = $"{Format.Clip(meeting.Subject, 120)} beitreten" });
            }
            children.Add(SmallButton($"open:{key}", Glyphs.OpenInNew, "In Outlook öffnen"));
            rows.Add(Row(children));
        }
        AddMore(rows, shown.Count - maxItems);
        return Card(rows);
    }

    public static Widget TasksCard(OutlookSnapshot snapshot, DateTime now, int maxItems, ActionKeys keys)
    {
        var urgent = snapshot.Tasks.Count(t => t.IsOverdue(now) || t.IsDueToday(now));
        var rows = new List<Widget> { Header(Icons.Tasks, "Aufgaben", urgent == 0 ? null : urgent.ToString(), OutlookKind.Task, "Neue Aufgabe") };
        if (snapshot.Tasks.Count == 0)
        {
            rows.Add(Empty("Keine offenen Aufgaben"));
        }
        foreach (var task in snapshot.Tasks.Take(maxItems))
        {
            var key = keys.For(task.Ref);
            var color = task.IsOverdue(now) ? BrimColors.Danger : task.IsDueToday(now) ? BrimColors.Warning : null;
            var children = new List<Widget>
            {
                SmallButton($"done:{key}", Icons.Checkbox, $"{Format.Clip(task.Subject, 120)} als erledigt markieren"),
                new StackWidget(
                    new TextWidget(Format.Clip(task.Subject.Length > 0 ? task.Subject : "(kein Betreff)", 120)),
                    new TextWidget(Format.Due(task.Due, now)) { Style = TextStyle.Caption, Color = color })
                {
                    Orientation = StackOrientation.Vertical,
                    Spacing = 2,
                    Align = CrossAlign.Stretch,
                    Grow = 1,
                },
            };
            if (task.Important)
            {
                children.Add(new IconWidget(Icons.Important) { Size = 12, Color = BrimColors.Danger, AccessibleName = "Wichtig" });
            }
            children.Add(SmallButton($"open:{key}", Glyphs.OpenInNew, "In Outlook öffnen"));
            rows.Add(Row(children));
        }
        AddMore(rows, snapshot.Tasks.Count - maxItems);
        return Card(rows);
    }

    /// <summary>Statt der drei Karten, solange Outlook nicht erreichbar ist.</summary>
    public static Widget UnavailableCard(string reason) => Card(
    [
        new StackWidget(
            new IconWidget(Icons.Mail) { Size = 14, Color = BrimColors.Slate },
            new TextWidget("Outlook") { Style = TextStyle.Secondary, FontSize = 12, Grow = 1 })
        {
            Spacing = 8,
        },
        new TextWidget(reason) { Style = TextStyle.Caption, Wrap = true },
        new StackWidget(
            new ButtonWidget("launch") { Kind = ButtonKind.Secondary, Label = "Outlook starten" },
            new ButtonWidget("refresh") { Kind = ButtonKind.Secondary, Label = "Erneut versuchen" })
        {
            Spacing = 8,
        },
    ]);

    // ───────── Aktivitäten ─────────

    /// <summary>Der Termin, der jetzt in die Notch gehört: beginnt innerhalb von <paramref name="lead"/> oder hat gerade begonnen.</summary>
    public static MeetingSummary? UpcomingMeeting(OutlookSnapshot snapshot, DateTime now, TimeSpan lead) =>
        snapshot.Meetings
            .Where(m => !m.AllDay && m.Start - now <= lead && now < m.Start + Linger && now < m.End)
            .OrderBy(m => m.Start)
            .FirstOrDefault();

    public static ActivityParams MeetingActivity(MeetingSummary meeting, DateTime now, ActionKeys keys)
    {
        var countdown = Format.Countdown(meeting.Start, now);
        var started = meeting.Start <= now;
        var color = started ? BrimColors.Charge : Accent;
        var key = keys.For(meeting.Ref);

        var actions = new List<Widget>();
        if (meeting.JoinLink is not null)
        {
            actions.Add(new ButtonWidget($"join:{key}") { Kind = ButtonKind.Secondary, Glyph = Icons.Video, Label = "Beitreten" });
        }
        actions.Add(new ButtonWidget($"open:{key}") { Kind = ButtonKind.Secondary, Label = "Öffnen" });

        var details = Format.TimeRange(meeting) + (meeting.Location is { Length: > 0 } location ? " · " + Format.Clip(location, 80) : "");
        return new ActivityParams
        {
            Id = "meeting",
            Rank = 6,
            Accent = Accent,
            Leading = new IconWidget(Glyphs.Calendar) { Size = 14, Color = color },
            Trailing = new TextWidget(countdown) { Style = TextStyle.Caption, FontSize = 12, Color = color },
            Bead = new IconWidget(Glyphs.Calendar) { Size = 8, Color = color },
            Expanded = new StackWidget(
                new StackWidget(
                    new IconWidget(Glyphs.Calendar) { Size = 16, Color = color },
                    new TextWidget(Format.Clip(meeting.Subject, 120)) { Style = TextStyle.Title, Grow = 1 },
                    new TextWidget(countdown) { Style = TextStyle.Mono, Color = color })
                {
                    Spacing = 8,
                },
                new TextWidget(details) { Style = TextStyle.Secondary },
                new StackWidget(actions) { Spacing = 8 })
            {
                Orientation = StackOrientation.Vertical,
                Spacing = 8,
                Align = CrossAlign.Stretch,
            },
            Announcement = $"{Format.Clip(meeting.Subject, 120)} {countdown}",
        };
    }

    /// <summary>Ungelesene E-Mails; im Tab „Jetzt“ immer, kompakt nur, wenn der Nutzer das will.</summary>
    public static ActivityParams MailActivity(OutlookSnapshot snapshot, DateTime now, bool compact) => new()
    {
        Id = "mail",
        Rank = 3,
        Background = !compact,
        Accent = Accent,
        Leading = new IconWidget(Icons.Mail) { Size = 14, Color = Accent },
        Trailing = new TextWidget(snapshot.UnreadCount.ToString()) { Style = TextStyle.Mono, FontSize = 12 },
        Bead = new IconWidget(Icons.Mail) { Size = 8, Color = Accent },
        Expanded = Summary(Icons.Mail, Format.Unread(snapshot.UnreadCount),
            snapshot.Mails.Take(3).Select(m => new SummaryLine(m.Sender, m.Subject, Format.Received(m.Received, now), null))),
    };

    /// <summary>Überfällige und heute fällige Aufgaben im Tab „Jetzt“; <c>null</c>, wenn keine drängt.</summary>
    public static ActivityParams? TasksActivity(OutlookSnapshot snapshot, DateTime now)
    {
        var urgent = snapshot.Tasks.Where(t => t.IsOverdue(now) || t.IsDueToday(now)).ToList();
        if (urgent.Count == 0)
        {
            return null;
        }
        var overdue = urgent.Count(t => t.IsOverdue(now));
        var color = overdue > 0 ? BrimColors.Danger : BrimColors.Warning;
        var title = urgent.Count == 1 ? "1 Aufgabe fällig" : $"{urgent.Count} Aufgaben fällig";
        return new ActivityParams
        {
            Id = "tasks",
            Rank = 2,
            Background = true,
            Leading = new IconWidget(Icons.Tasks) { Size = 14, Color = color },
            Trailing = new TextWidget(urgent.Count.ToString()) { Style = TextStyle.Mono, FontSize = 12, Color = color },
            Bead = new IconWidget(Icons.Tasks) { Size = 8, Color = color },
            Expanded = Summary(Icons.Tasks, title,
                urgent.Take(3).Select(t => new SummaryLine(t.Subject, null, Format.Due(t.Due, now), t.IsOverdue(now) ? BrimColors.Danger : BrimColors.Warning))),
        };
    }

    // ───────── Bausteine ─────────

    private sealed record SummaryLine(string Line, string? Detail, string Trailing, string? TrailingColor);

    private static Widget Summary(string glyph, string title, IEnumerable<SummaryLine> items)
    {
        var rows = new List<Widget>
        {
            new StackWidget(
                new IconWidget(glyph) { Size = 16, Color = Accent },
                new TextWidget(title) { Style = TextStyle.Title, Grow = 1 })
            {
                Spacing = 8,
            },
        };
        foreach (var (line, detail, trailing, trailingColor) in items)
        {
            rows.Add(new StackWidget(
                TwoLines(line, detail),
                new TextWidget(trailing) { Style = TextStyle.MonoCaption, Color = trailingColor })
            {
                Spacing = 8,
            });
        }
        return new StackWidget(rows) { Orientation = StackOrientation.Vertical, Spacing = 8, Align = CrossAlign.Stretch };
    }

    private static Widget Header(string glyph, string title, string? badge, OutlookKind kind, string newLabel)
    {
        var children = new List<Widget>
        {
            new IconWidget(glyph) { Size = 14, Color = Accent },
            new TextWidget(title) { Style = TextStyle.Secondary, FontSize = 12, Grow = 1 },
        };
        if (badge is not null)
        {
            children.Add(new TextWidget(badge) { Style = TextStyle.MonoCaption });
        }
        var target = kind.ToString().ToLowerInvariant();
        children.Add(SmallButton($"show:{target}", Glyphs.OpenInNew, $"{title} in Outlook zeigen"));
        children.Add(SmallButton($"new:{target}", Icons.Add, newLabel));
        return new StackWidget(children) { Spacing = 6 };
    }

    private static Widget TwoLines(string first, string? second)
    {
        var lines = new List<Widget> { new TextWidget(Format.Clip(first, 120)) };
        if (!string.IsNullOrEmpty(second))
        {
            lines.Add(new TextWidget(Format.Clip(second, 160)) { Style = TextStyle.Caption });
        }
        return new StackWidget(lines) { Orientation = StackOrientation.Vertical, Spacing = 2, Align = CrossAlign.Stretch, Grow = 1 };
    }

    private static Widget Row(IReadOnlyList<Widget> children) => new StackWidget(children) { Spacing = 8 };

    private static Widget SmallButton(string actionId, string glyph, string accessibleName) =>
        new ButtonWidget(actionId) { Glyph = glyph, Size = 28, GlyphSize = 14, AccessibleName = accessibleName };

    private static Widget Empty(string text) => new TextWidget(text) { Style = TextStyle.Caption };

    private static void AddMore(List<Widget> rows, int more)
    {
        if (more > 0)
        {
            rows.Add(new TextWidget(more == 1 ? "und 1 weiterer Eintrag" : $"und {more} weitere") { Style = TextStyle.Caption });
        }
    }

    private static Widget Card(IReadOnlyList<Widget> rows) => new StackWidget(rows)
    {
        Orientation = StackOrientation.Vertical,
        Spacing = 8,
        Align = CrossAlign.Stretch,
        Background = BrimColors.Graphite,
        CornerRadius = 14,
        Padding = (12, 10, 12, 10),
    };
}
