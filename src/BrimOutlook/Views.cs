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

    // ───────── Aktivität ─────────

    /// <summary>Höchster Rang, den Brim Plugins erlaubt; damit steht Outlook im Tab „Jetzt“ vor Akku, Kamera usw.</summary>
    public const int Rank = 9;

    /// <summary>Der Termin, der jetzt in die Notch gehört: beginnt innerhalb von <paramref name="lead"/> oder hat gerade begonnen.</summary>
    public static MeetingSummary? UpcomingMeeting(OutlookSnapshot snapshot, DateTime now, TimeSpan lead) =>
        snapshot.Meetings
            .Where(m => !m.AllDay && m.Start - now <= lead && now < m.Start + Linger && now < m.End)
            .OrderBy(m => m.Start)
            .FirstOrDefault();

    /// <summary>
    /// Eine Aktivität für alles: im Tab „Jetzt“ eine scrollbare Liste mit Termin, Kalender, Posteingang und
    /// Aufgaben. Kompakt erscheint sie nur, wenn ein Termin ansteht (Countdown) oder der Nutzer die ungelesenen
    /// E-Mails dort sehen will; sonst bleibt sie im Hintergrund.
    /// </summary>
    public static ActivityParams OutlookActivity(OutlookSnapshot snapshot, DateTime now, PluginSettings settings, ActionKeys keys)
    {
        var meeting = settings.MeetingActivity ? UpcomingMeeting(snapshot, now, settings.MeetingLead) : null;
        var items = new List<Widget>();
        if (meeting is not null)
        {
            items.Add(MeetingCard(meeting, now, keys));
        }
        items.Add(CalendarCard(snapshot, now, settings.MaxItems, keys));
        items.Add(MailCard(snapshot, now, settings.MaxItems, keys));
        items.Add(TasksCard(snapshot, now, settings.MaxItems, keys));
        var expanded = new ListWidget(items) { Spacing = 8, Scrollable = true };

        if (meeting is not null)
        {
            var countdown = Format.Countdown(meeting.Start, now);
            var color = meeting.Start <= now ? BrimColors.Charge : Accent;
            return new ActivityParams
            {
                Id = ActivityId,
                Rank = Rank,
                Accent = Accent,
                Leading = new IconWidget(Glyphs.Calendar) { Size = 14, Color = color },
                Trailing = new TextWidget(countdown) { Style = TextStyle.Caption, FontSize = 12, Color = color },
                Bead = new IconWidget(Glyphs.Calendar) { Size = 8, Color = color },
                Expanded = expanded,
                Announcement = $"{Format.Clip(meeting.Subject, 120)} {countdown}",
            };
        }
        return new ActivityParams
        {
            Id = ActivityId,
            Rank = Rank,
            Background = !(settings.ShowUnreadInNotch && snapshot.UnreadCount > 0),
            Accent = Accent,
            Leading = new IconWidget(Icons.Mail) { Size = 14, Color = Accent },
            Trailing = new TextWidget(snapshot.UnreadCount.ToString()) { Style = TextStyle.Mono, FontSize = 12 },
            Bead = new IconWidget(Icons.Mail) { Size = 8, Color = Accent },
            Expanded = expanded,
        };
    }

    /// <summary>Solange Outlook nicht erreichbar ist: Hinweis im Tab „Jetzt“, nichts in der kompakten Notch.</summary>
    public static ActivityParams UnavailableActivity(string reason) => new()
    {
        Id = ActivityId,
        Rank = Rank,
        Background = true,
        Leading = new IconWidget(Icons.Mail) { Size = 14, Color = BrimColors.Slate },
        Trailing = new TextWidget("aus") { Style = TextStyle.Caption, FontSize = 12 },
        Expanded = new ListWidget([UnavailableCard(reason)]) { Spacing = 8, Scrollable = true },
    };

    public const string ActivityId = "outlook";

    /// <summary>Der anstehende Termin groß, mit Countdown und „Beitreten“.</summary>
    private static Widget MeetingCard(MeetingSummary meeting, DateTime now, ActionKeys keys)
    {
        var countdown = Format.Countdown(meeting.Start, now);
        var color = meeting.Start <= now ? BrimColors.Charge : Accent;
        var key = keys.For(meeting.Ref);
        var children = new List<Widget>
        {
            new IconWidget(Glyphs.Calendar) { Size = 18, Color = color },
            TwoLines(Format.Clip(meeting.Subject, 120),
                Format.TimeRange(meeting) + (meeting.Location is { Length: > 0 } location ? " · " + Format.Clip(location, 80) : "")),
            new TextWidget(countdown) { Style = TextStyle.Mono, Color = color },
        };
        if (meeting.JoinLink is not null)
        {
            children.Add(new ButtonWidget($"join:{key}") { Kind = ButtonKind.Secondary, Glyph = Icons.Video, Label = "Beitreten" });
        }
        children.Add(SmallButton($"open:{key}", Glyphs.OpenInNew, "In Outlook öffnen"));
        return Card([Row(children)]);
    }

    // ───────── Bausteine ─────────

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
