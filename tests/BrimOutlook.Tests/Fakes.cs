using Brim.Core.Plugins;
using Brim.Core.Widgets;
using BrimOutlook.Outlook;

namespace BrimOutlook.Tests;

internal sealed class FakeOutlook : IOutlookSource
{
    public OutlookSnapshot Snapshot { get; set; } = new(0, [], [], []);
    public string? Unavailable { get; set; }
    public List<string> Calls { get; } = [];

    public Task<OutlookSnapshot> ReadAsync(DateTime now, DateTime until, int maxMails, CancellationToken cancellationToken = default) =>
        Unavailable is { } reason ? Task.FromException<OutlookSnapshot>(new OutlookUnavailableException(reason)) : Task.FromResult(Snapshot);

    public Task OpenAsync(ItemRef item) => Record($"open {item.EntryId}");
    public Task MarkReadAsync(ItemRef item) => Record($"read {item.EntryId}");
    public Task CompleteTaskAsync(ItemRef item) => Record($"done {item.EntryId}");
    public Task CreateAsync(OutlookKind kind) => Record($"new {kind}");
    public Task ShowFolderAsync(OutlookKind kind) => Record($"show {kind}");
    public Task LaunchAsync() => Record("launch");
    public void Dispose() { }

    private Task Record(string call)
    {
        Calls.Add(call);
        return Task.CompletedTask;
    }
}

internal sealed class RecordingSink : INotchSink
{
    public Dictionary<string, Widget> Views { get; } = [];
    public Dictionary<string, ActivityParams> Activities { get; } = [];
    public List<string> Huds { get; } = [];
    public int ActivitySets { get; private set; }

    public Task SetActivityAsync(ActivityParams activity)
    {
        ActivitySets++;
        Activities[activity.Id] = activity;
        return Task.CompletedTask;
    }

    public Task EndActivityAsync(string id)
    {
        Activities.Remove(id);
        return Task.CompletedTask;
    }

    public Task ShowHudAsync(HudParams hud)
    {
        Huds.Add(hud.Text ?? "");
        return Task.CompletedTask;
    }

    public Task SetViewAsync(string id, Widget content, int order)
    {
        Views[id] = content;
        return Task.CompletedTask;
    }

    public Task ClearViewAsync(string id)
    {
        Views.Remove(id);
        return Task.CompletedTask;
    }

    public Task LogAsync(string message, string level) => Task.CompletedTask;
}

internal static class WidgetTree
{
    /// <summary>Alle Bausteine eines Baums, für Abfragen in Tests.</summary>
    public static IEnumerable<Widget> All(Widget root)
    {
        yield return root;
        var children = root switch
        {
            StackWidget stack => stack.Children,
            ListWidget list => list.Items,
            _ => [],
        };
        foreach (var child in children.SelectMany(All))
        {
            yield return child;
        }
    }

    public static IEnumerable<string> Texts(Widget root) => All(root).OfType<TextWidget>().Select(t => t.Text);

    public static IEnumerable<string> Actions(Widget root) => All(root).OfType<ButtonWidget>().Select(b => b.ActionId);
}
