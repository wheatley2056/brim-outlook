using Brim.Core.Plugins;
using Brim.Core.Widgets;
using Brim.Sdk;

namespace BrimOutlook;

/// <summary>Was das Plugin bei Brim anzeigt. Getrennt von <see cref="BrimPlugin"/>, damit Tests mitschreiben können.</summary>
public interface INotchSink
{
    Task SetActivityAsync(ActivityParams activity);
    Task EndActivityAsync(string id);
    Task ShowHudAsync(HudParams hud);
    Task SetViewAsync(string id, Widget content, int order);
    Task ClearViewAsync(string id);
    Task LogAsync(string message, string level);
}

internal sealed class BrimSink(BrimPlugin brim) : INotchSink
{
    public Task SetActivityAsync(ActivityParams activity) => brim.SetActivityAsync(activity);
    public Task EndActivityAsync(string id) => brim.EndActivityAsync(id);
    public Task ShowHudAsync(HudParams hud) => brim.ShowHudAsync(hud);
    public Task SetViewAsync(string id, Widget content, int order) => brim.SetViewAsync(id, content, order);
    public Task ClearViewAsync(string id) => brim.ClearViewAsync(id);
    public Task LogAsync(string message, string level) => brim.LogAsync(message, level);
}
