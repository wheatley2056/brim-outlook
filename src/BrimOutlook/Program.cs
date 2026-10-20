// Outlook-Plugin für Brim: verbindet E-Mails, Termine und Aufgaben aus dem klassischen Outlook mit der Notch.
// Karten im Tab „Plugins“, der nächste Termin mit Countdown und „Beitreten“ in der Notch, neue E-Mails als HUD.
using System.Diagnostics;
using Brim.Sdk;
using BrimOutlook;
using BrimOutlook.Com;

if (!OperatingSystem.IsWindows())
{
    Console.Error.WriteLine("Das Outlook-Plugin läuft nur unter Windows.");
    return 1;
}

await using var brim = await BrimPlugin.ConnectAsync();
using var source = new ComOutlookSource();
var plugin = new OutlookPlugin(new BrimSink(brim), source, TimeProvider.System, OpenUrlAsync, PluginSettings.From(brim.Settings));

brim.ActionInvoked += (action, value) => _ = plugin.HandleActionAsync(action);
brim.SettingsChanged += settings => _ = plugin.ApplySettingsAsync(PluginSettings.From(settings));
brim.TabVisibilityChanged += visible =>
{
    if (visible)
    {
        _ = plugin.RefreshIfStaleAsync();
    }
};
brim.NotificationPosted += notification => _ = plugin.OnNotificationAsync(notification);

using var stop = new CancellationTokenSource();
var loop = plugin.RunAsync(stop.Token);
await brim.RunAsync();

// Brim gibt 2 Sekunden zum Beenden; auf einen hängenden Outlook-Aufruf wird nicht gewartet.
stop.Cancel();
await Task.WhenAny(loop, Task.Delay(500));
return 0;

static Task OpenUrlAsync(Uri link)
{
    Process.Start(new ProcessStartInfo(link.AbsoluteUri) { UseShellExecute = true })?.Dispose();
    return Task.CompletedTask;
}
