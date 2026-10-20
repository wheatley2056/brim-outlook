using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace BrimOutlook.Com;

[SupportedOSPlatform("windows")]
internal static partial class Interop
{
    public const int RpcCallRejected = unchecked((int)0x80010001);
    public const int RpcRetryLater = unchecked((int)0x8001010A);
    public const int RpcDisconnected = unchecked((int)0x80010108);
    public const int RpcServerUnavailable = unchecked((int)0x800706BA);
    public const int RpcCallFailed = unchecked((int)0x800706BE);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    public static extern int CLSIDFromProgID(string progId, out Guid clsid);

    // Marshal.GetActiveObject gibt es seit .NET 5 nicht mehr.
    [DllImport("oleaut32.dll")]
    public static extern int GetActiveObject(ref Guid clsid, IntPtr reserved, [MarshalAs(UnmanagedType.IUnknown)] out object? instance);

    [DllImport("ole32.dll")]
    public static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    /// <summary>Laufende Instanz von <paramref name="progId"/>; <c>null</c>, wenn sie nicht läuft.</summary>
    public static object? GetActive(string progId, out bool installed)
    {
        installed = CLSIDFromProgID(progId, out var clsid) >= 0;
        return installed && GetActiveObject(ref clsid, IntPtr.Zero, out var instance) >= 0 ? instance : null;
    }

    /// <summary>Gibt ein COM-Objekt sofort frei, damit Outlook keine Elemente offen hält.</summary>
    public static void Release(object? value)
    {
        if (value is not null && Marshal.IsComObject(value))
        {
            Marshal.ReleaseComObject(value);
        }
    }
}

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig]
    int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo);

    [PreserveSig]
    int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);

    [PreserveSig]
    int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
}

/// <summary>
/// Wiederholt Aufrufe, die Outlook ablehnt, weil es gerade beschäftigt ist (offener Dialog, Synchronisation),
/// statt sofort mit <c>RPC_E_CALL_REJECTED</c> zu scheitern. Nach 5 Sekunden wird aufgegeben.
/// </summary>
[SupportedOSPlatform("windows")]
internal sealed class MessageFilter : IOleMessageFilter
{
    private const int ServerCallIsHandled = 0;
    private const int ServerCallRetryLater = 2;
    private const int PendingMsgWaitDefProcess = 2;

    public static void Register() => Interop.CoRegisterMessageFilter(new MessageFilter(), out _);

    public static void Revoke() => Interop.CoRegisterMessageFilter(null, out _);

    public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr interfaceInfo) => ServerCallIsHandled;

    public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType) =>
        rejectType == ServerCallRetryLater && tickCount < 5000 ? 250 : -1;

    public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType) => PendingMsgWaitDefProcess;
}
