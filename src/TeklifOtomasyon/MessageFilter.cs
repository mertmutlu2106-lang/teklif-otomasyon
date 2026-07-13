using System.Runtime.InteropServices;

namespace TeklifOtomasyon;

/// <summary>
/// Excel COM "meşgul / çağrı reddedildi" (RPC_E_CALL_REJECTED) durumunda
/// çağrıyı otomatik yeniden dener. Böylece biri o an hücrede yazarken
/// program çökmez, kısa süre bekleyip tekrar dener.
/// </summary>
internal static class MessageFilter
{
    public static void Register()
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            CoRegisterMessageFilter(new FilterImpl(), out _);
    }

    public static void Revoke()
    {
        if (Thread.CurrentThread.GetApartmentState() == ApartmentState.STA)
            CoRegisterMessageFilter(null, out _);
    }

    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);

    private sealed class FilterImpl : IOleMessageFilter
    {
        // Gelen çağrı → işle (SERVERCALL_ISHANDLED)
        int IOleMessageFilter.HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) => 0;

        // Çağrı reddedildi → RETRYLATER ise ~250 ms sonra tekrar dene
        int IOleMessageFilter.RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
        {
            const int SERVERCALL_RETRYLATER = 2;
            if (dwRejectType == SERVERCALL_RETRYLATER)
                return 250;  // >=0 → bu kadar ms sonra tekrar dene
            return -1;       // iptal
        }

        // Bekleyen mesaj → varsayılan işleme bırak
        int IOleMessageFilter.MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType) => 2;
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"),
     InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IOleMessageFilter
    {
        [PreserveSig] int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig] int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig] int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }
}
