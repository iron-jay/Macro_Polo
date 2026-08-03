using System;
using System.Runtime.InteropServices;

namespace Macro_Polo.Core
{
    /// <summary>How Office came to load the add-in.</summary>
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    /// <summary>Why Office is unloading the add-in.</summary>
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    /// <summary>
    /// The interface Office calls to drive a COM add-in's lifetime.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Declared here rather than referenced from <c>extensibility.dll</c>. That primary interop
    /// assembly is not reliably present on machines that only have Office installed, and taking a
    /// dependency on it would reintroduce exactly the sort of "it works on the build machine"
    /// problem this conversion is meant to remove.
    /// </para>
    /// <para>
    /// Everything below is copied from the real declaration's metadata, and all of it matters. Two
    /// separate defects here each crashed the host process outright, with no managed exception and
    /// nothing in the log, because the failure happens inside the marshalling stub before any of
    /// this code runs:
    /// </para>
    /// <list type="bullet">
    ///   <item><description>
    ///     The interface is <b>dual</b>. The original carries no <see cref="InterfaceTypeAttribute"/>
    ///     at all, which means dual by default. Declaring it IDispatch-only exposes just the
    ///     IDispatch vtable, so Office's call lands past the end of the exposed slots.
    ///   </description></item>
    ///   <item><description>
    ///     <c>custom</c> is a <b>SAFEARRAY of VARIANT</b>. Without that descriptor the CLR falls
    ///     back to marshalling <see cref="Array"/> as a COM interface: it reads the SAFEARRAY
    ///     pointer Office passed as though it were an IDispatch pointer and dereferences it.
    ///   </description></item>
    /// </list>
    /// <para>
    /// The attributes are stated explicitly, never left to defaults, so none of this can be
    /// quietly "tidied up" back into a crash. MarshalingDescriptorTests compares every one of them
    /// against the primary interop assembly.
    /// </para>
    /// </remarks>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object application,
            [In] ext_ConnectMode connectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object addInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnDisconnection(
            [In] ext_DisconnectMode removeMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }
}
