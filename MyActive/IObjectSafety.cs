using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
namespace MyActive
{
    //Guid唯一，不可变更，否则将无法通过IE浏览器的ActiveX控件的安全认证  
    [ComImport, GuidAttribute("073A987E-2A7C-4874-8BEE-321E04F4E84E")]
    [InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown)]
    public interface IObjectSafety
    {
        [PreserveSig]
        int GetInterfaceSafetyOptions(ref Guid riid, [MarshalAs(UnmanagedType.U4)] ref int pdwSupportedOptions, [MarshalAs(UnmanagedType.U4)] ref int pdwEnabledOptions);

        [PreserveSig()]
        int SetInterfaceSafetyOptions(ref Guid riid, [MarshalAs(UnmanagedType.U4)] int dwOptionSetMask, [MarshalAs(UnmanagedType.U4)] int dwEnabledOptions);
    }
}