
[
    odl,
    uuid(C7D88A4E-2876-4dbf-B061-53523242341D),
    helpstring("Base Snarl extension interface.")
]

interface ISnarlExtension : stdole.IUnknown
{

    HRESULT Initialize(
        [out,retval] LONG *pVal);

    HRESULT TidyUp();

    HRESULT Start();

    HRESULT Stop();

    HRESULT GetConfigWindow(
        [out,retval] LONG *pVal);

    HRESULT Pulse();

    HRESULT GetInfo(
		[in,out] extension_info *Info);

    HRESULT LastError(
		[in,out] BSTR *Description);


}

