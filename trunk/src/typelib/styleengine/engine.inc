
[
    odl,
    uuid(DDA4FBE0-6F8E-443c-AA86-7A1E592032E0),
    helpstring("A Style Engine provides any number of styles, each with their own schemes.")
]

interface IStyleEngine : stdole.IUnknown
{

    HRESULT Initialize(
        [out,retval] M_RESULT *pVal);

    HRESULT TidyUp();

    HRESULT Version(
        [out,retval] LONG *pVal);

    HRESULT Revision(
        [out,retval] LONG *pVal);

    HRESULT Name(
        [out,retval] BSTR *pVal);

    HRESULT Description(
        [out,retval] BSTR *pVal);

    HRESULT Date(
        [out,retval] BSTR *pVal);

    HRESULT LastError(
        [out,retval] BSTR *pVal);

    HRESULT CountStyles(
        [out,retval] LONG *pVal);

    HRESULT Path(
        [out,retval] BSTR *pVal);

    HRESULT StyleAt(
	[in] LONG Index,
	[in,out] style_info *Style);

    /* CreateInstance() only needs the Style name now: IStyleInstance.UpdateContent() receives the scheme name and flags */

    HRESULT CreateInstance(
	[in] BSTR StyleName,
        [out,retval] IStyleInstance **pVal);

    /* GetConfigWindow() should return a valid handle to a suitable configuration window for the specified style */

    HRESULT GetConfigWindow(
	[in] BSTR StyleName,
        [out,retval] LONG *pVal);

}

