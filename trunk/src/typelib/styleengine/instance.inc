
[
    odl,
    uuid(DBB2113C-6701-4b62-B94A-246897037B9F),
    helpstring("Represents a specific instance of a particular style and scheme.")
]

interface IStyleInstance : stdole.IUnknown
{

    HRESULT AdjustPosition(
	[in, out] LONG *x,
	[in, out] LONG *y,
	[in, out] SHORT *Alpha,
	[in, out] BOOLEAN *Done);

    HRESULT GetContent(
	[out,retval] MImage **pval);

    HRESULT Pulse(
	[out,retval] BOOLEAN *Invalidated);

    HRESULT Show(
	[in] BOOLEAN Visible);

    HRESULT UpdateContent(
	[in,out] notification_info *NotificationInfo);


/*
    HRESULT Flags(
	[out,retval] SNARL_STYLE_FLAGS *pVal);
*/

}

