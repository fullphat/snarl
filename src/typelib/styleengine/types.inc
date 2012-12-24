
interface IStyleInstance;
interface IStyleEngine;

typedef enum S_STYLE_FLAGS {
    S_STYLE_MULTI_INSTANCE	= 0x00000000,	// can show multiple notifications
    S_STYLE_SINGLE_INSTANCE	= 0x00000001,	// can only show a single notification at any one time
    S_STYLE_IS_CONFIGURABLE	= 0x00000002,	// style has settings ui
    S_STYLE_IS_WINDOWLESS	= 0x00000004,	// style doesn't generate displayable content
    S_STYLE_DONT_QUEUE		= 0x00000008,	// for single instance styles, don't queue - replace immediately
    S_STYLE_CUSTOM_SHOW		= 0x00000010,	// custom show code
    S_STYLE_CUSTOM_HIDE		= 0x00000020,	// custom hide code
    S_STYLE_NO_DROPSHADOW	= 0x00000040,	// don't add a dropshadow
    S_STYLE_HANDLES_SOUNDS	= 0x00000080,	// will handle sound effects
    S_STYLE_WILL_RESIZE		= 0x00000100,	// will change size during custom show and hide stages
    S_STYLE_CLICK_THROUGH	= 0x00000200,	// style is always click-through irrespective of Snarl config
    S_STYLE_WANT_APP_NAME	= 0x00000400,	// UpdateContent() will include sending app name in the Title
	S_STYLE_V42_CONTENT		= 0x00000800,	// Title not used, Text contains unabridged notification content 
	S_STYLE_WANTS_VETO		= 0x00001000,	// For managed styles only - allows style to block notification appearing

    S_STYLE_PULSE_NEEDED	= 0x80000000	// wants a regular pulse

} S_STYLE_FLAGS;


typedef struct style_info { 
    BSTR		Name;
    BSTR		Description;
    LONG		Major;
    LONG		Minor;
    BSTR		Date;
    BSTR		Path;
    BSTR		Schemes;
    S_STYLE_FLAGS	Flags;
    BSTR		Copyright;
    BSTR		SupportEmail;
    BSTR		URL;
    BSTR		IconPath;

} style_info;




	/* S_NOTIFICATIONINFO->Flags */

typedef enum S_NOTIFICATION_FLAGS {
    S_NOTIFICATION_IS_PRIORITY	= 0x80000000,

} S_NOTIFICATION_FLAGS;

	/* Passed to UpdateContent() */

typedef struct notification_info { 
    BSTR			Title;
    BSTR			Text;
    BSTR			Icon;
    BSTR			Scheme;
    S_NOTIFICATION_FLAGS	Flags;

} notification_info;




