
interface ISnarlExtension;


typedef enum {
    SNARL_EXTN_IS_CONFIGURABLE	= 0x00000001,
	SNARL_EXTN_WANTS_PULSE		= 0x00000002,

} SNARL_EXTENSION_FLAGS;

typedef struct { 
    BSTR		Name;
    BSTR		Description;
    INT			Version;
    INT			Revision;
	SNARL_EXTENSION_FLAGS	Flags;

    BSTR		Path;
    BSTR		IconPath;
    MImage		*Icon;

	BSTR		Release;
    BSTR		Date;

	BSTR		Author;
    BSTR		Copyright;
    BSTR		SupportEmail;
    BSTR		URL;

} extension_info;










