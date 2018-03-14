
1. [Introduction](#introduction)
1. [Status Codes](#status-codes)  
   i. [Client Errors](#client-errors)  
   ii. [Server Errors](#server-errors)  
   iii. [Informational Codes](#informational-codes)  
   iv. [Event Identifiers](#event-identifiers)
1. [Command Reference](#command-reference)  
   [`addaction`](#addaction)  
   [`addevent`](#addevent)  
   [`clearevents`](#clearevents)  
   [`forward`](#forward) (V47)  
   [`hello`](#hello)  
   [`hide`](#hide)  
   [`isvisible`](#isvisible)  
   [`notify`](#notify)  
   [`register`](#register)  
   [`remevent`](#remevent)  
   [`subscribe`](#subscribe) (V47)  
   [`unregister`](#unregister)  
   [`unsubscribe`](#unsubscribe) (V47)  


# Introduction
This document provides an in-depth guide to using the Snarl API.  It is mainly targeted at developers who wish to add low-level support for Snarl into their applications, or those who are creating an intermediary library which provides Snarl support for a particular programming language or environment.

Whatever the reason, there's no better way to understand how Snarl works than by understanding its API.

The [Transports](#transports) section provides details about the built-in transports which Snarl supports, and how to interface directly with them; the [Command Reference](#command-reference) includes the various commands Snarl supports, and the arguments and responses associated with them.

First, a quick re-cap on the request structure:

## Request Structure
Applications talk to Snarl by sending it a _request_.  A request is typically an action and a series of supporting parameters and values.  How the request is formatted depends on the transport itself.

The following is an example Win32 transport request:

`notify?app-sig=test.my_app&title=Hello, world!&text=This is a notification`

And as an SNP/3.1 request:

    SNP/3.1 NOTIFY
    app-id: test.my_app
    title: Hello, world!
    text: This is a notification
    END

## Transports
How the request is actually sent to Snarl depends on the transport that is used.  A transport is a low-level mechanism that handles the communication between the application and Snarl itself.  Transports are currently built directly into Snarl, but in the future it may be possible to create add-on transports.

A transport can be likened to a device driver and an operating system kernel, with the application playing the part of the device, and Snarl playing the part of the kernel.  The transport thus has two parts to it: a logical part which defines the protocol used, and a physical part which translates the protocol into something Snarl can understand.

Like a device driver, a transport handler must be robust and efficient and, also like device drivers, they run at a privileged level of execution (in this case, running in Snarl's process space).

Currently, Snarl supports four transports:

|Name|Description|
|----|-----------|
|Win32|The Win32 transport uses the Windows messaging subsystem, consequently this transport can only be used for local-to-local communication (that is, the client must be running on the same machine as the server).|
|[Snarl Network Protocol (SNP)](SNP-Overview)|Pronounced _"Snap"_, is a network protocol that supports cross-platform local-to-local and remote-to-remote communications.|
|[SNP over HTTP (SNP/HTTP)](SNP-HTTP-Developer-Guide)|Pronounced _"Snap over HTTP"_, provides access to the API using a RESTful HTTP approach.|
|Growl Notification Transport Protocol (GNTP)|A network protocol defined by the Growl notification application.|

# Status Codes
All requests will return a _status code_, with a status code of zero indicating success and any other value indicating an error or informational response.

How the status code is returned to the client very much depends on the transport used.  For example, the Win32 transport is by its own nature limited to returning a `uint32` value, so it will simply return the raw status code.  Other transports are not as constrained and will therefore be able to return not only return the status code, but other meta data, and even additional human-readable content to make debugging problems easier.

MOVE THIS BIT INTO WIN32 TRANSPORT:

> _**Status code convention:**_ although errors are returned as negative values, convention is to convert these into positive numbers before presenting them to the user.  The transport must therefore provide a mechanism for distinguishing between success (a positive status code) and failure (a negative status code).

Status codes are grouped into a series of blocks:

|Range|Block|
|-----|-----|
|0|Success|
|100 to 199|Client errors|
|200 to 249|Server errors|
|250 to 299|Informational codes|
|300 to 349|Event identifiers|

In the same way that it is down to the transport to communicate the response back to the client, not all transports provide all status codes.  For example, SNP 3.1 defines a `GOODBYE` response, but this response does not use the `SNARL_NOTIFY_GOODBYE` event identifier.  See the individual transport documentation for details of which codes are supported.

## Client Errors
These codes indicate the failure was at the client end.  For this reason, the server cannot return many of these codes directly (for example `SNARL_ERROR_HOST_UNKNOWN` could never be returned by the server), however they are provided so that intermediary libraries can return a consistent set of status codes to client applications:

CHANGE THESE TO USE THE SNARLLIB ENUM:

|Name|Code|Description|
|--------|---------|--------|
|`SNARL_ERROR_NOT_INITIALISED`|`100`|TBA|
|`SNARL_ERROR_FAILED`|`101`|General failure of some sort|
|`SNARL_ERROR_UNKNOWN_COMMAND`|`102`|The command was not recognised|
|`SNARL_ERROR_TIMED_OUT`|`103`|Snarl took too long to respond|
|`SNARL_ERROR_BAD_SOCKET`|`106`|The communication socket was closed unexpectedly|
|`SNARL_ERROR_BAD_PACKET`|`107`|Badly formed SNP request|
|`SNARL_ERROR_INVALID_ARG`|`108`|An invalid parameter was provided|
|`SNARL_ERROR_ARG_MISSING`|`109`|A required argument was missing|
|`SNARL_ERROR_SYSTEM`|`110`|Internal system error|
|`SNARL_ERROR_ACCESS_DENIED`|`121`|The command was not allowed|
|`SNARL_ERROR_HOST_UNKNOWN`|`TBA`|The specified host could not be resolved|
|`SNARL_ERROR_CONNECTION_FAILED`|`TBA`|A connection to the host and port could not be made|

## Server Errors
The following indicate that a successful connection to the server was made, but the server was not able to complete the request:

CHANGE THESE TO USE THE SNARLLIB ENUM:

|Name|Code|Description|
|--------|---------|--------|
|`SNARL_ERROR_NOT_RUNNING`|`201`|Snarl isn't running on the local (or, in the case of a network request, remote) computer|
|`SNARL_ERROR_NOT_REGISTERED`|`202`|An attempt was made to create a notification before the application was registered|
|`SNARL_ERROR_ALREADY_REGISTERED`|`203`|The application is already registered|
|`SNARL_ERROR_CLASS_ALREADY_EXISTS`|`204`|The class specified is already registered|
|`SNARL_ERROR_CLASS_BLOCKED`|`205`|The user has disabled notifications from this class|
|`SNARL_ERROR_CLASS_NOT_FOUND`|`206`|The specified class was not found|
|`SNARL_ERROR_NOTIFICATION_NOT_FOUND`|`207`|The specified notification was not found|
|`SNARL_ERROR_FLOODING`|`208`|The notification was not displayed as it would cause flooding of the display|
|`SNARL_ERROR_DO_NOT_DISTURB`|`209`|The user has enabled Do Not Disturb mode|
|`SNARL_ERROR_COULD_NOT_DISPLAY`|`210`|No enough screen space exists to display the notification|
|`SNARL_ERROR_AUTH_FAILURE`|`211`|Password mismatch|
|`SNARL_ERROR_DISCARDED`|`212`|The notification was discarded, usually because the sending application was in the foreground|
|`SNARL_ERROR_NOT_SUBSCRIBED`|`213`|Subscriber does not exist

## Event Identifiers
Event identifiers are used when Snarl notifies the sending that something has happened.  Typically this will be due to some element of user interaction, or lack thereof.  Each transport defines its own methods for communicating an event, and not all transports support events - hence the documentation for the particular transport in use should be consulted for further information.

EXTRA NOTES:

* Win32 sends event response to reply-to window, _GOODBYE is sent to subscribers.

CHANGE THESE TO USE THE SNARLLIB ENUM:

|Name|Code|Description|
|----|----|-----------|
|`SNARL_NOTIFY_GOODBYE`|301|Sent to subscribers when the server they're subscribed to is closing|
|`SNARL_NOTIFY_CLICK`|302|Deprecated: notification was right-clicked|
|`SNARL_NOTIFY_EXPIRED`|303|Notification timed out|
|`SNARL_NOTIFY_INVOKED`|304|Notification was clicked by the user|
|`SNARL_NOTIFY_MENU`|305|Deprecated: item was selected from the notification's menu|
|`SNARL_NOTIFY_EX_CLICK`|306|Deprecated: user clicked the middle mouse button on the notification|
|`SNARL_NOTIFY_CLOSED`|307|User clicked the notification's close gadget|
|`SNARL_NOTIFY_ACTION`|308|User selected an action from the notification's actions menu|


# Command Reference
The following sections describe the various commands supported by Snarl.  Not all versions of Snarl support all commands, as newer iterations introduce new features.  Where commands are only supported by certain versions of Snarl, this is called out separately.

Over time, commands may be renamed to make them more relevant (or, at least, less ambiguous).  The previous name(s) will remain supported, but new applications should use the current command.  Previous names are listed in the 'aliases' section.

## `addaction`
Adds an action to an existing notification (deprecated).

### Aliases
None

### Arguments

|Parameter|Requirement|Description|
|app-sig|Required|The application's signature.|
|uid|Required|The notification's unique identifier|
|label|Required|The text to be displayed to the user in the Actions drop-down menu|
|cmd|Required|Command to be invoked by Snarl (see notes)|
|password|Optional|Password used during registration.|

### Return Value
Return value is `SNARL_SUCCESS` if the action was added successfully, `SNARL_ERROR_MISSING_ARG` if either label or cmd are missing, `SNARL_ERROR_NOTIFICATION_NOT_FOUND` if the notification wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
* This command is now deprecated; applications should use `notify` to add actions to notifications
* The cmd argument can take one of three forms, as follows:

Type|Format|Action Taken|Examples
----|------|------------|--------
|Static callback|URL or path to file or folder|Specified file, folder or URL is launched by Snarl|c:\my_batch_file.bat http://www.getsnarl.info|
|Dynamic callback|`@{identifier}`|Snarl notifies the application with a `SNARL_NOTIFY_ACTION` callback, passing `{identifier}` back to the application|`@123456` `@email::1a2b3c4d5e`|
|System command|`!{command}`|Invokes a pre-defined system command - see here for details of currently implemented commands.|`!taskmgr` 

* The return value (and thus end result) is unpredictable if an action with the same command already exists.

## `addevent`
Adds a notification class to a previously registered application.

### Aliases
`addclass`

### Arguments

Parameter	Requirement	Description
app-sig	Required 	The application's signature used during registration.
id	Required	The identifier used by subsequent notifications.
password	Optional	The password used during registration.
name 	Optional 	The user-friendly name of the class.
enabled 	Optional 	Should be one if the class is to be enabled by default, or zero otherwise.  The default is one.
callback	Optional 	Default callback to be used if a notification created using this class does not specify one.
title 	Optional 	Default title to be used if a notification created using this class does not specify one.
text 	Optional	Default text to be used if a notification created using this class does not specify one.
icon 	Optional 	Default icon to be used if a notification created using this class does not specify one.
sound 	Optional 	Default sound to be used if a notification created using this class does not specify one.

### Return Value

Returns `SNARL_SUCCESS` if the notification was successfully hidden (or removed from the missed list), `SNARL_ERROR_NOT_REGISTERED` if the application wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
* As of V43, you can use `class-id` in place of `id` and `class-name` in place of `name`
* `name` is what the end user sees in Snarl's preferences window and thus should be more meaningful than the identifier.  If it is not specified, `id` is displayed instead

## clearactions
Removes all actions associated with the specified notification.

### Aliases
None

### Arguments

Parameter	Requirement	Description
app-sig	Required 	The application's signature.
uid 	Required 	The notification's unique identifier.
password	Optional	Password used during registration. 

### Return Value
Returns `SNARL_SUCCESS` if the successfully removed, `SNARL_ERROR_NOTIFICATION_NOT_FOUND` if the notification wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
Consideration should be taken when using this command as it will undoubtedly be confusing for an end-user if the actions assigned to a notification are removed while it is still visible on-screen.

## `clearevents`
Removes all event classes associated with a particular application.

### Aliases
`clearclasses`

### Arguments

|Parameter|Requirement|Description|
|---------|-----------|-----------|
|`app-sig`|Required|The application's signature|
|`password`|Optional|Password used during registration|

### Return Value
Returns `SNARL_SUCCESS` if the events were successfully removed, `SNARL_ERROR_NOT_REGISTERED` if the application wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
This is a more convenient way of doing `remclass?app-sig=some/app&all=1`

## `hello`
Effectively a no-op.  It can be useful however to:

* Confirm Snarl is running
* Confirm Snarl is accepting notifications via the particular transport
* Determine which version of Snarl is running (although `version` should be used for this)

### Arguments
None

### Return Value
Returns the version number of Snarl

## `hide`
Removes a notification from the screen

### Arguments

|Parameter|Requirement|Description|
|---------|-----------|-----------|
|`app-sig`|Required|The application's signature|
|`uid` or `token`|Required|The notification's unique identifier or token|
|`password`|Optional|Password used during registration|

### Return Value
Returns `SNARL_SUCCESS` if the notification was removed from the screen, `SNARL_ERROR_NOTIFICATION_NOT_FOUND` if the notification wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
Careful consideration should be given to programmatically removing a notification.  There can be good reason for removing a notification without any user interaction (for example, because the original reason for displaying the notification no longer exists), however often it can be confusing or frustrating for the user.  See the following blog post for more information.

## `isvisible`
Indicates whether or not the specified notification is still visible on-screen.

### Arguments

|Parameter|Requirement|Description|
|---------|-----------|-----------|
|`app-sig`|Required|The application's signature|
|`uid` or `token`|Required|The notification's unique identifier or token|
|`password`|Optional|Password used during registration|

### Return Value
Returns `SNARL_SUCCESS` if the notification is still visible, `SNARL_ERROR_NOTIFICATION_NOT_FOUND` if the notification wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
This command should be used carefully - or, at least, the result should be treated with suspicion as it's perfectly feasible that this command will return indicating that the notification is still visible on screen, only to have the notification immediately disappear because the user dismissed it, it expired, or it otherwise vanished.

## `notify`
Display a notification.

### Arguments

Parameter	Requirement	Description
app-sig	Required 	The application's signature used during registration.
password	Optional	The password used during registration.
id 	Optional 	The identifier of the class to use.
title 	Optional 	Notification title.
text	Optional 	Notification body text.
timeout 	Optional 	The amount of time, in seconds, the notification should remain on screen.
icon 	Optional	The icon to use, see the notes below for details of supported formats.
icon-base64 	Optional 	Base64-encoded bytes to be used as the icon.
sound 	Optional 	The path to a sound file to play. 
callback 	Optional 	Callback to be invoked when the user clicks on the main body area of the notification.
uid 	Optional	A unique identifier for the notification so the sending application can track events related to it and update it if required. 
priority	Optional	The urgency of the notification, 1 indicates high priority; -1 indicates low priority; 0 indicates normal priority, see Priorities in the Developer Guide for more information on how Snarl deals with different priorities. 
sensitivity	Optional	The sensitivity of the notification.  See the Sensitivity section in the Developer Guide for more information. 
replace-uid	Optional	The uid of an existing notification to replace. 
merge-uid	Optional	The uid of an existing notification to merge the content of this notification with.  
value-percent	Optional	A decimal percent value to be included with the notification.  Certain styles can display this value as a meter or other visual representation.  See Custom Values in the Developer Guide for more information. 
action 	Optional 	An action to add to the notification, see Actions in the Developer Guide for more information.
log	Optional	If set to "0", the notification is not logged in the missed list or history.

### Return Value
Returns a token (positive integer) if the notification was created successfully (not that it may not be displayed on screen depending on how the user has configured Snarl to respond), SNARL_ERROR_NOT_REGISTERED if the application wasn't found, or SNARL_ERROR_AUTH_FAILURE if an incorrect password was supplied.

### Notes

* Although individually optional, at least one of `title`, `text` or `icon` must be supplied for the command to succeed;
* A `timeout` of zero means the notification should remain on screen until the user dismisses it (behaviour commonly referred to as "sticky"); a `timeout` of -1 (the default) indicates the notification should use the default duration - this is the recommended practice;
* `icon` can be a fully qualified path to an image on a local or remote filesystem, a URL, a standard Snarl icon (if prefixed with '!') or an HICON (if prefixed with '#').  If no icon is provided, Snarl will attempt to use the application's own icon - see the Icons section in the Developer Guide for more information;
* `icon-base64` allows icon data to be included with the notification, which can be useful if the notification is being sent from one computer to another computer.  If both `icon` and `icon-base64` are supplied, `icon` will take precedence;
* If `id` is not specified, the notification will be displayed using the _All/Other Notifications_ class;
* If any of `title`, `text`, `timeout`, `callback` or `sound` are not supplied, then the corresponding class default will be used, if it exists;
* `sound` can be a WAV or MP3 file, or (if prefixed with an exclamation mark) a Windows system sound;
* `callback` is invoked when the user clicks the body area of the notification (any part of the notification not covered by a gadget); however, applications should consider using actions in preference to the callback feature of notifications, see the Notification Guidelines section in the Developer Guide for more information;
* Any number of `action` arguments may be included within a single notify request.  See Actions in the Developer Guide for more details;
* Depending on the state of the user at the time the notification appears, Snarl may alter the priority or timeout values of the notification before displaying it.  Consequently, these values should be considered as recommendations from the sending application but no more than that;
* `uid` can be used to uniquely identify the notification after it has been displayed by Snarl, for example to update its content, replace it or merge additional content in.  The `uid` is also included in an event generated by the notification;
* Support for `data-` arguments depends on the particular style used to display the notification.  See Custom Values in the Developer Guide for more information;
* The `log` command is useful for applications which generate transient notifications which are of no use to the users from a historical perspective.  Such notifications could include an hourly reminder or the changing of the currently playing music track.  However, generally speaking, it's good practice to allow the user to decide which notifications should be logged.  It should be noted that, as of Snarl V45, all notifications are written to `APP_DATA\full phat\snarl\snarl_log.txt`, irrespective of the user's preferences and the value of this argument.


## register
Registers an application with Snarl.

### Aliases
`reg`

### Arguments

Parameter	Requirement	Description
app-sig	Required 	The application's signature  
uid	Required	The notification's unique identifier
title 	Required 	The application's name 
icon 	Optional 	The icon to use - see notes for more details 
reply-to 	Optional 	Win32 transport only: the handle to a window (HWND) which Snarl should post events to 
reply-with 	Optional 	Win32 transport only: the message Snarl should use to post to the reply-to Window  
password	Optional	Password used during registration
keep-alive 	Optional	Prevent application from being removed during garbage collection

### Notes
* `title` is the actual name of the application which is displayed when the application first registers with Snarl and also appears in the applications list within Snarl's preferences UI. Unlike the signature, it does not need to be unique (however, logic dictates it should be);
* `icon` can be a fully qualified path to an image on a local or remote filesystem, a URL, a standard Snarl icon (if prefixed with '!') or an HICON (if prefixed with '#').  If no icon is provided, Snarl will attempt to use the executable icon - see the Icons section in the Developer Guide for more information;
* `reply-with` and `reply-to` are both optional but one cannot be specified without the other.  These parameters apply to the Win32 transport only - see the documentation on this transport for more information;
* From release 2.5.1 onwards, Snarl will periodically check for applications which appear to be orphaned - that is, an application which has registered but the process which registered it no longer exists.  For some environments, especially script and command-line ones, this can create a situation where a process launches to register an application, but then disappears once the command has been carried out.  To prevent such a registration from being automatically removed during the garbage collection, use `keep-alive=1` when registering.  Your application must unregister itself manually and must (if it uses one) manage the registration password correctly.

### Return Value
Returns a value greater than zero if the application was successfully registered, `SNARL_ERROR_ARG_MISSING` if any of the required arguments are missing, or `SNARL_ERROR_AUTH_FAILURE` if Snarl detects a re-registration without a matching password.

## remclass

Removes a particular notification class from a registered application.

Arguments

Parameter	Requirement	Description
app-sig	Required 	The application's signature  
password	Optional 	Password used during registration 
id 	Optional 	The identifier of the class to remove 
all 	Optional 	If 1, removes all classes associated with the specified application 

### Return Value
Return value is SNARL_SUCCESS if the class was removed okay, SNARL_ERROR_CLASS_NOT_FOUND if the specified class doesn't exist, SNARL_ERROR_NOT_REGISTERED if the application wasn't found, SNARL_ERROR_ARG_MISSING if neither all nor id is supplied, or SNARL_ERROR_AUTH_FAILURE if an incorrect password was supplied.

## `subscribe`
Requests to subscribe to all notifications or notifications from specific applications.  Multiple application identifiers can be provided and the filter can be set to include notifications from only those applications (inclusive), or all applications _except_ those applications (exclusive).  A unique identifier must be supplied, which is used in subsequent `unsubscribe` requests.

### Aliases
None.

### Arguments

|Parameter|Requirement|Description|
|---------|-----------|-----------|
|`filter`|Optional|One or more application identifiers to subscribe to, separated by semi-colons.|
|`filter-type`|Optional|Either `inclusive` or `exclusive`.  If not supplied, `inclusive` is assumed.|
|`forward-to`|Optional*|Url to forward notifications to.  One of `forward-to` or `reply-port` must be provided.|
|`reply-port`|Optional*|Local TCP port to forward notifications to.  One of `forward-to` or `reply-port` must be provided.|
|`uid`|Required|The unique identifier for the subscription.|

### Return Value
Returns `SUCCESS` if ok, `ALREADY_SUBSCRIBED` if already subscribed, `ARG_MISSING` if neither `reply-port` or `forward-to` provided, `INVALID_ARGS` if both `reply-port` and `forward-to` provided and `NOT_IMPLEMENTED` if the transport doesn't support it.

### Notes
* Requires Snarl 5.0 or later
* One of `reply-port` or `forward-to` must be supply, however both should not be supplied
* Not all transports support this request
* If `reply-port` is used, this should be listening for incoming forwarded notifications _before_ the subscribe request is issued
* The receiver _must_ respond to each forwarded notification
* The receiver _must_ accept a `goodbye` message.  This will indicate that no further notifications will be received from the server

## `test`
Displays a notification if Snarl is running in Debug Mode.

### Aliases
None.

### Arguments
None.

### Return Value
Returns `SNARL_SUCCESS` if Snarl is running in debug mode and the test notification was displayed successfully, `SNARL_UNKNOWN_COMMAND` if Snarl is running but is not in debug mode.



## `unregister`
Unregisters an application

### Aliases
`unreg`

### Arguments

|Parameter|Requirement|Description|
|---------|-----------|-----------|
|`app-sig`|Required|The application's signature|
|`password`|Optional|Password used during registration|

### Return Value
Returns `SNARL_SUCCESS` if the application was unregistered okay, `SNARL_ERROR_NOT_REGISTERED` if the application wasn't found, or `SNARL_ERROR_AUTH_FAILURE` if an incorrect password was supplied.

### Notes
(say about R5 and persistent registrations).

## `update`
Updates the content of an existing on-screen notification.

> **_This command is deprecated:_** to update an existing notification, applications should re-issue a `notify` command using the same `uid` - if the notification is still visible on-screen, it will be updated, otherwise a new notification will be created.

## `updateapp`
Updates specific details about an already registered application.

> _**This command is deprecated:**_ to update an existing registration, call `register` providing the updated details. 

## `version`
Effectively a no-op.  It can be useful however to:

* Confirm Snarl is running;
* Confirm Snarl is accepting notifications via the appropriate transport medium;
* Determine which version of Snarl is running.

### Arguments
None.

### Return Value
Returns the system version number of Snarl.
