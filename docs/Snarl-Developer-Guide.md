# Snarl Developer Guide

1. [Introduction](#introduction)  
   i. [Displaying notifications](#displaying-notifications)  
   ii. [Registering and unregistering](#registering-and-unregistering)  
   iii. [Event classes](#event-classes)
1. [Features](#features)  
   i. [Icons](#icons)  
   ii. [Callbacks](#callbacks)  
   iii. [Playing a Sound](#playing-a-sound)  
   iv. [Setting the Duration](#setting-the-duration)  
   v. [Assigning a Priority](#assigning-a-priority)  
   vi. [Supplying Custom Data](#supplying-custom-data)  
   vii. [Headers](#headers)
1. [Advanced Topics](#advanced-topics)  
   i. [Application identifiers](#application-identifiers)  
   ii. [Icons](#icons-2)   
   iii. [Actions](#actions)     
   iv. [Indicating Sensitive Content](#indicating-sensitive-content)  
   v. [Determining Registration Status](#determining-registration-status)   

# Introduction

> ![Info](http://fullphat.net/docs/icons/info.png) _The examples in this document use the Snarl Win32 IPC transport.  You can try them out using the HeySnarlCS application which you can find in the Snarl tools folder_


## Displaying Notifications
Showing a notification on-screen is very straightforward:

    notify?title=Hello, world!

Will generate the following:

![Hello, world](http://fullphat.net/docs/dev/hello%20world.png)

Alternatively, you can have just an icon:

    notify?icon-stock=misc-chair

![Icon, only](http://fullphat.net/docs/dev/icon%20only.png)

And with a title and some text:

    notify?title=Hello, world!&text=Chairs, chairs, chairs...&icon-stock=misc-chair

![Chairs, chairs, chairs](http://fullphat.net/docs/dev/chairs%20chairs%20chairs.png)

You'll note that all of these notifications are displayed using the Grampf application - this is a built-in application provided by Snarl.  Grampf handles notifications not assigned to a registered application.

Unlike the built-in Snarl app, the user can disable Grampf, effectively blocking these "anonymous" notifications.  To properly display notifications, register an application and generate notifications using the application's identifier:

`notify?app-sig=chairs_r_us&icon-stock=misc-chair`


## Registering and Unregistering

.. to follow ..


## Event Classes

.. to follow ..


# Features

## Icons
Snarl can display icons from various sources, including:

* Image files
* Stock icons included with the Snarl distributable (and available via the web)
* URLs
* Windows resource files (DLL, EXE, etc.)
* Icon data included in the request itself

To display an icon, include the `icon` property in the request as follows:

    icon={data}

`data` can be a URL, path or a file, or stock icon that's included with the Snarl distributable.  Some examples include:

    icon=http://someserver.net/some/path/to/icon.png
    icon=!misc-hotdog
    icon=c:\my pics\icons\icon.png

See the [Advanced Topics](#icon-types) section for more details on specifying icons.

> _**Supported formats:** most image formats are supported, however PNG or JPEG are preferred as these include support for transparency.  Animated GIFs are not supported._


## Callbacks
Displaying a notification is something an application does - interacting with a notification is down to the user.  A user may dismiss a notification (by clicking anywhere on the notification body), leave it to expire naturally by not interacting with it at all, or "accept" the notification.  To be able to accept a notification, the notification must present a button that the user can click; to display the button requires that the notification is created with a _callback_.

This is a notification with a callback:

(notification with a callback)

And the request that generated it:

    notify?title=Your parcel has been delivered&icon=!misc-gift&callback-invoked=url:http://some.url/

As can be seen, the callback is assigned as part of a `notify` request using the following parameters:

|Name|Description|Status|
|----|-----------|------|
|`callback-invoked`|The action to take when the user accepts the notification|Required|
|`callback-invoked-args`|Additional arguments to include in the callback|Optional|
|`callback-label`|The label to use in the button|Optional|

### Callback Format
The callback action follows a standard pattern, as follows:

    callback-invoked={prefix}:{data}

The different types of callback are as follows:

|Prefix|Description|Data Value|
|------|-----------|----------|
|`file`|Launches a file|The full path to the file.  This can be a standard path (e.g. `c:\temp\myapp.exe`) or a UNC (e.g. `\\myserver\myshare\myapp.exe`)|
|`get`|Sends a HTTP GET request to a particular URL.  The notification's `uid` and `data-` parameters are included in the URL query|The URL to send the request to|
|`post`|Sends a HTTP POST request to a particular URL.  The notification's `uid` and `data-` parameters are included in the request as JSON content|The URL to send the request to|
|`uid`|Causes Snarl to send a message using the same transport that created the notification.  See Dynamic Callbacks below.|Leave blank|
|`url`|Launches the specified URL using the client's default browser|The URL to open|

### Dynamic Callbacks
If `uid:` is specified as the callback, Snarl will send a message to the application that created the notification using the same transport that generated the notification.  This will be sent as an _Out-of-Cycle message_ - see the individual transport documentation for how that is handled by the transport being used.

The message will include a status code of `NOTIFY_INVOKED` and will include the notification `uid` parameter and any `data-` parameters.


### Other Callbacks
In addition to triggering a callback when a notification is accepted by the user, client applications can also receive callbacks when a notification expires, is snoozed, or is dismissed by the user.  The following table describes these callbacks:

|Parameter|Description|Status Code|
|---------|-----------|-----------|
|`callback-dismissed`|Triggered when the notification is dismissed by the user|`NOTIFY_DISMISSED`|
|`callback-expired`|Triggered when the notification disappears naturally without being interacted with by the user|`NOTIFY_EXPIRED`|
|`callback-snoozed`|Triggered when the notification is snoozed by the user|`NOTIFY_SNOOZED`|

All of the above follow the same format as `callback-invoked` described above.  Additional arguments can also be supplied in the same way (e.g. `callback-expired-args`), however `callback-label` only applies if `callback-invoked` is used.








## Playing a Sound

Snarl can play a sound when it displays the notification, as follows:

`sound={prefix}:{data}`

Prefix can be stock or file.

> Although applications can provide their own sound for the notification to play, the user may configure Snarl to play a specific sound.


## Setting the Duration
The amount of time a notification should remain on-screen can be set using the `duration` property, as follows:

**`duration={seconds}`**

There are two special values which can be used: zero, to indicate that the notification should remain on-screen until specifically dismissed by the user, and -1 (the default) to indicate that the notification should remain on-screen for the default amount of time (which is user-configurable).

> Any value supplied should be considered as a _suggested_ duration.  Snarl may impose restrictions on the value supplied - especially if the notification has a low priority, or will be displayed non-interactively.  Generally speaking, applications should avoid prescribing the amount of time a notification should remain on screen and leave it to Snarl and the user to decide.

## Assigning a Priority
A priority is assigned using the `priority` property:

`priority={-1|0|1}`

A priority of -1 indicates the notification is of low importance and, as such, isn't detrimental to the user if they don't actually see if appear.  A priority of 1 indicates the notification is of high importance, in which case Snarl will attempt to ensure the notification is displayed to the user; a priority of zero is effectively a no-op - Snarl treats the notification as a regular notification.

With the release of Snarl 2.4, each application may only have one low priority notification may be on-screen at any one time.  When a new low priority notification appears, any existing low priority notification the application may have created is removed before the new one is displayed.  Additionally, low priority notifications are not displayed or logged as missed while the user is away or busy.

Conversely, high priority notifications are displayed even when Snarl's Do Not Disturb mode is enabled. For this reason alone they should be used very sparingly (see the Notification Guidelines section for more information).  Examples of notifications for which a high priority might be suitable include:

* Low power warning on a battery-powered system;
* Anti-Virus infection detection;
* UPS power supply activating.

Of course, if the user has to leave home at 6pm to catch his or her 11pm flight, they might consider an alarm or hourly reminder suitable for high priority status. The difference here though is that this would be a user choice, not an application choice. Generally speaking, it's best to assume your notification does not need to be a high priority one and leave it to the user to decide for themselves.




## Data and Headers
Notification requests can include additional data and header information.  Additional data is included by prefixing the key name with `data-`, as follows:

**`data-{key}={value}`**

For example:

`data-number=23&data-superhero=Deadpool`

Snarl ignores all data parameters, but they are included in callback responses (if the transport supports it), forwarded notifications and when notifications are redirected to external services.

Headers are considered to contain metadata about the notification itself.  They are included by prefixing the key name with `x-`, as follows:

**`x-{key}={value}`**

For example:

`x-clacks-overhead=GNU Terry Pratchett&x-received-from=Ankh Morpork`

Again, Snarl ignores headers, but they are included in forwarded notifications and those redirected to external services.  Headers are not included in callbacks however.



# Advanced Topics

## Application Identifiers
Each application that registers with Snarl must have a unique identifier.  The recommended way to create a suitable application identifier is to use [reverse domain name notation](https://en.wikipedia.org/wiki/Reverse_domain_name_notation) and include your company (or personal) name to ensure the identifier is suitably unique.  You may also use the [MIME media type](https://en.wikipedia.org/wiki/Media_type) format.

There are certain limitations on what can be used to create application identifiers:

* Use only alphanumeric characters, underscores, hyphens, full stops (periods) and forward slashes
* Avoid using characters outside of the 7 bit ASCII range
* Don't use spaces

For example: if your company's web domain is `acmeproducts.net` and your application is called "Hello, World!", then the application identifier could be any of the following:

`net.acmeproducts.hello_world`  
`net.acmeproducts.hello-world`  
`net.acmeproducts.helloworld`  
`application/x.vnd-acmeproducts.hello_world`  

The identifier is not usually displayed to the end user but is used behind-the-scenes by Snarl to manage rules the user creates that target the application.  It should therefore remain the same unless there is a particular need for it to change.  You might, for example, want to include a version number within it if there's a chance the user could be running different versions of your application concurrently and there's a need to distinguish between notifications from each version.

## Icon Types
Snarl 5.0 expands on how icon types can be specified, and how different types of icon can be provided.  These enhancements remain backwards-compatible with applications designed to work with earlier versions of Snarl.

### Icon Identifiers
The `icon` property has been enhanced to include _identifiers_.  These allow client applications to clearly indicate the type of icon being supplied.  Usually it can be inferred quite easily by Snarl as to what type of icon is being provided, however using a prefix ensures there is no ambiguity.

The `icon` property is therefore defined as:

**`icon={prefix}:{data}`**

Defined prefixes are as follows:

|Type|Description|Examples|
|----|-----------|--------|
|`file`|Absolute or relative path to the file to be used - can be local or remote.|`file:c:\temp\myicon.png`,  `file:\\myserver\myshare\myicon.png`|
|`phat64`|A Phat64-encoded string containing the image data itself.  Phat64 encoding is simply Base64 encoding, but modified to work with the Snarl API format.|n/a|
|`resource`|Path to resource file to use and (optionally) the index of the icon within the resource file to use.|  `resource:shell32.dll`, `resource:shell32.dll,34`, `resource:c:\myapp\mylib.dll,56`|
|`stock`|The name of a stock icon to use|`stock:misc-chair`, `stock:system-warning`|
|`url`|A valid URL to the image to use|`url:http://some.server.org/pics/some_pic.png`|

If an icon is specified and the prefix is not recognised, or isn't provided, Snarl will attempt to infer the icon type being provided in the same way that earlier versions did.  This ensures that applications designed to work with previous releases of Snarl remain compatible.

### Multiple Icons
Also introduced in Snarl 5.0 is the ability to define multiple icon types for a notification.  This can be useful if a notification is intended to be displayed in multiple ways (for example, it may be displayed on-screen, but the user may also forward it to another computer, or send it to a web-based push notification service).  While a local filesystem icon can be displayed on-screen, it may not be accessible from the remote computer, and certainly won't be accessible to the notification service.

Different icon types are specified using the `icon-*` property as follows:

**`icon-{type}={data}`**

The following `type` values are defined:

|Type|Description|Examples|
|----|-----------|--------|
|`file`|Absolute or relative path to the file to be used - can be local or remote.|`c:\temp\myicon.png`,  `\\myserver\myshare\myicon.png`|
|`phat64`|A Phat64-encoded string containing the image data itself.  Phat64 encoding is simply Base64 encoding, but modified to work with the Snarl API format.|n/a|
|`resource`|Path to resource file to use and (optionally) the index of the icon within the resource file to use.|  `shell32.dll`, `shell32.dll,34`, `c:\myapp\mylib.dll,56`|
|`stock`|The name of a stock icon to use|`misc-chair`, `system-warning`|
|`url`|A valid URL to the image to use|`http://some.server.org/pics/some_pic.png`|

Multiple icon types can be supplied, so it is possible to include a path to an icon on the local filesystem, a URL to the same icon, and the icon data encoded within the notification.  The displaying server can then decide which format is best to use (for example, a web-based push notification service may only be able to use the URL variant).  An example multi-icon request could look as follows:

**`notify?title=Icons galore!&icon-file=c:\temp\some_icon.png&icon-url=http://server.com/icons/icon.png&icon-phat64=..encoded_bytes..`**

### Inferred Icons
With the addition of support for multiple icon types within a notification, how the `icon` property is handled by Snarl has also been expanded to support multiple icon types.  Previously Snarl would take the supplied icon and include it in any forwarded or redirected notifications, now Snarl will create additional icon types based on the icon supplied, as follows:

|Type Supplied|Other Types Created|Examples|
|-------------|-------------------|--------|
|`file`|`phat64`|
|`stock`|`file`, `phat64`, `url`|
|`url`|None|

For example, if an application provides the following in a request:

`notify?title=Hello, world!&icon=stock:misc-hotdog`

Snarl will supply the following in any subsequently forwarded or redirected notifications:

    icon-stock=misc-hotdog
    icon-url=http://...
    icon-phat64={hex-encoded bytes}

> _**Priority:** The `icon` property is processed first by Snarl, then other icon types.  This allows applications to provide an icon in the traditional way (e.g. `icon=!misc-hotdog`) and then supplement it with additional icon types._


## Actions




## Indicating Sensitive Content
With the release of Snarl 2.5.1, an application can assign a sensitivity rating to a notification.  Currently, Snarl doesn't take any specific action based on the rating assigned; instead the value is passed through to the handling style, which may then take action.  The sensitivity is assigned as follows:

`sensitivity={0|16|32|48|64|80|96}`

An example may be a style which checks the user's IP address and then blocks or redacts the content of a notification if the user is not on their corporate LAN.

## Determining Registration Status
Snarl R5.0 removes the need for applications to register every time they wish to generate notifications.  While this removes a lot of extra application overhead, if the user decides to remove your application's registration information, subsequent notifications generated by your application may not be displayed.

To test if you application is (still) registered, send an 'empty' `notify` request, as follows:

`notify?app-sig=myapp`

or, if you've secured your application:

`notify?app-sig=myapp&password=mypassword`

If your application is still registered, you'll receive an `ArgMissing` (109) error; if your application is not registered with Snarl, you'll receive a `NotRegistered` (202) error instead.  Note that if your application is registered, but you've used the wrong password in the `notify` request, you'll receive an `AuthFailure` (211) error.

> _Play nicely:_ if your application is no longer registered, this may be because the user no longer wants to receive notifications from the application.  Consequently, 'nag' messages to re-register the application should be avoided.  Instead, one option may be to display an alert box asking if the user wants to re-register the application again, but provide the user with a "No, and don't ask me again" option.

> You might also want to consider the option of providing a separate executable or script that re-registers your application with Snarl - the user can then run this if they decide later they do want to receive notifications from your application, or if they unregister your application by accident.


***



## Forwarding
A forwarded notification is one that is sent from a source of notifications to Snarl.  Unlike standard notifications, forwarded notifications may be sent to a server without prior registration of the source application.  While this may seem to be a way to circumnavigate registering an application, developers should be aware of limitations around forwarded notifications:

* Forwarded notifications are handled by the Snarl built-in `ForwardedNotification` event class.  This may be disabled by the end user to prevent any client from forwarding notifications;

* The listener used by the client to transmit the forwarded notification to Snarl may be configured to reject forwarded notifications;

* The listener used by the client to transmit the forwarded notification to Snarl may be configured to require authorisation before accepting the message;

* Forwarded notifications displayed by Snarl receive a specific badge icon and application icon which cannot be changed by the client application (although the notification icon itself can).


## Subscriptions
A subscription is a request from a client to a server where the client asks to receive notifications from the server.  The notifications will be sent as forwarded messages, rather than traditional notifications.  The client can request all notifications, or only certain notifications, however the server ultimately decides which notifications it will forward.

### Subscribing
Subscriptions can take two forms: notifications can be forwarded to a URL as an HTTP `POST` request, or they can be sent as a message to a network port on the same computer as the client.

Subscriptions must include a unique identifier (uid) which is used when unsubscribing and should also be sent by the server when it needs to issue a `GOODBYE` response.  The uid should be a string of alphanumeric characters (a [GUID](https://msdn.microsoft.com/en-us/library/system.guid(v=vs.110).aspx) is a good example).  The uid is passed in the `uid` parameter.

The server may require authentication - if so, the authentication details must be included in the request header in the normal way.  See the [Security](security) section for further information.

> Note: subscription servers _should_ handle both types of forwarding, however they are not obligated to do so.  When sending a `subscribe` request, the client _must_ check the returned status to ensure the server accepted the request.

See [this guide](Subscribing-to-notifications) for more detail - and a worked example - on subscriptions.

### Unsubscribing
If the client no longer wants to receive notifications, it should unsubscribe.  It can do this by issuing an `unsubscribe` request from the same source that was used for the initial subscription.  It must also include the `uid` that was used in the initial subscription.

Once a client is unsubscribed, it may resubscribe using either the same or a different password.

### Handling Subscribers
An application may provide subscription handling functionality.  If it does so, it should obey the following:

* The server should support both `reply-port` and `forward-to` subscriptions.  If it doesn't support both, it should gracefully reject the request with a `NOT_IMPLEMENTED` failure response
* The server must support multiple subscribers.  How it manages individual subscribers is down to the server implementation itself
* The server should support authentication

Should a server need to stop serving notifications (for example, it may need to restart after being updated, or if the session it's running on is ending).  In this instance it must notify any existing subscribers that it is stopping by issuing a `Goodbye` response to each connected subscriber.  The `Goodbye` response must be sent to the `reply-port` or `forward-to` URL and must include the `uid` included in the individual subscription.

On receiving a `Goodbye` response, a client should assume it will no longer receive any further forwarded notifications from the server.  The client may choose to resubscribe, however it should wait a small amount of time (no less than 30 seconds) before doing so.

## Authentication
Authentication ensures that only trusted clients can communicate with a server as they both will share a common secret - specifically a password - that isn't communicated directly between the two.  Listeners may be password protected, meaning only clients that know the password can successfully issued requests.  Subscriptions can also be password protected so only trusted subscribers are allowed.

To authenticate a message, the password must be translated into a _key hash_ using a supported _hash algorithm_ and and _salt_.  This information must be included in the request.

### Algorithm
Defines the algorithm used to create the key hash. Supported algorithms are:

* `MD5`
* `SHA256`
* `SHA512`

### Key Hash
This is the hash of the result of hashing the password with a cryptographically secure salt.  The following steps describe how to translate a password into a key hash:

* The password should be converted into a UTF-8 byte array
* A cryptographically secure (between 8 and 32 bytes) salt should be generated
* The salt bytes are appended to the password bytes to form the _key basis_
* The _key_ is generated by computing the hash of the key basis using one of the supported hashing algorithms
* The _key hash_ is then produced by computing the hash of the key (using the same hashing algorithm) and hex-encoding as a string

### Salt
The cryptographically secure salt byte array used to generate the key hash.  Like the key hash, this should be encoded as a hexadecimal string.

## Encryption

To encrypt a message, the header must contain a valid authentication algorithm and associated key hash and salt plus a valid encryption algorithm and initialisation value (iv).

> Not all transports support encryption - refer to the individual transport documentation for details.
