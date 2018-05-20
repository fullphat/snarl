
Before we dive into Snarl, let's take a moment to talk about how Snarl is structured internally.  

Snarl provides the **Core API** which external applications use to register, manage events and create notifications.  The Core API is accessed through **Transports** which are code modules that run within Snarl and translate incoming requests into Core API requests for processing.

### Core API

The Core API has developed significantly over time, and continues to do so.  The current version of the Core API was introduced in Snarl R2.4 (V42) and is known as **Hyride** or simply **V42**.  You can read more about the [Hydride API here](Hydride-API).

### Transports

Transports translate external requests into a Snarl API request.  There are two ways to talk to Snarl: via Windows IPC using `WM_COPYDATA` messages, or across the network.

Some transports may take a request and pass it almost verbatim to the Snarl API with litle to no translation inbetween, while some transports may implement their own format which needs considerable translation before the request can be submitted.  The V42 Win32 IPC transport and SNP/HTTP (v0 and v1) transports are good examples of the former; SNP 3.1 is an example of the latter.

Transports can also be written to translate existing protocols into a Snarl API request.  The GNTP transport is such an example and, although it doesn't exist yet, a transport which takes incoming SysLog UDP messages could very easily translate these into Snarl API requests.

### Listeners

While the API and Transports live in Snarl space; Listeners are created and configured by the end-user.  A listener combines a Transport with a specific TCP port which listens for incoming notification requests.














