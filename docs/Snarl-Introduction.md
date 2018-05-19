
Welcome to Snarl!

Before we dive into Snarl, let's take a moment to talk about how Snarl is structured internally.  Snarl adopts a hierarchical approach, with individual **APIs** implemented by **Transports** which can then be used by one or more **Listeners**.

## Terms

### API

This is the core functionality built into Snarl that allows external applications to talk to Snarl.  Snarl currently supports two APIs, with a third - Oxide - in development.  No matter what mechanism is used to get a request to Snarl, that request must ultimately conform to one of the supported APIs for it to be successful.

### Transports

Transports translate external requests into a Snarl API request.  There are two ways to talk to Snarl - either via Windows IPC using WM_COPYDATA messages, or via network-based requests.

Some transports may take a request and pass it almost verbatim to the Snarl API - the V42 Win32 IPC transport and SNP/HTTP (v0 and v1) transports do exactly this - while some transports may implement their own format which needs considerable translation - SNP 3.1 is an example of such a transport.

Transports can also be written to translate existing protocols into a Snarl API request.  The GNTP transport is such an example and, although it doesn't exist yet, a transport which takes incoming SysLog UDP messages could very easily translate these into Snarl API requests.

### Listeners

While the API and Transports live in Snarl space; Listeners are created and configured by the end-user.  A listener combines a Transport with a specific TCP port which listens for incoming notification requests.

## The APIs

### Hydride

The current API used within Snarl is known as Hydride and was introduced in Snarl system version 42 (Release 2.4).  Hydride uses a very simple format which should be instantly recognisable:

```
notify?title=A Notification&text=Hello, world!
```













