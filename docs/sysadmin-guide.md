# System Administration Guide

> ![Info](http://fullphat.net/docs/icons/info.png) _Note that this documentation applies to Snarl R5.0 Beta 6 and later._

## Introduction

Snarl provides system administrators with granular control over how local installations can be configured and secured against unauthorised change.  Individual installations of Snarl can be controlled via a single configuration files stored on a web server.  The web server will typically be within the same corporate environment, but it need not be.

> ![Info](http://fullphat.net/docs/icons/info.png) _This document is aimed at system administrators in corporate or educational environments who wish to control access to deployed instances of Snarl running on multiple computers._


## Overview

* At startup, Snarl wll look for a file called `redirect.rc` in its working directory
* If this file exists, Snarl will attempt to load the configuration referred to by the `url` entry in the `targets` section
* If this file doesn't exist, or Snarl was unable to load the referenced file, it will look for a file called `sysadmin.json` in the working directory

## The Configuration Files

### redirect.rc

This is used to redirect Snarl to a remote configuration file.  Using remote configuration files is preferred as :

* They are easier to maintain
* They are, by nature, not modifiable by the end user
* They allow the option of providing different configurations to different groups of users

If provided, `redirect.rc` should contain the following:

    [targets]
    url=..path to configuration file..

"path to configuration file" must be a fully-qualified URL to the configuration file to use.  For example `http://myserver/snarl/finance_team/admin.json`.

To avoid user interference with `redirect.rc`, administrators should ensure that users only have read access to this file on their computer.

### Sysadmin

This file details which features of Snarl are to be restricted.  It can either be located on a remote server (preferred) or it can be included in the same folder as Snarl itself.  If it's located in the same folder as Snarl it must be called `sysadmin.json`; if it's located on a remote server, it can be called anything.  Unlike `redirect.rc`, this configuration file must be formatted as JSON.

The following settings are currently available:

|`HideIcon`|`bool`|Controls whether Snarl's icon is displayed in the System Tray. Note that this value over-rides the icon control setting in the .snarl configuration but it's also important to note that setting this entry to zero will not force the icon to be visible|
|`InhibitPrefs`|`bool`|Controls whether Snarl's Preference Panel can be accessed. This setting prevents all access to the panel, including via the API and hot-key shortcut|
|`InhibitMenu`|`bool`|Controls whether or not the menu appears when the user right-clicks on the Snarl System Tray icon|
|`InhibitQuit`|`bool`|Controls whether or not Snarl can be stopped by the user. Enabling this option will prevent the user from closing Snarl via the System Tray Menu and most other methods but it does not protect the Snarl.exe process from being terminated via, for example, Task Manager.|

Some points to note:

* If a setting isn't provided, it's assumed to be `False`
* Setting an entry to `False` does not forcibly disable the restriction

## Example Configuration File

    {
        "HideIcon": false,
        "InhibitPrefs": true,
    }


## Worked Example

To try this out, do the following:

* Ensure Snarl is not running
* Create a text file called `sysadmin.json` in the same folder as Snarl
* Paste the following into the file and save it:

    {
        "HideIcon": false,
        "InhibitPrefs": true,
    }

* Launch Snarl
* Right-click Snarl's tray icon - you should see access to Snarl's preferences is blocked
