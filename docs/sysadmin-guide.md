# System Administration Guide

## Introduction

> ![Info](http://fullphat.net/docs/icons/info.png) _Applies to Snarl R5.0 Beta 6 and later_

Snarl provides system administrators with greater control over how local installations can be configured and secured against unauthorised change.  Individual installations of Snarl can be controlled via a single - or several - configuration files stored on a web server.  The web server will typically be within the same corporate environment, but it need not be.

> ![Info](http://fullphat.net/docs/icons/info.png) _This document is aimed at system administrators in corporate or educational environments who wish to control access to deployed instances of Snarl running on multiple computers._


## Overview

Briefly, the process is as follows:

* At startup, Snarl wll look for a file called `redirect.rc` in its working directory
* This file should contain a single section (`targets`)
* Within the `targets` section, there should be an entry called `url`
* `url` should contain the full URL path to a JSON-formatted configuration file

## The Configuration Files

Snarl administration is based around simple configuration files which are easy to edit and distribute.

### redirect.rc

When Snarl starts it will look to see if there is a file called sysconfig.ssl in the current program directory (typically this will be /Program Files/full phat products/Snarl/). This file is a simple INI-style text file which should contain a single entry:

    [targets]
    url=http://myserver/snarl/finance_team/admin.json

Where `http://myserver/snarl/finance_team/admin.json` is an example URL to the JSON configuration file.

To avoid user interference with the sysconfig.ssl, administrators should ensure that users only have read access to this file on their computer.

## JSON configuration

This is a new file which contains administration settings specific to Snarl. Note that the contains of this file may be subsumed into the existing .snarl file in future revisions.  This file must exist in the supplied Snarl configuration folder for it to be loaded by Snarl.

The following settings are currently available:

* HideIcon=[0|1] - controls whether Snarl's icon is displayed in the System Tray. Note that this value over-rides the icon control setting in the .snarl configuration but it's also important to note that setting this entry to zero will not force the icon to be visible
* InhibitPrefs=[0|1] - controls whether Snarl's Preference Panel can be accessed. This setting prevents all access to the panel, including via the API and hot-key shortcut
* InhibitMenu=[0|1] - controls whether or not the menu appears when the user right-clicks on the Snarl System Tray icon
* InhibitQuit=[0|1] - controls whether or not Snarl can be stopped by the user. Enabling this option will prevent the user from closing Snarl via the System Tray Menu and most other methods but it does not protect the Snarl.exe process from being terminated via, for example, Task Manager.
* TreatSettingsAsReadOnly=[0|1] - changes made to Snarl's local configuration (i.e. anything that would usually be saved to the .snarl file) are not saved.

Some points to note:

For entries which require a zero or one value, one will enable the restriction. To leave a restriction disabled, it is acceptable to not omit the entry entirely
For entries which require a zero or one value, setting the restriction to zero does not forcibly disable that restriction.

## Worked Example

To try this out, do the following:

* Ensure Snarl is not running
* Create a text file called `redirect.rc` in Snarl's installation folder
* Edit the text file
* Save the file
* Create a folder called snarl in c:\
* Create a folder called etc in c:\snarl
* Create a text file called snarl.admin in c:\snarl\etc
* Edit snarl.admin and enter the following:

    InhibitPrefs=1
    TreatSettingsAsReadOnly=1

* Save snarl.admin
* Launch Snarl
* Right-click (or double-click) Snarl's tray icon - you should see access to Snarl's preferences is blocked
* Quit Snarl
* Open the c:\Snarl\etc folder
* Note that no Snarl configuration files have been written to this folder


