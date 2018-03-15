# System Administration Guide

## Introduction

Beginning with Release 2.4, system administrators will have more control over local Snarl installations within their corporate environment. Control will initially be based around centrally managed configuration files but may be expanded into Group Policy Objects over time.

This page explains the current proposed functionality available to system administrators. Note that this is very much work in progress.

Target Audience
This document is aimed at system administrators in corporate or educational environments who wish to control access to deployed instances of Snarl running on multiple computers.

## Overview

Briefly, the process is as follows:

* Snarl looks for a file called 'sysconfig.ssl' in its working directory
* sysconfig.ssl should contain a single entry ("target") which should contain a path to Snarl's configuration folder
* If the path contained in sysconfig.ssl appears valid, it is used instead of the standard path (APPDATA\full phat\Snarl)
* If the path contained in sysconfig.ssl is valid, Snarl looks for a "snarl.admin" file and applies any restrictions found in it

## The Configuration Files

Snarl administration is based around simple configuration files which are easy to edit and distribute.

### sysconfig.ssl

When Snarl starts it will look to see if there is a file called sysconfig.ssl in the current program directory (typically this will be /Program Files/full phat products/Snarl/). This file is a simple INI-style text file which should contain a single entry:

   [targets]
   url=<path_to_sysadmin.json>

Where `path_to_sysadmin.json` is the URL to (but not including) the sysadmin.json file.

The configuration folder can have any name and can reside at any (accessible to the user) location. This allows system administrators to provide rudimentary role-based Snarl configuration by deploying multiple sysconfig.ssl files with different target folders to different groups of computers. For example, the Development Team may use configuration folder //server1/snarl/dev_team/ while the Call Centre staff may use the more restrictive //server1/snarl/callcentre/ folder instead. This is achieved by deploying different sysconfig.ssl files to the appropriate machines.

To avoid user interference with the sysconfig.ssl, administrators should ensure that users only have read access to this file on their computer.

## `sysadmin.json`
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

* Ensure R2.31d2 is installed but Snarl is not running
* Create a text file called sysconfig.ssl in Snarl's installation folder
* Edit the text file and add the single line target=c:\snarl
* Save the sysconfig.ssl file
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


