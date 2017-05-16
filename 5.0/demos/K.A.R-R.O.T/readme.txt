
 ____  __.       _____      __________         __________     ________     ___________
|    |/ _|      /  _  \     \______   \        \______   \    \_____  \    \__    ___/
|      <       /  /_\  \     |       _/  ______ |       _/     /   |   \     |    |   
|    |  \     /    |    \    |    |   \ /_____/ |    |   \    /    |    \    |    |   
|____|__ \ /\ \____|__  / /\ |____|_  /         |____|_  / /\ \_______  / /\ |____|   
========\/=\/=========\/==\/========\/=================\/==\/=========\/==\/==========

[ A Snarl 5.0 Demonstration App ]


IMPORTANT
=========

THIS SOFTWARE COMES WITH NO WARRANTY.  USE ONLY AS DIRECTED.


Synopsis
========

K.A.R-R.O.T demonstrates several new features of Snarl R5.0:

  o libSnarlWin32 - a wrapper library for the Win32 Transport

  o Win32 Transport V5 - included with Snarl R5.0

  o The ability to receive full callback information using the Win32 V5 transport


Usage
=====

This demo mimics an errant AI bot called K.A.R-R.O.T.  The user initiates the demo by clicking the red "LAUNCH" button, which generates an introductory notification.  If the user cancels the notification, or allows it to expire, a further notification will be generated.  The user can simply close the demo, or wait for the secondary notification to expire.  If the user dismisses the secondary notification, the cycle restarts.


Warning
=======

This demo shows how to respond to user interaction with notifications.  Snarl notification guidelines discourage reacting to the user dismissing a notification, or notification expiry.  Developers should carefully consider the logic of their application before reacting to these type of callbacks.




-------------------------------------------------------------------------------

K.A.R-R.O.T source code is provided free for use within other applications
Copyright (c) 2017 full phat products



