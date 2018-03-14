# Introduction

Snarl provides significant flexibility and customisation around how notifications are presented.  With this flexibility however comes a responsibility on the part of the developer to ensure all notifications an application creates are of benefit to the user.  This means avoiding situations where:

* The user is distracted by a large number of notifications appearing at the same time
* Ambiguous or confusing notifications appear that require the user to stop what they're doing to investigate the source of the notification
* Notifications attempt to imitate operating system functions, or other applications
* Irrelevant or unsolicited content is provided to the user (e.g. advertising of associated products, requests for user feedback, etc.)
* The user's working context is ignored (e.g. not displaying non urgent notifications when the user is busy)

This document provides guidance on how applications should best present notifications to the user.  It deals with two key parts: the notification content itself, and the contextualisation of the notification.  These are only recommendations however; it is down to you as the developer to decide what is the right approach.

# Inside the Notification

## Title and Text

* **_Keep the title and text short:_** Notifications appear unsolicited and, as such, the user will want to be able to digest the meaning of the notification as quickly as possible without interrupting their flow.  Titles should be kept as short as possible, typically no more than four words; the notification text may be longer but should still be as concise as possible. If you need to pass a lot of information to the user, consider using a static callback or action which will direct the user to a document that contains it.

## Icons

* _**Use icons wherever possible:**_ icons are multi-lingual and can convey greater meaning with less effort.
* _**Choose wisely:**_ a well-chosen icon can convey the meaning of the notification without the user having to read the associated text.  Snarl includes a large number of stock icons which the user will become familiar with over time.
* **_Ensure the icon relates to the urgency of the notification:_** icons such as `!system-critical` should be reserved for only the most urgent or important notifications.  Consider using `!system-info` or `!system-warning` instead.

# Outside the Notification

## Dynamically Modifying Notifications

* **_Avoid programmatically hiding notifications:_** The hide command should be avoided as it is impossible for your application to tell whether the user has been able to fully digest the notification content. The only time hide may be of use is if the event the notification is alerting the user to no longer exists (for example, a “power disconnected” notification when the power supply has been restored).

## Priority

* **_Avoid high-priority notifications:_** Generally speaking the user should be left to determine which notifications constitute high-priority status. When deciding whether a notification should be displayed as high priority or not you should consider the context your notification will be displayed in and that, although you may feel the notification requires urgent from the user, the user may disagree.

Do Not Disturb (DND) mode can be enabled manually by the user or Snarl can be set to automatically enable it if the foreground application appears to be running in full-screen mode. In addition to this global on/off switch, the user can also define which notification classes should be treated as priority notifications so, in reality, all your application can do is recommend which notification classes should be treated as priority; the user has the final say in the matter.

## Duration

* _**Avoid sticky notifications:**_ notifications which do not disappear automatically quickly take up screen space and can be frustrating to the user as they must then stop what they’re doing, navigate to the notification and dismiss it manually.  The process is compounded if the notification has a default callback as the user must navigate specifically to the Close gadget on the notification – a much smaller area of screen than the notification itself.

* _**Use the Snarl default duration:**_ there are very few examples where using something other than the default notification duration is suitable and using the default duration ensures consistency across notifications and allows the user to determine how long they would like notifications to remain on screen.
