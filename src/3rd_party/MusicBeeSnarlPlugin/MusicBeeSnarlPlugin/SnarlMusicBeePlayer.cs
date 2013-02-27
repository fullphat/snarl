using System;
using Snarl.V44;
using System.Diagnostics;

namespace MusicBeeSnarlPlugin
{
    public sealed class SnarlMusicBeePlayer
    {
        private const int WindowReplyMsg = 0xBFFF; // Last WM_APP msg

        private enum SnarlCallbackAction : ushort
        {
            PlayerNext, // Default button action
            PlayerPrev,
            PlayerPause,
            PlayerStop,
        }

        private SnarlInterface snarl;
        private MusicBeePlugin.Plugin.MusicBeeApiInterface mbApiInterface;
        private String defaultIcon = "";


        public SnarlMusicBeePlayer(MusicBeePlugin.Plugin.MusicBeeApiInterface apiInterface)
        {
            mbApiInterface = apiInterface;
        }

        public void Register(String defaultIcon)
        {
            if (snarl != null)
                throw new Exception("Already registered.");

            this.defaultIcon = defaultIcon;

            snarl = new SnarlInterface();
            snarl.GlobalSnarlEvent += (s, e) =>
                {
                    if (e.GlobalEvent == SnarlGlobalEvent.SnarlLaunched)
                        DoRegister();
                };
            snarl.CallbackEvent += (s, e) =>
                {
                    if (e.SnarlEvent == SnarlStatus.CallbackInvoked)
                    {
                        mbApiInterface.Player_PlayNextTrack();
                    }
                    else if (e.SnarlEvent == SnarlStatus.NotifyAction)
                    {
                        if (e.Parameter == (ushort)SnarlCallbackAction.PlayerPrev)
                            mbApiInterface.Player_PlayPreviousTrack();
                        else if (e.Parameter == (ushort)SnarlCallbackAction.PlayerStop)
                            mbApiInterface.Player_Stop();
                        else if (e.Parameter == (ushort)SnarlCallbackAction.PlayerPause)
                            mbApiInterface.Player_PlayPause();
                    }
                };

            DoRegister();
        }

        private void DoRegister()
        {
            IntPtr windowHandle = mbApiInterface.MB_GetWindowHandle();

            snarl.RegisterWithEvents("MusicBee", "MusicBee", "dosgbdf", defaultIcon, windowHandle, WindowReplyMsg);
        }

        internal void Unregister()
        {
            snarl.Unregister();
        }

        internal void SongChanged(string artist, string trackTitle, string icon, string iconBase64)
        {
            var actions = new SnarlAction[3];
            actions[0] = new SnarlAction() { Label="Player: Prev", Callback="@" + ((int)SnarlCallbackAction.PlayerPrev).ToString() };
            actions[1] = new SnarlAction() { Label="Player: Pause", Callback = "@" + ((int)SnarlCallbackAction.PlayerPause).ToString() };
            actions[2] = new SnarlAction() { Label="Player: Stop", Callback="@" + ((int)SnarlCallbackAction.PlayerStop).ToString() };

            Int32 notificationId = snarl.Notify(uId: "SongChanged",
                                                title: artist,
                                                text: trackTitle,
                                                icon: icon,
                                                iconBase64: iconBase64,
                                                callback: "@" + (int)SnarlCallbackAction.PlayerNext,
                                                callbackLabel: "Next",
                                                actions: actions);
            Debug.WriteLine("MusicBeeSnarlPluginn: notificationId=" + notificationId);
        }

        internal void Test()
        {
            //for (int i = 0; i < 127; i++)
            //{
            //snarl.Notify(title: "Test",
            //     icon: "shell32.dll,-36");

            //snarl.Notify(title: "Test",
            //                 icon: defaultIcon + ",-32512");
            //snarl.Notify(title: "Test",
            //                     icon: defaultIcon + ",-32511");
            //snarl.Notify(title: "Test",
            //                     icon: defaultIcon + ",-32513");
            //snarl.Notify(title: "Test",
            //                     icon: defaultIcon + ",32512");
            //snarl.Notify(title: "Test",
            //                     icon: defaultIcon + ",32511");
            //snarl.Notify(title: "Test",
            //                     icon: defaultIcon + ",32513");
            //    Thread.Sleep(1050);
            //}
        }
    }
}
