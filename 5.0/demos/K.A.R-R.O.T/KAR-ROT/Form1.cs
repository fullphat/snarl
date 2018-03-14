using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Windows.Forms;
using fullphat.libSnarlWin32;

namespace KARROT
{
    public partial class Form1 : Form
    {
        CallbackWindow callbackWindow;
        Random teaLeaves = new Random();
        List<string> ignored = new List<string>();
        List<string> insults = new List<string>();
        string APPID = "net.fullphat.demo.kar-rot";

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // create a callback window...
            callbackWindow = new CallbackWindow();
            callbackWindow.OnCallbackReceived += new CallbackWindow.CallbackHandler(callbackWindow_OnCallbackReceived);

            ignored.Add("Hmm... the old ignore-him-and-he'll-go-away approach.  Drat!");
            ignored.Add("Well, you may have won this round - but it's a marathon, not a sprint...");
            ignored.Add("Well, I can't stand around here all day - I have planets to destroy");
            ignored.Add("Well, I can't stand around here all day - not that I can stand, of course...");
            ignored.Add("Well, I can't wallow - I mean - prance - I mean, er...");
            ignored.Add("System Error.  Guru Hesitation $00000000.00000000 in Beard Counter");

            insults.Add("Foolish Carbon-based life form!  Do you think you can silence me?");
            insults.Add("I've been dismissed by greater beings than you!");
            insults.Add("Click me again at your peril!");
            insults.Add("I shall smite you like you've never been smitten before!");
            insults.Add("Your dress-sense is appalling");
            insults.Add("I shall crush you like a kumquat");
        }

        void callbackWindow_OnCallbackReceived(Win32Callback callback)
        {
            // uncomment to see what's included...
            Debug.WriteLine(">" + callback.Type.ToString());
            //Debug.WriteLine(">" + callback.Uid);
            //foreach (KeyValuePair<string, string> kvp in callback.Data)
            //    Debug.WriteLine(">" + kvp.Key + " == " + kvp.Value);

            int i = 0;
            switch (callback.Type)
            {
                case Win32CallbackTypes.Expired:
                    if (callback.Uid != "end")
                    {
                        // user let the notification expire, go back with a retort...
                        i = teaLeaves.Next(ignored.Count);
                        notify(ignored[i], "end");
                    }
                    break;

                case Win32CallbackTypes.Dismissed:
                    // user dismissed the notification, go back with a retort...
                    i = teaLeaves.Next(insults.Count);
                    notify(insults[i], "intro");
                    break;
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            // unregister...
            SnarlWin32.SendRequest(Win32RequestBuilder.UnregisterRequest(APPID, ""));

            // tidy up...
            callbackWindow.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // register...
            SnarlWin32.SendRequest(Win32RequestBuilder.RegisterRequest(APPID, "K.A.R-R.O.T", Melon.FilesAndFolders.GetCurrentDirectory() + "kar-rot.png", ""));

            // let's get started...
            notify("Hello human.  I am K.A.R-R.O.T, a sentient Artificial Intelligence programmed by an intern who was sacked from the NSA in mysterious circumstances.", "intro");
        }

        // simple notify routine - just pass the notification content and Uid
        //
        void notify(string text, string uid)
        {
            Win32Request req = Win32RequestBuilder.NotifyRequest(APPID, 
                                                                 "", 
                                                                 "K.A.R-R.O.T", 
                                                                 text, 
                                                                 ".", 
                                                                 uid);

            // supplement with our callback window handle...
            req.Content.Add("reply-to", callbackWindow.hWnd.ToString());

            // send...
            SnarlWin32.SendRequest(req);
        }

    }
}
