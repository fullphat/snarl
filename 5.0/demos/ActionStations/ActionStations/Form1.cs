using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using fullphat.libSnarlWin32;

namespace ActionStations
{
    public partial class Form1 : Form
    {
        CallbackWindow callbackWindow;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            callbackWindow = new CallbackWindow();
            callbackWindow.OnCallbackReceived += new CallbackWindow.CallbackHandler(callbackWindow_OnCallbackReceived);
        }

        void callbackWindow_OnCallbackReceived(Win32Callback callback)
        {
            //System.Diagnostics.Debug.WriteLine("!!");
            //System.Diagnostics.Debug.WriteLine(callback.ToString());

            switch (callback.Type)
            {
                case Win32CallbackTypes.Expired:
                case Win32CallbackTypes.Dismissed:
                    reset();
                    break;

                case Win32CallbackTypes.ActionSelected:
                    System.Diagnostics.Debug.WriteLine(">" + callback.Uid);

                    switch (callback.Uid)
                    {
                        case "fury":
                            label1.BackColor = Color.Purple;
                            label1.ForeColor = Color.White;
                            label1.Text = "F U R Y   M O D E";
                            break;

                        case "deadpool":
                            label1.BackColor = Color.White;
                            label1.ForeColor = Color.DarkBlue;
                            label1.Text = "1-800-DEAD-POO\nService Unavailable";
                            break;

                    }

                    break;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Win32Request req = Win32RequestBuilder.NotifyRequest("", "", "Launch detection!", "Monitoring station Zulu Five Zero has detected potential launch plumes from sector NVZX-71.", "!system-critical");
            req.Content.Add("reply-to", callbackWindow.hWnd.ToString());

            req.Content.Add("action-1-id", "1");
            req.Content.Add("action-1-label", "Panic");
            req.Content.Add("action-1-callback", "uid:panic");
            //req.Content.Add("action-1-icon", "");

            req.Content.Add("action-2-id", "2");
            req.Content.Add("action-2-label", "Call Deadpool");
            req.Content.Add("action-2-icon", "stock:spiderman");
            req.Content.Add("action-2-callback", "uid:deadpool");

            req.Content.Add("action-3-id", "3");
            req.Content.Add("action-3-label", "Unleash Fury");
            req.Content.Add("action-3-icon", "stock:bomb");
            req.Content.Add("action-3-callback", "uid:fury");

            req.Content.Add("action-4-id", "4");
            req.Content.Add("action-4-label", "Snooze");
            req.Content.Add("action-4-icon", "stock:snooze");
            req.Content.Add("action-4-callback", "sys:snooze");

            Win32Response result = SnarlWin32.SendRequest(req);
            if (result.IsSuccess)
            {
                label1.BackColor = Color.Firebrick;
                label1.ForeColor = Color.Red;
                label1.Text = "-Launch Detected!-";
            }
            else
            {
                MessageBox.Show("Curses!", result.StatusCode.ToString(), MessageBoxButtons.OK);
            }

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            callbackWindow.Dispose();
        }

        void reset()
        {
            label1.BackColor = Color.OliveDrab;
            label1.ForeColor = Color.GreenYellow;
            label1.Text = "-System Offline-";
        }

    }
}
