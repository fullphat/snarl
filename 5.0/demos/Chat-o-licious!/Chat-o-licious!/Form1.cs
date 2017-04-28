using System;
using System.Collections.Generic;
using System.ComponentModel;
//using System.Collections.Specialized;
//using System.Configuration;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using fullphat.libSnarl;
using fullphat.libSNP31;
//using System.Diagnostics;

namespace Chat_o_licious_
{
    public partial class Form1 : Form
    {
        string _userPic = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string[] destAndPort = textBox2.Text.Split(new char[] { ':' }, 2);
            if (destAndPort.Length != 2)
                return;

            int port = 0;
            if (!int.TryParse(destAndPort[1], out port))
            {
                textBox3.Text = "Bad port number!";
                return;
            }

            if (port < 1 || port > 65535)
            {
                textBox3.Text = "Bad port number!";
                return;
            }

            SNP31Request chatMsg = SNP31.ForwardRequest("Chat-o-licious!", string.Format("New chat from {0}", System.Environment.UserName), textBox1.Text, SNP31.GetFileAsBytes(_userPic));
            SNP31Response reply = SNP31.SendRequest(destAndPort[0], port, chatMsg);

            if (reply.Type == ResponseTypes.Failed)
            {
                textBox3.Text = reply.GetContentValue("error-name");
            }
            else
            {
                textBox3.Text = "";
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            button1.Enabled = (textBox1.Text.Length > 0);
            //Properties.Settings.Default["host"] = textBox1.Text;
            //Properties.Settings.Default.Save();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Chat-o-licious! V1.0\n\nThis is a very simple Chat application that uses Snarl to do all the work.  Enter the IP address and port of a server running Snarl and listening for SNP/3.1, then type some text and click 'send'.\n\nIf you're both running Snarl, you can set up a SNP/3.1 listener as well and have a two-way conversation!", "About Chat-o-licious!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            string path = _userPic;
            if (path == "")
            {
                path = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);
            }
            else
            {
                path = System.IO.Path.GetDirectoryName(path);
            }

            openFileDialog1.FileName = path;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pictureBox1.Image = Image.FromFile(openFileDialog1.FileName);
                _userPic = openFileDialog1.FileName;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _userPic = System.Environment.CurrentDirectory + @"\chat.png";
            try
            {
                pictureBox1.Image = Image.FromFile(_userPic);
            }
            catch { }
        }
    }
}
