using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Timers;
using fullphat.libSNP31;
using Melon.SettingsBundles;
using Melon.Con;
using Melon;

namespace NotificationPump
{
    class Program
    {
        static List<string> _destinations = new List<string>();
        static List<string> _quotes = new List<string>();
        static bool _keepRunning = true;
        static Random _teaLeaves = new Random();
        static byte[] _icon;

        static void Main(string[] args)
        {
            //Print.Info("\nNotification Pump");
            Print.Info("melon.melonlibrary: " + Melon.Melon.Version.ToString());
            Print.Info("snp31.melonlibrary: " + SNP31.Version.ToString());

            // read config file...
            IniFile _config = new IniFile(FilesAndFolders.GetCurrentDirectory() + "config.rc");

            // get [destinations] section...
            List<IniEntry> dests = _config.GetSectionContent("destinations");
            if (dests.Count == 0)
            {
                Print.Error("No destinations configured!");
                return;
            }

            // get [quotes] section...
            List<IniEntry> quotes = _config.GetSectionContent("quotes");
            if (quotes.Count == 0)
            {
                Print.Error("No quotes defined!");
                return;
            }

            // (critical startup complete)

            // cache icon...
            _icon = SNP31.GetFileAsBytes(FilesAndFolders.GetCurrentDirectory() + "icon.png");

            // install CTRL+C handler...
            Console.CancelKeyPress += new ConsoleCancelEventHandler(Console_CancelKeyPress);

            // set timer...
            Timer quoteMonkey = new Timer(60 * 60 * 1000);
            quoteMonkey.AutoReset = true;
            quoteMonkey.Elapsed += new ElapsedEventHandler(quoteMonkey_Elapsed);
            Print.Info(string.Format("Sending a quote every {0} seconds...", quoteMonkey.Interval));

            Print.Info("Will send to the following...");
            foreach (IniEntry ie in dests)
            {
                Print.Info("  " + ie.Name);
                _destinations.Add(ie.Name);
            }

            // get quotes...
            foreach (IniEntry ie in quotes)
            {
                if (!ie.Name.StartsWith("#"))
                    _quotes.Add(ie.Name);
            }

            // send first quote now...
            Print.Note("Running...");
            quoteMonkey_Elapsed(null, null);

            // then start timer...
            quoteMonkey.Start();

            // run...
            while (_keepRunning)
            {
                System.Threading.Thread.Sleep(100);
            }

        }

        static void quoteMonkey_Elapsed(object sender, ElapsedEventArgs e)
        {
            int i = _teaLeaves.Next(_quotes.Count);
            string s = _quotes[i];
            Print.Info(s);
            send(s);
        }

        static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            _keepRunning = false;
        }

        static void send(string text)
        {
            foreach (string dest in _destinations)
            {
                QuoteAndDestCombo qadc = new QuoteAndDestCombo(dest, text);
                System.Threading.Thread t = new System.Threading.Thread(senderThread);
                t.IsBackground = true;
                t.Start(qadc);
            }
        }

        static void senderThread(object quoteAndDestCombo)
        {
            // extract the object...
            QuoteAndDestCombo qadc = quoteAndDestCombo as QuoteAndDestCombo;
            SNP31Request req = SNP31.ForwardRequest("Notification Pump", "Quote of the Day", qadc.Quote, _icon);
            SNP31Response rep = SNP31.SendRequest(qadc.Host, qadc.Port, req);

            if (rep.Type == ResponseTypes.Success)
            {
                Print.Ok(string.Format("{0}:{1}>{2}", qadc.Host, qadc.Port, rep.Type.ToString()));
            }
            else
            {
                Print.Warn(string.Format("{0}:{1}>{2}", qadc.Host, qadc.Port, rep.Type.ToString()));
            }
        }

    }

    class QuoteAndDestCombo
    {
        public string Host = "";
        public int Port = 0;
        public string Quote = "";

        public QuoteAndDestCombo(string dest, string quote)
        {
            //Print.Error(dest);
            KeyValuePair<string, string> kvp = Formatting.SplitPair(dest, ':');
            if (!string.IsNullOrEmpty(kvp.Key))
            {
                Host = kvp.Key;
                Port = Goodies.StringToNumber(kvp.Value, 0);
                Quote = quote;
            }
        }

    }

}
