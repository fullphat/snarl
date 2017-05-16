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

        static void Main(string[] args)
        {
            Console.WriteLine("\nNotification Pump");
            Console.WriteLine("melon.melonlibrary: " + Melon.Melon.Version.ToString());
            Console.WriteLine("snp31.melonlibrary: " + SNP31.Version.ToString());

            IniFile _config = new IniFile(FilesAndFolders.GetCurrentDirectory() + "config.rc");
            List<IniEntry> dests = _config.GetSectionContent("destinations");
            if (dests.Count == 0)
            {
                Print.Error("No destinations configured!");
                return;
            }

            List<IniEntry> quotes = _config.GetSectionContent("quotes");
            if (quotes.Count == 0)
            {
                Print.Error("No quotes defined!");
                return;
            }

            foreach (IniEntry ie in dests)
                _destinations.Add(ie.Name);

            foreach (IniEntry ie in quotes)
                _quotes.Add(ie.Name);

            Console.CancelKeyPress += new ConsoleCancelEventHandler(Console_CancelKeyPress);

            Timer quoteMonkey = new Timer(60 * 60 * 1000);
            quoteMonkey.AutoReset = true;
            quoteMonkey.Elapsed += new ElapsedEventHandler(quoteMonkey_Elapsed);
            quoteMonkey.Start();

            Print.Info(string.Format("Next quote in {0} seconds...", quoteMonkey.Interval));

            while (_keepRunning)
            {
                System.Threading.Thread.Sleep(100);
            }

        }

        static void quoteMonkey_Elapsed(object sender, ElapsedEventArgs e)
        {
            int i = _teaLeaves.Next(_quotes.Count);
            string s = _quotes[i];
            Print.Note(s);
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
            SNP31Request req = SNP31.ForwardRequest("Notification Pump", "Quote of the Day", qadc.Quote, FilesAndFolders.GetCurrentDirectory() + "icon.png", "");
            SNP31Response rep = SNP31.SendRequest(qadc.Host, qadc.Port, req);
            Print.Note(string.Format("{0}:{1}>{2}", qadc.Host, qadc.Port, rep.Type.ToString()));
        }

    }

    class QuoteAndDestCombo
    {
        public string Host = "";
        public int Port = 0;
        public string Quote = "";

        public QuoteAndDestCombo(string dest, string quote)
        {
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
