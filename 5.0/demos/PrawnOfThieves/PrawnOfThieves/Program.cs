using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Diagnostics;
using System.Threading;
using System.IO;
using Melon;
using fullphat.libSnarl;
using fullphat.libSNP31;

namespace PrawnOfThieves
{
    class Program
    {
        static bool _keepRunning = true;
        static ConcurrentDictionary<string, IPEndPoint> _subscribers = new ConcurrentDictionary<string, IPEndPoint>();
        static Random _random = new Random(23);
        static bool _fastMode = false;
        static string _password = "";

        static List<string> _happenings = new List<string>();
        static byte[] _theIcon = SNP31.GetFileAsBytes(Directory.GetCurrentDirectory() + @"\cheese.png");

        static void Main(string[] args)
        {
            int port = 7070;

            if (args.Length > 0)
            {
                int i = Goodies.StringToNumber(args[0], 0);
                if (i > 0 && i < 65536)
                    port = i;

                if (args.Length > 1)
                    _password = args[1];

                //switch (args[0])
                //{
                //    case "--fast":
                //        _fastMode = true;
                //        break;
                //}
            }

            getHappenings();

            TcpListener server = null;
            try
            {
                //IPAddress localAddr = IPAddress.Parse("127.0.0.1");
                server = new TcpListener(IPAddress.Any, port);
                server.Start();

                Console.CancelKeyPress += new ConsoleCancelEventHandler(Console_CancelKeyPress);

                // intro text

                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("\nPrawn of Thieves V1.2 Copyright (C) 1987 full phat scrolls");
                if (_fastMode)
                    Console.WriteLine("<fast mode enabled>");

                if (_password != "")
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("<password protection enabled>");
                }

                Console.ResetColor();
                Console.WriteLine(string.Format("Listening for SNP 3.1 subscriptions on port {0}...\n", port.ToString()));

                // create the sender thread...

                Thread sender = new Thread(senderThread);
                sender.IsBackground = true;
                sender.Start();

                // enter the listening loop...

                while (_keepRunning)
                {
                    TcpClient client = server.AcceptTcpClient();
                    IPEndPoint ip = (IPEndPoint)client.Client.RemoteEndPoint;
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(string.Format("Client {0}:{1} connected...", ip.Address.ToString(), ip.Port.ToString()));
                    Console.ResetColor();

                    NetworkStream stream = client.GetStream();
                    byte[] rx = new byte[client.ReceiveBufferSize];
                    int cb = stream.Read(rx, 0, (int)client.ReceiveBufferSize);
                    string request = Encoding.UTF8.GetString(rx, 0, cb);

                    // process the request here
                    SNP31Request sr = SNP31.DecodeRequest(request);
                    string uid = sr.GetContentValue("uid");

                    // must be valid and can't be a forward-to subscription...
                    if (sr.Action == Actions.Subscribe && sr.GetContentValue("forward-to") == "")
                    {
                        doSubscription(sr, stream, ip);
                    }
                    else if (sr.Action == Actions.Unsubscribe)
                    {
                        if (uid == "")
                        {
                            Console.ForegroundColor = ConsoleColor.Magenta;
                            Console.WriteLine(string.Format("Unsubscribe from {0} missing uid!", ip.Address.ToString()));
                            Console.ResetColor();
                        }
                        else
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.WriteLine(string.Format("Unsubscribe from {0} with uid {1}", ip.Address.ToString(), uid));
                            IPEndPoint ipx;
                            Console.WriteLine("Remove: " + _subscribers.TryRemove(uid, out ipx).ToString());
                            Console.ResetColor();
                        }
                    }
                    else
                    {
                        // if not, send error reply
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Invalid request: must be a SNP 3.1 {subscribe} request with a valid reply-port.");
                        Console.WriteLine(">" + SNP31.SwizleCRLFs(request));
                        Console.ResetColor();

                        // send the error back...
                        send(stream, SNP31.FailedResponse(StatusCodes.BadPacket, "Can only handle reply-port subscriptions").ToString());
                    }
                    client.Close();
                }
            }
            catch (SocketException e)
            {
                Console.WriteLine("SocketException: {0}", e);
            }
            server.Stop();
        }

        static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Got SIGINT: ending...");
            Console.ResetColor();
            _keepRunning = false;

            Console.ForegroundColor = ConsoleColor.White;
            foreach (KeyValuePair<string,IPEndPoint> kvp in _subscribers)
            {
                Console.WriteLine("Sending GOODBYE to " + kvp.Value.Address.ToString() + ":" + kvp.Value.Port.ToString());
                Network.SendData(kvp.Value, SNP31.GoodbyeResponse(kvp.Key).ToString());
            }
            Console.ResetColor();
        }

        static void doSubscription(SNP31Request request, NetworkStream stream, IPEndPoint ip)
        {
            if (_password != "")
            {
                // password required
                AuthenticationResults ar = Snarl.Authenticate(_password, request.HashTypeKeyAndSalt);
                if (ar != AuthenticationResults.Success)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    //Console.BackgroundColor = ConsoleColor.Yellow;
                    Console.WriteLine(string.Format("Authentication required: {0}", ar.ToString()));
                    Console.ResetColor();
                    // reply with error
                    send(stream, SNP31.FailedResponse(StatusCodes.AuthenticationFailure, ar.ToString()).ToString());
                    return;
                }
            }

            // get args...
            int replyPort = Goodies.StringToNumber(request.GetContentValue("reply-port"), 0);
            string uid = request.GetContentValue("uid");

            // success...
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine(string.Format("Got subscribe request (reply-to {0})", replyPort.ToString()));
            Console.WriteLine(string.Format("UID: {0}", uid));
            Console.ResetColor();

            // create an IPEndPoint using the sender's IP address and the reply-port provided...
            IPEndPoint sub = new IPEndPoint(ip.Address, replyPort);

            // the IPEndPoint is the subscriber...
            _subscribers.TryAdd(uid, sub);

            // reply with success
            send(stream, SNP31.SuccessResponse(null).ToString());

            // send initial message...
            sendToSubscriber(uid, sub, "You are standing in a meadow.  There is a bull munching on dandelions in the south-east corner.");

            // done!
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine(string.Format("Sent intro message to {0}:{1}...", sub.Address.ToString(), sub.Port.ToString()));
            Console.ResetColor();
        }

        static void send(NetworkStream stream, string data)
        {
            byte[] tx = Encoding.UTF8.GetBytes(data);
            stream.Write(tx, 0, (int)tx.Length);
        }

        static void senderThread()
        {
            while (true)
            {
                int i = _random.Next(4, 20);
                if (!_fastMode)
                    i = i * 60;

                Console.ForegroundColor = ConsoleColor.DarkYellow;
                Console.WriteLine("Sleeping for " + i.ToString());
                Console.ResetColor();

                Thread.Sleep(i * 1000);

                string s = _happenings[_random.Next(_happenings.Count)];
                Console.WriteLine(string.Format("Sending {0} to subscribers...", Formatting.Quotify(s)));

                foreach (KeyValuePair<string,IPEndPoint> kvp in _subscribers)
                    sendToSubscriber(kvp.Key, kvp.Value, s);
            }
        }

        static void sendToSubscriber(string uid, IPEndPoint subscriber, string notificationText)
        {
            SNP31Request forwardMessage = SNP31.ForwardRequest("Prawn of Thieves", "Grassy Meadow", notificationText, _theIcon);
            ForwardPacket fp = new ForwardPacket(subscriber, forwardMessage, uid);

            Thread t = new Thread(forwarderThread);
            t.IsBackground = true;
            t.Start(fp);
        }

        static void forwarderThread(object forwardPacket)
        {
            ForwardPacket fp = forwardPacket as ForwardPacket;
            string reply = Network.SendDataAndGetReply(fp.Subscriber, fp.ForwardMessage.ToString());
            if (reply == "")
            {
                // didn't reply for some reason so kick them...
                //_subscribers.Remove(fp.Subscriber);
                IPEndPoint ipe = null;
                _subscribers.TryRemove(fp.SubscriberUid, out ipe);

                Console.ForegroundColor = ConsoleColor.Magenta;
                Console.WriteLine(fp.Subscriber.Address.ToString() + ":" + fp.Subscriber.Port.ToString() + ": subscriber kicked for no reply");
                Console.ResetColor();
            }
            else
            {
                SNP31Response response = SNP31.DecodeResponse(reply);
                if (response.Type == ResponseTypes.Failed)
                {
                    Console.ForegroundColor = ConsoleColor.Magenta;
                    Console.WriteLine(fp.Subscriber.Address.ToString() + ":" + fp.Subscriber.Port.ToString() + ": failed:" + response.GetContentValue("error-name"));
                    //Console.WriteLine(replyMessage.Type.ToString());
                    Console.WriteLine(SNP31.SwizleCRLFs(reply));
                    Console.ResetColor();
                }

            }

        }

        static void getHappenings()
        {
            _happenings.Add("The clouds break and the sun comes out.");
            _happenings.Add("The bull stops munching dandelions and glares at you for what seems like an eternity - but is really only 10 seconds - and then returns to his dandelions.");
            _happenings.Add("The sun disappears behind the thickening clouds.");
            _happenings.Add("The wind picks up, rippling the grass.");
            _happenings.Add("Some ramblers stride past hotly debating which way North is and the merits of Sat Nav.");
            _happenings.Add("The bull wanders to another part of the meadow.");
            _happenings.Add("A few spots of rain start to fall, causing the bull momentary confusion.");
        }

    }

    class ForwardPacket
    {
        public IPEndPoint Subscriber;
        public string SubscriberUid;
        public SNP31Request ForwardMessage;

        public ForwardPacket(IPEndPoint subscriber, SNP31Request message, string uid)
        {
            Subscriber = subscriber;
            SubscriberUid = uid;
            ForwardMessage = message;
        }
    }

}