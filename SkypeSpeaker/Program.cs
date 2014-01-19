using System;
using System.Collections.Generic;
using System.Speech.Synthesis;
using System.Text.RegularExpressions;
using System.Timers; // Use UiacomWrapper for better eventing
using System.Windows.Automation;

namespace SkypeSpeaker
{
    class Program
    {
        SpeechSynthesizer _Speech = new SpeechSynthesizer();
        AutomationElement _ChatWindow;
        string _LastMessage = null;
        Dictionary<string, string> _VoiceMap = new Dictionary<string, string>();

        public Program(string[] args)
        {
            for (var i = 0; i < args.Length; i += 2)
            {
                if (i + 1 < args.Length)
                {
                    _VoiceMap.Add(args[i].ToLower(), args[i + 1]);
                    Console.WriteLine(string.Format("{0} verbalized with {1}", args[i], args[i + 1]));
                }
                else
                {
                    Help();
                    return;
                }
            }

            if (args.Length == 0)
            {
                _VoiceMap = null; // All mode
            }

            Console.WriteLine("Attaching to Skype...");
            StartListening();
            Console.WriteLine("Ready for Skype chat events. Set the conversation view type to Compact. Press any key to exit.");
            Console.ReadKey();
        }

        void Help()
        {
            Console.WriteLine("Usage: <self.exe> [UserName1 Voice1] [UserName2 Voice2]");
            Console.WriteLine("Installed Voices:");
            foreach(var v in _Speech.GetInstalledVoices())
            {
                Console.WriteLine(v.VoiceInfo.Name);
            }
        }

        void StartListening()
        {
            var SkyeMainWindowCondition = new AndCondition(
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Window),
                new PropertyCondition(AutomationElement.ClassNameProperty, "tSkMainForm"));

            var SkypeWindow = AutomationElement.RootElement.FindFirst(TreeScope.Children, SkyeMainWindowCondition);
            _ChatWindow = SkypeWindow.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, "Chat Content List"));
            Automation.AddStructureChangedEventHandler(_ChatWindow, TreeScope.Children, (sender, e) =>
            {
                // Without UiaComWrapper, e.InvalidatedChildren is only fired on GridPattern supported elements.
                CheckMessages();
            });

            var t = new Timer();
            t.Elapsed += (_, __) => CheckMessages();
            t.Interval = 5000;
            t.Start();

            CheckMessages(); // Set _LastMessage
        }

        void CheckMessages()
        {
            var Messages = new Stack<string>();

            var lastChild = TreeWalker.RawViewWalker.GetLastChild(_ChatWindow);
            do
            {
                if (lastChild != null)
                {
                    // Handles the case where ' New' is appended to the end
                    var Msg = lastChild.Current.Name;
                    if ((_LastMessage != null && Msg.Contains(_LastMessage)) || 
                        (_LastMessage != null && _LastMessage.Contains(Msg)))
                    {
                        break; // We've seen this message before.
                    }
                    else if (Msg.EndsWith(" is typing")) // Kill typing indicator
                    {
                        continue;
                    }
                    else
                    {
                        if (_LastMessage == null) // Set last marker for startup
                        {
                            _LastMessage = Msg;
                            break;
                        }
                        Messages.Push(Msg);
                    }
                }
            } while ((lastChild = TreeWalker.RawViewWalker.GetPreviousSibling(lastChild)) != null);

            while (Messages.Count > 0) Speak(Messages.Pop());
        }

        private void Speak(string text)
        {
            _LastMessage = text;
            // Remove links.
            text = new Regex(@"https?://([\w\.]*)/?.*?([\s,])", 
                RegexOptions.None).Replace(text, @"$1$2");
            // Remove timestamp
            var match = Regex.Match(text, @"From (?<From>.*?), (?<Message>.*), sent on", RegexOptions.IgnoreCase);
            if (match.Success)
            {
                var from = match.Groups["From"].Value;
                var message = match.Groups["Message"].Value;

                if (_VoiceMap == null)
                {
                    Console.WriteLine("{0}: {1}", from, message);
                    _Speech.SpeakAsync(string.Format("{0} says {1}", from, message));
                }
                else
                {
                    Console.WriteLine("{0}: {1}", from, message);

                    if (_VoiceMap.ContainsKey(from.ToLower()))
                    {
                        var prevVoice = _Speech.Voice.Name;
                        _Speech.SelectVoice(_VoiceMap[from.ToLower()]);
                        _Speech.Speak(message);
                        _Speech.SelectVoice(prevVoice);
                    }
                }
            }
            else
            {
                Console.WriteLine("Non-verbalized message: " + text);
            }
        }

        static void Main(string[] args)
        {
            new Program(args);
        }
    }
}