using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Automation;
using System.Windows.Forms;
using Eneter.Messaging.EndPoints.TypedMessages;
using Eneter.Messaging.MessagingSystems.MessagingSystemBase;
using Eneter.Messaging.MessagingSystems.TcpMessagingSystem;
using MouseKeyboardActivityMonitor;
using MouseKeyboardActivityMonitor.WinApi;

namespace Dota2Map
{
    public struct KeyQA
    {
        public string Q { get; set; }
        public string A { get; set; }
    }

    public struct AppQA
    {
        public string Q { get; set; }
        public List<string> Answers { get; set; }
    }

    public struct MyResponse
    {
        public int ViberationCount { get; set; }
    }

    public partial class Main : Form
    {
        List<KeyQA> _listQA = new List<KeyQA>();
        private readonly MouseHookListener _mMouseHookManager;
        private static IDuplexTypedMessageSender<AppQA, MyResponse> _mySender;

        public Main()
        {
            try
            {
                InitializeComponent();
                ProcessFile();
                WriteFile();
                MessageBox.Show("Done!");
                return;
                _mMouseHookManager = new MouseHookListener(new GlobalHooker()) {Enabled = true};
                _mMouseHookManager.MouseUp += mMouseHookManager_MouseUp;

                ReadKeyFile(Properties.Settings.Default.KeyPath);

                IDuplexTypedMessagesFactory aSenderFactory = new DuplexTypedMessagesFactory();
                _mySender = aSenderFactory.CreateDuplexTypedMessageSender<AppQA, MyResponse>();

                IMessagingSystemFactory aMessaging = new TcpMessagingSystemFactory();
                IDuplexOutputChannel anOutputChannel =
                    aMessaging.CreateDuplexOutputChannel("tcp://" + Properties.Settings.Default.IpAddress + ":8060/");

                // Attach the input channel and start to listen to messages.
                _mySender.AttachDuplexOutputChannel(anOutputChannel);
            }
            catch (Exception exception)
            {
                WriteAnswer(Properties.Settings.Default.AnswerPath, exception.Message);
            }
        }

        void mMouseHookManager_MouseUp(object sender, MouseEventArgs e)
        {
            try
            {
                string copiedText = CopyText().Trim();
                if (copiedText.Equals(string.Empty) || !Regex.IsMatch(copiedText, "^A.", RegexOptions.Multiline))
                    return;

                WriteAnswer(Properties.Settings.Default.AnswerPath, copiedText);

                string appQaPattern = @"([\s\S]*?)^A\.([\s\S]*?)^B\.([\s\S]*?)";

                appQaPattern += Regex.IsMatch(copiedText, @"^C\.", RegexOptions.Multiline)
                    ? @"^C\.([\s\S]*?)"
                    : string.Empty;
                appQaPattern += Regex.IsMatch(copiedText, @"^D\.", RegexOptions.Multiline)
                    ? @"^D\.([\s\S]*?)"
                    : string.Empty;
                appQaPattern += Regex.IsMatch(copiedText, @"^E\.", RegexOptions.Multiline)
                    ? @"^E\.([\s\S]*?)"
                    : string.Empty;
                appQaPattern += Regex.IsMatch(copiedText, @"^F\.", RegexOptions.Multiline)
                    ? @"^F\.([\s\S]*?)"
                    : string.Empty;
                appQaPattern += Regex.IsMatch(copiedText, @"^G\.", RegexOptions.Multiline)
                    ? @"^G\.([\s\S]*?)"
                    : string.Empty;
                appQaPattern += Regex.IsMatch(copiedText, @"^H\.", RegexOptions.Multiline)
                    ? @"^H\.([\s\S]*?)"
                    : string.Empty;
                appQaPattern += Regex.IsMatch(copiedText, @"^I\.", RegexOptions.Multiline)
                    ? @"^I\.([\s\S]*?)"
                    : string.Empty;

                appQaPattern = appQaPattern.Remove(appQaPattern.Length - 2, 1);
                var appQaMatch = Regex.Match(copiedText, appQaPattern, RegexOptions.Multiline);

                var appQa = new AppQA
                {
                    Q = appQaMatch.Groups[1].Value.Trim(),
                    Answers = new List<string>()
                    {
                        appQaMatch.Groups[2].Value.Trim(),
                        appQaMatch.Groups[3].Value.Trim(),
                        appQaMatch.Groups[4].Value.Trim(),
                        appQaMatch.Groups[5].Value.Trim(),
                        appQaMatch.Groups[6].Value.Trim(),
                        appQaMatch.Groups[7].Value.Trim(),
                        appQaMatch.Groups[8].Value.Trim(),
                        appQaMatch.Groups[9].Value.Trim(),
                        appQaMatch.Groups[10].Value.Trim()
                    }
                };

                var listFoundQa =
                    _listQA.Where(qa => qa.Q.Contains(appQa.Q) && appQa.Answers.Contains(qa.A)).OrderBy(qa => qa.A);
                foreach (KeyQA qa in listFoundQa)
                {
                    string answer = (qa.A + string.Empty).Trim();
                    if (!answer.Equals(string.Empty))
                    {
                        for (int i = 0; i < appQa.Answers.Count; i++)
                        {
                            if (answer.Equals(appQa.Answers[i]))
                            {
                                _mySender.SendRequestMessage(new MyResponse { ViberationCount = i + 1 });
                            }
                        }
                    }
                }

                if (listFoundQa.Any())
                    WriteAnswer(Properties.Settings.Default.AnswerPath, "---");
            }
            catch (Exception ex)
            {
                WriteAnswer(Properties.Settings.Default.AnswerPath, "***" + ex.Message + "***");
            }

        }

        private void WriteAnswer(string path, string data)
        {
            using (var sw = new StreamWriter(path, true))
            {
                sw.WriteLine(data);
            }
        }

        private void ReadKeyFile(string path)
        {
            using (var sr = new StreamReader(path))
            {
                string text = sr.ReadToEnd();

                foreach (Match match in Regex.Matches(text, @"([\s\S]*?)\*([\s\S]*?)\*", RegexOptions.Multiline))
                {
                    _listQA.Add(new KeyQA { Q = match.Groups[1].Value.Trim(), A = match.Groups[2].Value.Trim() });
                }
            }
        }

        private string CopyText()
        {
            AutomationElement element = null;
            try
            {
                element = AutomationElement.FocusedElement;
            }
            catch { }

            if (element != null)
            {
                object pattern;
                if (element.TryGetCurrentPattern(TextPattern.Pattern, out pattern))
                {
                    var tp = (TextPattern)pattern;
                    var sb = new StringBuilder();

                    foreach (var r in tp.GetSelection())
                    {
                        sb.AppendLine(r.GetText(-1));
                    }

                    var selectedText = sb.ToString();

                    return selectedText;
                }
            }

            return string.Empty;
        }

        #region MAKE_KEY_FILE

        private void WriteFile()
        {
            using (var sw = new StreamWriter(@"C:\Users\Duc\Desktop\xx_fixed.txt"))
            {
                foreach (string s in _listQAxxx)
                {
                    sw.WriteLine(s);
                    sw.WriteLine();
                }
            }
        }
        List<string> _listQAxxx = new List<string>();
        void ProcessFile()
        {
            using (var sr = new StreamReader(@"C:\Users\Duc\Desktop\xx.txt"))
            {
                string text = sr.ReadToEnd();

                foreach (Match match in Regex.Matches(text, @"[\s\S]*?Answer ([ABCDEFGH])", RegexOptions.Multiline))
                {
                    //_listQAxxx.Add(Regex.Replace(match.Value.Replace(match.Groups[1].Value + ".	", "*" + match.Groups[1].Value + ".	"), @"ANSWER:	\w", "").Trim());

                    string QA = Regex.Replace(match.Value.Replace(match.Groups[1].Value + ". ", 
                                                                  "*" + match.Groups[1].Value + ". "),
                                              @"Answer \w",
                                              "")
                                     .Trim();

                    QA = Regex.Replace(QA +"\r\n", @"^[ABCDEF]\.[\s\S]*?\r\n", "", RegexOptions.Multiline).Trim() + "*";
                    QA = Regex.Replace(QA + "\r\n", @"^\*[ABCDEF]\.[\s\S]*?", "*", RegexOptions.Multiline).Trim();
                    _listQAxxx.Add(QA);
                }
            }
        }
        #endregion

        private void Main_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (Char) Keys.F12)
            {
                Hide();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Hide();
        }
    }
}
