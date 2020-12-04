using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using AmazonSES;
using System.Threading;
using System.Runtime.InteropServices;
using System.Linq;
using System.Net.NetworkInformation;
using System.Net;
using System.Text;
using System.Collections;
using System.Web;
using System.Text.RegularExpressions;
using System.Globalization;

namespace NTPEmailMarketing
{
    public partial class NTPEmailMarketing : Form
    {
        public NTPEmailMarketing()
        {
            InitializeComponent();
            //StartUpManager.AddApplicationToCurrentUserStartup("Gmail Marketing");
        }
        public static string RootFolder = Directory.GetParent(Directory.GetCurrentDirectory()).FullName;
        public static string EmailFolder = Directory.GetParent(RootFolder).FullName + "\\EmailList\\";
        public static string MaleEmailFolder = Directory.GetParent(RootFolder).FullName + "\\EmailList_Male\\";
        public static string FemaleEmailFolder = Directory.GetParent(RootFolder).FullName + "\\EmailList_Female\\";
        public static string PostFolder = Directory.GetParent(RootFolder).FullName + "\\Categories\\";
        public static string groupsEmailPath = Directory.GetParent(RootFolder).FullName + "\\EmailList_Groups\\EmailList.txt";
        public static string logFilePath = Application.StartupPath + "\\logs.txt";
        public static string rejectedFilePath = Application.StartupPath + "\\rejected.txt";
        public static string delayFilePath = Application.StartupPath + "\\delay.txt";
               
        const string ACCESS_KEY = "AKIAJWG23UERWOMK6H7A";//"AKIAJOSY7K6WCYKPW4UA";
        const string SECRET_ACCESS_KEY = "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";
        //;"roKekEtTvFJNzxODEVbMy8AztfYb046cbTlb9gv5";        
        public const string PATTERN_ALL = "*.*";
        public const string PATTERN_WORD = "*.docx";
        public const string PATTERN_HTML = "*.html";
        public static string[] Category = { "SanKhuyenMai", "SanKhuyenMaiMoiNgay", "ThuThuatMayTinh", "ThuGianMoiNgay", "KienThucMoiNgay", "KienThucDanhChoNam", "KienThucDanhChoNu" };//, "KiemTienOnline" };        
        Thread[] listThreads;
        private bool isRunning;

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;

        [DllImport("user32.dll")]
        static extern bool AnimateWindow(System.IntPtr hWnd, int time, AnimateWindowFlags flags);
        [System.Flags]
        enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x00000001,
            AW_HOR_NEGATIVE = 0x00000002,
            AW_VER_POSITIVE = 0x00000004,
            AW_VER_NEGATIVE = 0x00000008,
            AW_CENTER = 0x00000010,
            AW_HIDE = 0x00010000,
            AW_ACTIVATE = 0x00020000,
            AW_SLIDE = 0x00040000,
            AW_BLEND = 0x00080000
        }

        private void WindowsInSystemTray(bool inTray)
        {
            if (inTray)
            {
                this.ShowInTaskbar = false;
                AnimateWindow(this.Handle, 50, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_HIDE);
                myNotifyIcon.Visible = true;
                myNotifyIcon.ShowBalloonTip(500);
            }
            else
            {
                this.ShowInTaskbar = true;
                this.WindowState = FormWindowState.Normal;
                AnimateWindow(this.Handle, 700, AnimateWindowFlags.AW_BLEND | AnimateWindowFlags.AW_ACTIVATE);
                this.Activate();
                myNotifyIcon.Visible = false;
            }
        }

        private void PostCardGroup()
        {
            string[,] results = Handler.GetPromotionNews();
            string to = "dichvunaptienthecao@groups.facebook.com";
            string toMyself = "foryouforme87@facebook.com";
            for (int i = 0; i < results.Length; i++)
            {
                string subject = "[" + results[i, 0] + "] THÔNG TIN KHUYẾN MÃI";
                string message = results[i, 1] + Environment.NewLine;
                message += "Mại dô! Mại dô! :)";
                SendGmail(subject, to, message);
                //Thread.Sleep(10 * 1000);
                //SendGmail(subject, toMyself, message);
                Thread.Sleep(10 * 60 * 1000);
            }
        }

        private void SendGmail(string subject, string to, string body)
        {
            MailAccount oMailAccount = new MailAccount();
            oMailAccount.DisplayName = "Dịch Vụ Nạp Tiền - Thẻ Cào";
            oMailAccount.Username = "tommynguyen24612@gmail.com";//"phat72@gmail.com";
            oMailAccount.Password = "phatIT0687!@#123";//"pA55WordFormE!@#";            
                     
            // Send email
            if (!EmailMarketing.IsExistTitle(body, DateTime.Now.ToString("dd/MM/yyyy")))
            {
                Handler.SendEmail(oMailAccount, subject, to, body, true);

                // Log start sending email
                StringBuilder sbOperation = new StringBuilder();
                sbOperation.AppendLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                sbOperation.AppendLine("Start sending email!");
                sbOperation.AppendLine("Subject: " + subject);
                sbOperation.AppendLine("From: " + oMailAccount.Username);
                sbOperation.AppendLine("To: " + to);
                Handler.Log(logFilePath, sbOperation.ToString());   
            }                             
        }

        private void SendGmail(object Index)
        {
            string subject = "testing", to = "tommynguyen987@gmail.com", body = "this is testing";
            string[] emailFileList = null;
            string[] postFileList = null;
            MailAccount oMailAccount = new MailAccount();
            int index = (int)Index;

            Handler.ConvertFile(PostFolder + Category[index]);
            postFileList = Handler.GetFiles(PostFolder + Category[index], PATTERN_HTML);
            delayFilePath = delayFilePath.Replace(".txt", index + ".txt");
            switch (index)
            {
                case 0:                    
                    oMailAccount.DisplayName = "Săn Khuyến Mãi";
                    oMailAccount.Username = "tommynguyen24612@gmail.com";//"phat72@gmail.com";
                    oMailAccount.Password = "phatIT0687!@#123";//"pA55WordFormE!@#";
                    break;
                case 1:
                    oMailAccount.DisplayName = "Săn Khuyến Mãi";
                    oMailAccount.Username = "tinkhuyenmaimoingay@gmail.com";
                    emailFileList = Handler.GetFiles(EmailFolder, PATTERN_ALL);
                    break;
                case 2:                    
                    oMailAccount.DisplayName = "Thủ thuật Máy Tính";
                    oMailAccount.Username = "thuthuatmaytinhmoingay@gmail.com";
                    emailFileList = Handler.GetFiles(EmailFolder, PATTERN_ALL);
                    break;
                case 3:
                    oMailAccount.DisplayName = "Thư Giãn Mỗi Ngày";
                    oMailAccount.Username = "tinthugianmoingay@gmail.com";
                    emailFileList = Handler.GetFiles(EmailFolder, PATTERN_ALL);
                    break;
                case 4:
                    emailFileList = Handler.GetFiles(EmailFolder, PATTERN_ALL);
                    break;
                case 5:
                    emailFileList = Handler.GetFiles(MaleEmailFolder, PATTERN_ALL);
                    break;
                case 6:
                    emailFileList = Handler.GetFiles(FemaleEmailFolder, PATTERN_ALL);
                    break;
            }

            if (postFileList.Length == 0)
            {
                return;
            }
                
            if (index == 0)
            {
                foreach (var postFile in postFileList)
                {
                    FileInfo file = new FileInfo(postFile);
                    subject = file.Name.Replace(file.Extension, "");

                    string time = subject.Split('-')[1];
                    string temp = time.Replace(time.Substring(time.IndexOf(']')), "").Trim();
                    DateTime datetime = DateTime.Parse(temp);
                    if (datetime.Year == DateTime.Now.Year)
                    {
                        datetime = datetime.AddDays(1);
                    }

                    if (DateTime.Now > datetime)
                    {
                        file.Delete();
                        string tempWord = postFile.Replace(".html", ".docx");                                            
                        file = new FileInfo(tempWord);                        
                        file.Delete();
                        string tempFolder = postFile.Replace(".html", "_files");
                        DirectoryInfo dir = new DirectoryInfo(tempFolder);
                        FileInfo []files = dir.GetFiles();
                        foreach (var item in files)
                        {
                            item.Delete();
                        }
                        dir.Delete();
                    }
                    else
                    {
                        subject =  subject.Replace("Săn Khuyến Mãi - ", "");
                        temp = subject.Substring(subject.IndexOf("]") + 1).Trim();
                        time = subject.Replace("[ ", "").Replace("] " + temp, "").Trim();
                        time = DateTime.Parse(time).ToString("dd/MM/yyyy");
                        subject = " [ " + time + " ] " + temp;
                        body = HttpUtility.HtmlDecode(File.ReadAllText(postFile, Encoding.UTF8));
                        string[] emailsList = File.ReadAllLines(groupsEmailPath)
                                                .Where(arg => !string.IsNullOrWhiteSpace(arg))
                                                .Distinct()
                                                .ToArray();

                        if (EmailMarketing.IsExistName(subject))
                        {
                            continue;
                        }

                        bool isInternetConnected = true;
                        do
                        {
                            isInternetConnected = Handler.IsAvailableNetworkActive();
                        } while (!isInternetConnected);

                        int delay = Handler.GetDelay(delayFilePath);
                        if (delay > 0)
                        {
                            Thread.Sleep(delay * 60 * 60 * 1000);
                        }

                        // Log start sending email
                        StringBuilder sbOperation = new StringBuilder();
                        sbOperation.AppendLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                        sbOperation.AppendLine("Start sending email!");
                        sbOperation.AppendLine("Subject: " + subject);
                        sbOperation.AppendLine("From email file: " + groupsEmailPath);
                        Handler.Log(logFilePath, sbOperation.ToString());  

                        // Send email
                        Handler.SendEmail(oMailAccount, subject, emailsList, body);                                             
                    }
                }
            }
            else
            {
                if (emailFileList.Length == 0)
                {
                    return;
                }

                foreach (var postFile in postFileList)
                {
                    foreach (var emailFile in emailFileList)
                    {
                        FileInfo file = new FileInfo(postFile);
                        subject = file.Name.Replace(file.Extension, "");
                        body = HttpUtility.HtmlDecode(File.ReadAllText(postFile, Encoding.UTF8));
                        string[] emailsList = File.ReadAllLines(emailFile)
                                                .Where(arg => !string.IsNullOrWhiteSpace(arg))
                                                .Distinct()
                                                .ToArray();

                        if (!string.IsNullOrEmpty(body))
                        {
                            for (int i = 0; i < emailsList.Length; i++)
                            {
                                to = Handler.GetValidEmail(emailsList[i]);
                                if (Handler.isValidEmail(to))
                                {
                                    int delay = Handler.GetDelay(delayFilePath);
                                    if (delay > 0)
                                    {
                                        Thread.Sleep(delay * 60 * 60 * 1000);
                                    }
                                    Handler.InsertDelay(delayFilePath, 0);

                                    if (EmailMarketing.IsExistName(to, subject) || EmailMarketing.IsRejected(to))
                                    {
                                        continue;
                                    }

                                    bool isInternetConnected = true;
                                    do
                                    {
                                        isInternetConnected = Handler.IsAvailableNetworkActive();
                                    } while (!isInternetConnected);

                                    // Log start sending email
                                    StringBuilder sbOperation = new StringBuilder();
                                    sbOperation.AppendLine(DateTime.Now.ToString("dd/MM/yyyy hh:mm:ss tt"));
                                    sbOperation.AppendLine("Start sending email!");
                                    sbOperation.AppendLine("To: " + to);
                                    sbOperation.AppendLine("Subject: " + subject);
                                    sbOperation.AppendLine("From email file: " + emailFile);
                                    Handler.Log(logFilePath, sbOperation.ToString());
                                    // Send email
                                    Handler.SendEmail(oMailAccount, subject, to, body, false);
                                }
                            }
                        }
                    }
                }
            }            
                              
            MessageBox.Show("Sending email finished!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void SendAmazonSESEmail()
        {
            try
            {
                string to = "tommynguyen987@gmail.com";
                string from = "chiasekienthucmoingay@gmail.com";
                string replyto = "chiasekienthucmoingay@gmail.com";
                string subject = "Testing Amazon SES";
                string body = "this is testing amazon ses";
                AmazonSESWrapper amazonses = new AmazonSESWrapper(ACCESS_KEY, SECRET_ACCESS_KEY);
                //amazonses.VerifyEmailAddress(from);
                //amazonses.VerifyEmailAddress(to);
                AmazonSentEmailResult result = amazonses.SendEmail(to, from, replyto, subject, body);
                //MessageBox.Show("MessageId: " + result.MessageId + Environment.NewLine + "HasError: " + result.HasError + Environment.NewLine + "ErrorException: " + result.ErrorException);
                foreach (DataGridViewRow row in grvEmails.Rows)
                {
                    //AmazonSESWrapper amazonses = new AmazonSESWrapper("AKIAJWG23UERWOMK6H7A", "6k5d2NNTAQ67kJlnp5Zse8xjofMAbg2lT1bW3py4");
                    //row.Cells[2].Value = "Sending...";
                    //AmazonSentEmailResult result = amazonses.SendEmail(row.Cells[1].Value.ToString(), from, replyto, subject, body);
                    //if (!result.HasError)
                    //{
                    //    row.Cells[2].Value = "Finished.";
                    //}
                    //MessageBox.Show("MessageId: " + result.MessageId + Environment.NewLine + "HasError: " + result.HasError + Environment.NewLine + "ErrorException: " + result.ErrorException);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Xảy ra lỗi khi gửi mail! " + Environment.NewLine + ex.Message);
            }
        }

        private void showMainToolStripMenuItem_Click(object sender, EventArgs e)
        {
            WindowsInSystemTray(false);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            myNotifyIcon.Dispose();
            this.Close();
            this.Dispose();            
        }

        private void myNotifyIcon_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            WindowsInSystemTray(false);
        }

        private void NTPEmailMarketing_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                WindowsInSystemTray(true);
            }
        }

        private void NTPEmailMarketing_FormClosing(object sender, FormClosingEventArgs e)
        {
            e.Cancel = true;
            WindowsInSystemTray(true);
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            var result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog1.FileName;
                var lines = File.ReadAllLines(txtFilePath.Text);
                foreach (var line in lines)
                {
                    int index = grvEmails.Rows.Add();
                    grvEmails.Rows[index].Cells[0].Value = (index + 1);
                    grvEmails.Rows[index].Cells[1].Value = line;
                }                
            }
        }

        private void btnSend_Click(object sender, EventArgs e)
        {
            btnBrowse.Enabled = false;
            btnDelete.Enabled = false;
            backgroundWorker1.RunWorkerAsync();                        
        } 

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {            
            if(e.Cancel)
            {                
                backgroundWorker1.CancelAsync();
            }
            else
            {                
                listThreads = new Thread[6];
                for (int i = 0; i < 6; i++)
                {
                    int input = i;
                    listThreads[i] = new Thread(new ParameterizedThreadStart(SendGmail));
                    listThreads[i].Start(input);
                }
                //SendGmail(0);
            }            
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            myStatus.Text = "Sending email...";
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {                
                myStatus.Text = "Sending email cancelled";
            }
            else
            {
                myStatus.Text = "Done";
            }
            btnBrowse.Enabled = true;
            btnDelete.Enabled = true;
        }

        private void grvEmails_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            for (int i = 0; i < grvEmails.Rows.Count; i++)
            {
                grvEmails.Rows[i].Cells[0].Value = i + 1;
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            var confirm = MessageBox.Show("Bạn có chắc muốn xóa dòng này không?", "Xác Nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirm == DialogResult.Yes)
            {
                foreach (DataGridViewRow row in grvEmails.Rows)
                {
                    if (row.Selected)
                    {
                        int index = row.Index;
                        grvEmails.Rows.Remove(row);
                    }
                }
                grvEmails.Refresh();
            }
        }

        private void NTPEmailMarketing_Load(object sender, EventArgs e)
        {
            myNotifyIcon.BalloonTipText = "Your application is still working" + System.Environment.NewLine + "Double click into icon to show application.";
            WindowsInSystemTray(true);            
            backgroundWorker1.RunWorkerAsync();
        }

        private void myReminder_Tick(object sender, EventArgs e)
        {
            if (DateTime.Now.Hour >= 8)
            {
                PostCardGroup();
            }
            Handler.UpdateRejected(rejectedFilePath);
        }        
    }
}
