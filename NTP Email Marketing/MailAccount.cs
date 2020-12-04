using System.Net;
using System.Net.Mail;

namespace NTPEmailMarketing
{
    public class MailAccount
    {
        public string DisplayName = "Kiến Thức Mỗi Ngày";
        public string SMTPHostname = "smtp.gmail.com";
        public string Username = "chiasekienthucmoingay@gmail.com";
        public string Password = "Admin123!@#";
        public int SMTPPort = 587;
        public bool SMTPSSL = true;

        public SmtpClient GetSmtpClient()
        {
            var client = new SmtpClient(SMTPHostname, SMTPPort)
            {
                DeliveryMethod = SmtpDeliveryMethod.Network,
                EnableSsl = SMTPSSL,
                Host = SMTPHostname,
                Credentials = new NetworkCredential(Username, Password)
            };

            return client;
        }

        public MailAccount GetDefaultMailAccount()
        {
            MailAccount oMailAccount = new MailAccount();
            var _with1 = oMailAccount;
            _with1.DisplayName = DisplayName;
            _with1.Username = Username;
            _with1.Password = Password;
            _with1.SMTPHostname = SMTPHostname;
            _with1.SMTPPort = SMTPPort;
            _with1.SMTPSSL = true;
            return oMailAccount;

            //MailAccount oMailAccount = null;
            //oMailAccount = new MailAccount();
            //var _with1 = oMailAccount;
            //_with1.Name = "Default";
            //_with1.SMTPHostname = "email-smtp.us-east-1.amazonaws.com";
            //_with1.Username = "AKIAJ6IKW7YVMLIDTFQQ";
            //_with1.Password = "AjN+lZ0O2QiV5de7yNwQoMk3WhHxs2q9CgG3AHAdujk8";
            //_with1.SMTPPort = 465;
            //_with1.SMTPSSL = true;
            //return oMailAccount;
        }
    }
}
