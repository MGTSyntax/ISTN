using MySql.Data.MySqlClient;
using System;
using System.Configuration;
using System.Net.Mail;
using System.Net;
using System.ServiceProcess;
using System.Timers;
using System.IO;
using Serilog;
using System.Diagnostics.Contracts;
using Google.Protobuf.WellKnownTypes;
using System.Threading.Tasks;
using System.Net.Mime;
using static System.Net.Mime.MediaTypeNames;

namespace ISTNotification
{
    public partial class ISTNotification : ServiceBase
    {
        MySqlConnection mibCon;
        MySqlCommand cmd;
        private Timer timer;
        private bool isProcessing = false;
        private static string serverValue = ConfigurationManager.AppSettings["ServerValue"];
        private static string userValue = ConfigurationManager.AppSettings["UserValue"];
        private static string passwordValue = ConfigurationManager.AppSettings["PasswordValue"];
        private static string portValue = ConfigurationManager.AppSettings["PortValue"];
        private readonly string[] mibValues = new string[] { ConfigurationManager.AppSettings["MIB1Value"], ConfigurationManager.AppSettings["MIB2Value"], ConfigurationManager.AppSettings["MIB3Value"], ConfigurationManager.AppSettings["MIB4Value"] };
        private readonly string senderEmail = ConfigurationManager.AppSettings["SenderValue"];
        private readonly string senderPassword = ConfigurationManager.AppSettings["SenderPasswordValue"];
        private readonly string smtpServer = ConfigurationManager.AppSettings["SMTPServerValue"];
        private readonly string daysValue = ConfigurationManager.AppSettings["DaysValue"];
        private readonly string ccValue = ConfigurationManager.AppSettings["CCValue"];
        private readonly string bccValue = ConfigurationManager.AppSettings["BCCValue"];
        private readonly string photoFileName = ConfigurationManager.AppSettings["PhotoFileName"];
        private readonly string photoFileType = ConfigurationManager.AppSettings["PhotoFileType"];

        public ISTNotification()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            StartTimer();
        }

        private void StartTimer()
        {
            // Log the current time
            WriteLogs("Timer started at: " + DateTime.Now.ToString());

            // Calculate the time until 12:01 AM tomorrow
            DateTime now = DateTime.Now;
            DateTime targetTime = now.Date.AddDays(1).AddMinutes(1); // 12:01 AM tomorrow
            //DateTime targetTime = now.Date.AddHours(17).AddMinutes(33);
            if (targetTime <= now)
            {
                // If the specified time has already passed today, schedule it for tomorrow
                targetTime = targetTime.AddDays(1);
            }

            // Calculate the interval in milliseconds
            double intervalMilliseconds = (targetTime - now).TotalMilliseconds;

            // Dispose of the old timer if it exists
            timer?.Dispose();

            // Create a new timer
            timer = new Timer(intervalMilliseconds);
            timer.Elapsed += TimerElapsed;
            timer.Start();
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            // Check if processing is already in progress
            if (isProcessing)
            {
                // Processing is already in progress, so skip this timer event
                return;
            }

            // Set the flag to indicate that processing is now in progress
            isProcessing = true;

            try
            {
                // Connect to the MySQL database
                mibCon = new MySqlConnection($"Server={serverValue};User={userValue};Password={passwordValue};Port={portValue};");

                foreach (string mibValue in mibValues)
                {
                    if (!string.IsNullOrWhiteSpace(mibValue))
                    {
                        ProcessISTN(mibValue);
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogs($"Error: {ex.Message}");
            }
            finally
            {
                // Reset the flag when processing is complete
                isProcessing = false;
                mibCon?.Close();
            }

            // Log the current time
            WriteLogs("Timer elapsed at: " + DateTime.Now.ToString());

            // Restart the timer for the next execution
            StartTimer();
        }

        private void ProcessISTN(string mibName)
        {
            try
            {
                mibCon.Open();

                // Construct and execute your SQL query for the specific MIB
                string query = "SELECT " +
                    "a.ji_empNo as empno, " +
                    "CONCAT(a.ji_lname,CASE WHEN TRIM(a.ji_extname) <> '' THEN CONCAT(' ', a.ji_extname) ELSE '' END, ', ',a.ji_fname,' ',LEFT(a.ji_mname, 1),'.') as empname, " +
                    "c.email_add as empeadd " +
                    $"FROM {mibName}.trans_basicinfo a " +
                    $"INNER JOIN {mibName}.trans_jobinfo b ON a.ji_empNo = b.ji_empNo " +
                    "AND b.ji_active = 1 AND b.ji_dateReg <> '' AND b.ji_jobStat <> 'Processing for Clearance' " +
                    $"LEFT JOIN {mibName}.trans_emailadd c ON a.ji_empNo = c.ji_empNo " +
                    $"LEFT JOIN {mibName}.trans_persinfo d ON a.ji_empNo = d.ji_empNo " +
                    "WHERE " +
                    $"DATE_FORMAT(DATE_SUB(STR_TO_DATE(d.pi_dbirth, '%m/%d/%Y'), INTERVAL {daysValue} DAY),'%m-%d') = DATE_FORMAT(CURDATE(), '%m-%d')";

                cmd = new MySqlCommand(query, mibCon);

                using (MySqlDataReader dr = cmd.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string empno = dr["empno"].ToString();
                        string empname = dr["empname"].ToString();
                        string empeadd = dr["empeadd"].ToString();

                        Task.Run(() => SendEmailWithEmbeddedPhoto(empeadd, empno, empname, mibName));
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogs($"{mibName}: {ex.Message}");
            }
            finally
            {
                mibCon.Close();
            }
        }

        private void SendEmail(string recipientEmailVal, string empno, string empname, string mibName)
        {
            try
            {
                // Create a new MailMessage
                MailMessage mail = new MailMessage(senderEmail, recipientEmailVal);

                mail.Bcc.Add("it@mibsec.com.ph");

                // Set the subject and body of the email
                mail.Subject = "Exclusive In-Service Training Opportunity – Act Now!";
                mail.Body = $"Hi {empname},<br><br>" +
                    "Exciting news! You have a golden opportunity to take in-service training, and you've got exactly 60 days to make the most of it.<br><br>" +
                    "At MIB, we value your growth and want to help you succeed. That's why we're offering this special in-service training just for you.<br><br>" +
                    "🚀 Why You Should Jump In:<br>" +
                    "<ul>" +
                    "<li>Elevate Your Skills: Learn new things that can make you even better at your job.</li>" +
                    "<li>Career Boost: This training can help you advance in your career.</li>" +
                    "<li>Quick and Easy: It's straightforward and won't take too much of your time.</li><br>" +
                    "🎓 Plus, the best part is, when you finish a training course within this time frame, you'll earn a certificate to show off your achievement!<br><br>" +
                    "If you ever need help or have questions, reach out to us:<br>" +
                    "at mntacata@mibsec.com.ph - Marchael N. Tacata<br>" +
                    "at gpzulueta@mibsec.com.ph - Gener P. Zulueta<br>" +
                    "at aehenson@mibsec.com.ph - Annie E. Henson<br><br>" +
                    "Don't miss out – seize the chance to grow and succeed in the next 60 days!<br><br>" +
                    "Cheers to your success 🌟,<br><br>" +
                    "M I B";
                mail.IsBodyHtml = true;

                // Create a new SmtpClient
                SmtpClient smtpClient = new SmtpClient(smtpServer); // Replace with your SMTP server

                // Set the SMTP port and enable SSL if needed
                smtpClient.Port = 587; // For example, use port 587 for Gmail
                smtpClient.EnableSsl = true;

                // Set the sender's credentials
                smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);

                // Send the email
                smtpClient.Send(mail);

                WriteLogs($"Notification for {empno} - {empname} in {mibName} was sent.");
            }
            catch (Exception ex)
            {
                WriteLogs($"SendEmail - {ex.Message}");
            }
        }

        private void SendEmailWithEmbeddedPhoto(string recipientEmailVal, string empno, string empname, string mibName)
        {
            try
            {
                // Create a new MailMessage
                MailMessage mail = new MailMessage();

                // Set the sender
                mail.From = new MailAddress(senderEmail);

                // Set the recipient
                mail.To.Add(new MailAddress(recipientEmailVal));

                // Add CC addresses
                string[] splitCC = ccValue.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string cc in splitCC)
                {
                    mail.CC.Add(new MailAddress(cc));
                }

                // Add BCC address
                mail.Bcc.Add(new MailAddress(bccValue));

                // Set the subject and body of the email
                mail.Subject = "Exclusive In-Service Training Opportunity – Act Now!";
                mail.IsBodyHtml = true;

                // Create HTML content with image
                string htmlBody = $@"Hi {empname}!<br><br>
                             Exciting news! You have a golden opportunity to take in-service training, 
                             and you'll be delighted to know that the training kicks off {daysValue} days before your birthday.<br><br>
                             <img src=""cid:istnads"" alt=""Advertisement Image""><br><br>
                             Don't miss out – seize the chance to grow and succeed starting {daysValue} days before your birthday!<br><br>
                             Cheers to your success 🌟,<br><br>
                             M I B";

                // Create an alternative view with HTML content
                AlternateView htmlView = AlternateView.CreateAlternateViewFromString(htmlBody, null, "text/html");

                string imgPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "istnads.jpg");

                // Load the image from the URL
                LinkedResource imageResource = new LinkedResource(imgPath, "image/jpeg");
                imageResource.ContentId = "istnads"; // Set a unique identifier
                htmlView.LinkedResources.Add(imageResource);

                // Add the alternative view to the email
                mail.AlternateViews.Add(htmlView);

                // Create a new SmtpClient
                using (SmtpClient smtpClient = new SmtpClient(smtpServer))
                {
                    // Set the SMTP port and enable SSL if needed
                    smtpClient.Port = 587; // For example, use port 587 for Gmail
                    smtpClient.EnableSsl = true;

                    // Set the sender's credentials
                    smtpClient.Credentials = new NetworkCredential(senderEmail, senderPassword);

                    // Send the email
                    smtpClient.Send(mail);
                }

                WriteLogs($"Notification for {empno} - {empname} in {mibName} was sent.");
            }
            catch (Exception ex)
            {
                WriteLogs($"SendEmailWithEmbeddedPhoto - {ex.Message}");
            }
        }

        private void WriteLogs(string message)
        {
            string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");
            Directory.CreateDirectory(logPath);

            string logFilePath = Path.Combine(logPath, $"ServiceLog_{DateTime.Now.ToShortDateString().Replace('/', '_')}.txt");

            using (StreamWriter sw = File.AppendText(logFilePath))
            {
                sw.WriteLine(message);
            }
        }

        protected override void OnStop()
        {
            // Stop the timer and clean up resources
            timer.Stop();
            timer.Dispose();
            mibCon?.Close();
        }
    }
}
