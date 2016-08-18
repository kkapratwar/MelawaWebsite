using System;
using System.Net;
using System.Net.Mail;

namespace MelavaWebsite.Common
{
    public class Helper
    {
        public bool SendEmail(string emailId)
        {
            MailMessage msg = new MailMessage();
            try
            {
                msg.From = new MailAddress("durgakalyankinwat@gmail.com");
                msg.To.Add("durgakalyankinwat@gmail.com");
                msg.Subject = "Hello world! " + DateTime.Now.ToString();
                msg.Body = "hi to you ... :)";
                SmtpClient client = new SmtpClient();
                client.UseDefaultCredentials = true;
                client.Host = "smtp.gmail.com";
                client.Port = 587;
                client.EnableSsl = true;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Credentials = new NetworkCredential("durgakalyankinwat@gmail.com", "first#1234");
                client.Timeout = 20000;
                client.Send(msg);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                msg.Dispose();
            }
        }
    }
}