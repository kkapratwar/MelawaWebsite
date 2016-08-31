using System;
using System.Configuration;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;
using MelavaWebsite.Models;

namespace MelavaWebsite.Common
{
    public class Helper
    {
        public bool SendEmail(PersonDetails personDetails)
        {
            MailMessage msg = new MailMessage();
            try
            {
                string fromEmail = ConfigurationManager.AppSettings["EmailAddress"];
                string password = ConfigurationManager.AppSettings["EmailPassword"];
                msg.From = new MailAddress(fromEmail);
                msg.To.Add(string.IsNullOrEmpty(personDetails.Email)
                    ? fromEmail
                    : personDetails.Email);
                msg.Bcc.Add(fromEmail);
                msg.Subject = "Do Not Reply:Melava Registration.";
                StringBuilder body =
                    new StringBuilder(File.ReadAllText(AppDomain.CurrentDomain.BaseDirectory + ConfigurationManager.AppSettings["EmailTemplate"]));
                body.Replace("{Candidate Id}", personDetails.Id.ToString());
                body.Replace("{Full Name}", personDetails.Name);
                body.Replace("{Birth Name}", personDetails.BirthName);
                body.Replace("{First Gotra}", personDetails.FirstGotra);
                body.Replace("{Second Gotra}", personDetails.SecondGotra);
                body.Replace("{Gender}", personDetails.Gender);
                body.Replace("{Age}", personDetails.Age.ToString());
                body.Replace("{Height}", Convert.ToString(personDetails.Height));
                body.Replace("{Education}", personDetails.Education);
                body.Replace("{Occupation}", personDetails.Occupation);
                body.Replace("{Address}", personDetails.Address);
                body.Replace("{Contact Number}", personDetails.ContactNumber);
                body.Replace("{Anubandh Id}", Convert.ToString(personDetails.AnubandhId));
                body.Replace("{Email-Id}", personDetails.Email);
                msg.Body = body.ToString();
                msg.IsBodyHtml = true;
                SmtpClient client = new SmtpClient
                {
                    UseDefaultCredentials = true,
                    Host = "smtp.gmail.com",
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = new NetworkCredential(fromEmail, password),
                    Timeout = 20000
                };
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