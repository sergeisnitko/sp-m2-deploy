using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class Mail
    {
        /// <summary>
        /// Method for sending mail messages with smtp client
        /// <param name="Context">SharePoint CSOM client context</param>
        /// <param param name="to">email of the recipient</param>
        /// <param name="from">email of the sender</param>
        /// <param name="host">smtp server address</param>
        /// <param name="title">the title of the message</param>
        /// <param name="message">the body of the message</param>
        /// <param name="EnableSsl">set this param in true, if you need to use SSL in smtp client</param>
        /// </summary> 
        public static void SpfSendMail(this ClientContext Context, string to, string from, string host, string title, string message, bool EnableSsl = false)
        {
            var Site = Context.Site;
            Context.Load(Site);
            Context.ExecuteQuery();

            var NSymb = "";
            if (Site.ServerRelativeUrl == "/")
            {
                NSymb = "/";
            }
            message = message.Replace("href=\"" + Site.ServerRelativeUrl, "href=\"" + Site.Url + NSymb).Replace("&#160;", "&nbsp;");

            title = String.IsNullOrEmpty(title) ? "SPF Message" : title;
            var mail = new MailMessage(from, to);

            var client = new SmtpClient();
            client.Host = host;
            client.EnableSsl = EnableSsl;
            mail.Subject = title;
            mail.Body = message;
            mail.IsBodyHtml = true;

            client.Send(mail);
        }
    }
}
