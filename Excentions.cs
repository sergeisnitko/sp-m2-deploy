using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace SPF
{
    public static class Excentions
    {
        public static void SetWelcomePage(this List list, string pageUrl)
        {
            ClientContext clientContext = (ClientContext)list.Context;
            var rootFolder = list.RootFolder;
            rootFolder.WelcomePage = pageUrl;
            rootFolder.Update();
            clientContext.ExecuteQuery();
        }

    }
}
