using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class SpfWeb
    {
        public static List GetListByUrl(this Web web, string listUrl)
        {
            var clientContext = (ClientContext)web.Context;
            if (!web.IsPropertyAvailable("ServerRelativeUrl"))
            {
                clientContext.Load(web, currentWeb => currentWeb.ServerRelativeUrl);
                clientContext.ExecuteQuery();
            }

            string finalUrl = web.ServerRelativeUrl + "/" + listUrl;
            finalUrl = finalUrl.Replace("//", "/");
            List list = web.GetList(finalUrl);
            return list;
        }
        public static List LoadListByUrl(this Web web, string listUrl)
        {
            var ctx = (ClientContext)web.Context;
            var listFolder = web.GetFolderByServerRelativeUrl(listUrl);
            ctx.Load(listFolder.Properties);
            ctx.ExecuteQuery();
            var listId = new Guid(listFolder.Properties["vti_listname"].ToString());
            var list = web.Lists.GetById(listId);
            ctx.Load(list);
            ctx.ExecuteQuery();
            return list;
        }
    }
}
