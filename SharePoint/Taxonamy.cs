using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class Taxonamy
    {
        public static int GetWssId(this ClientContext Context, Guid TermId, Guid TermSetId)
        {
            var RootWeb = Context.Site.RootWeb;
            Context.Load(RootWeb);
            Context.ExecuteQuery();

            var TaxonomyList = RootWeb.GetListByUrl("Lists/TaxonomyHiddenList");
            var Items = TaxonomyList.GetItemsByTwoStrings("IdForTerm", TermId.ToString("d"), "IdForTermSet", TermSetId.ToString("d"));
            if (Items.Count > 0)
            {
                var Item = Items[0];
                return Item.Id;
            }
            return -1;
        }
    }
}
