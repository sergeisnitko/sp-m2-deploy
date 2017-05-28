using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class Query
    {
        public static ListItemCollection GetItemsByLookup(this List InList, string FieldName, string Id)
        {
            var Query = "<View><Query><Where><Eq><FieldRef Name='{0}' LookupId='TRUE'/><Value Type='Lookup'>{1}</Value></Eq></Where></Query></View>";
            Query = string.Format(Query, FieldName, Id);
            return InList.GetItemsByCAML(Query);
        }
        public static ListItemCollection GetItemsByTwoLookup(this List InList, string FieldName1, string Id1, string FieldName2, string Id2)
        {
            var Query = "<View><Query><Where><And><Eq><FieldRef Name='{0}' LookupId='TRUE'/><Value Type='Lookup'>{1}</Value></Eq><Eq><FieldRef Name='{2}' LookupId='TRUE'/><Value Type='Lookup'>{3}</Value></Eq></And></Where></Query></View>";
            Query = string.Format(Query, FieldName1, Id1, FieldName2, Id2);
            return InList.GetItemsByCAML(Query);
        }
        public static ListItemCollection GetItemsByTwoStrings(this List InList, string FieldName1, string Id1, string FieldName2, string Id2)
        {
            var Query = "<View><Query><Where><And><Eq><FieldRef Name='{0}'/><Value Type='Text'>{1}</Value></Eq><Eq><FieldRef Name='{2}'/><Value Type='Text'>{3}</Value></Eq></And></Where></Query></View>";
            Query = string.Format(Query, FieldName1, Id1, FieldName2, Id2);
            return InList.GetItemsByCAML(Query);
        }
        public static ListItemCollection GetItemsByKey(this List InList, string FieldName, string Value)
        {
            var Query = "<View><Query><Where><Eq><FieldRef Name='{0}'/><Value Type='Text'>{1}</Value></Eq></Where></Query></View>";
            Query = string.Format(Query, FieldName, Value);
            return InList.GetItemsByCAML(Query);
        }
        public static ListItemCollection GetItemsByCAML(this List InList, string CAML)
        {
            var Query = new CamlQuery();
            Query.ViewXml = CAML;
            var InItems = InList.GetItems(Query);
            InList.Context.Load(InItems);

            InList.Context.ExecuteQuery();
            return InItems;
        }

        public static ListItemCollection GetAllItems(this List InList)
        {
            var Query = CamlQuery.CreateAllItemsQuery();
            var InItems = InList.GetItems(Query);
            InList.Context.Load(InItems);

            InList.Context.ExecuteQuery();
            return InItems;
        }
        public static List<ListItem> BatchQueryIn(this List InList, ListItemCollection InItems, string FieldWithId)
        {
            var ListData = new List<ListItem>();
            return BatchQueryIn(InList, InItems, FieldWithId, ListData);
        }
        public static List<ListItem> BatchQueryIn(this List InList, ListItemCollection InItems, string FieldWithId, string ExtendedCaml)
        {
            var ListData = new List<ListItem>();
            return BatchQueryIn(InList, InItems, FieldWithId, ExtendedCaml, ListData);
        }
        public static List<ListItem> BatchQueryIn(this List InList, ListItemCollection InItems, string FieldWithId, List<ListItem> ListData)
        {
            return BatchQueryIn(InList, InItems, FieldWithId, "", ListData);
        }
        public static List<ListItem> BatchQueryIn(this List InList, ListItemCollection InItems, string FieldWithId, string ExtendedCaml, List<ListItem> ListData)
        {
            if (InItems == null)
            {
                return ListData;
            }

            var BatchLen = 100;
            var QueryTemplate = "<View><Query><Where><In><FieldRef Name='ID'/><Values>{0}</Values></In></Where></Query></View>";
            if (!String.IsNullOrEmpty(ExtendedCaml))
            {
                QueryTemplate = "<View><Query><Where><And>" + ExtendedCaml + "<In><FieldRef Name='ID'/><Values>{0}</Values></In></And></Where></Query></View>";
            }

            var InQuery = "";

            var BatchQuery = new CamlQuery();
            var BatchTempI = BatchLen <= InItems.Count ? BatchLen : InItems.Count;
            foreach (var Item in InItems)
            {
                var IdField = (FieldLookupValue)Item[FieldWithId];

                InQuery += "<Value Type='Number'>" + IdField.LookupId + "</Value>";
                BatchTempI -= 1;
                if (BatchTempI == 0)
                {
                    InQuery = string.Format(QueryTemplate, InQuery);
                    BatchQuery.ViewXml = InQuery;
                    var BatchItems = InList.GetItems(BatchQuery);
                    BatchTempI = BatchLen;
                    InQuery = "";

                    InList.Context.Load(BatchItems);
                    InList.Context.ExecuteQuery();
                    ListData.AddRange(BatchItems);
                }
            }

            if ((BatchTempI > 0) && (BatchTempI < BatchLen))
            {
                InQuery = string.Format(QueryTemplate, InQuery);
                BatchQuery.ViewXml = InQuery;
                var BatchItems = InList.GetItems(BatchQuery);
                BatchTempI = BatchLen;
                InQuery = "";

                InList.Context.Load(BatchItems);
                InList.Context.ExecuteQuery();
                ListData.AddRange(BatchItems);
            }

            return ListData;
        }
    }
}
