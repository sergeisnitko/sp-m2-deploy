using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPF.Extentions
{
    public static class Simple
    {
        public static string StringValueOrEmpty(this object Value)
        {
            return Value != null ? Value.ToString() : "";
        }

        public static DateTime DateTimeByContext(this object value, ClientContext Context)
        {
            if (value == null)
                return DateTime.MaxValue;


            var spTimeZone = Context.Web.RegionalSettings.TimeZone;
            var SPDateTime = spTimeZone.UTCToLocalTime((DateTime)value);
            Context.Load(spTimeZone);
            Context.ExecuteQuery();

            return SPDateTime.Value;
        }

        public static string DateTimeValueOrEmpty(this object value, ClientContext Context)
        {
            return value.DateTimeValueOrEmpty(Context, "o");
        }

        public static string DateTimeValueOrEmpty(this object value, ClientContext Context, string format)
        {
            value = value.DateTimeByContext(Context);
            return (((DateTime)value != DateTime.MinValue) && ((DateTime)value != DateTime.MaxValue)) ? ((DateTime)value).ToString(format) : "";
        }

        public static string DateTimeValueOrEmpty(this object value)
        {
            return value.DateTimeValueOrEmpty("o");
        }

        public static string DateTimeValueOrEmpty(this object value, string format)
        {
            return value != null ? ((DateTime)value).ToString(format) : "";
        }

        public static int IntValueOrEmpty(this object Value)
        {
            var returnValue = 0;
            if (Value != null)
            {
                Int32.TryParse(Value.ToString(), out returnValue);
            }
            return returnValue;
        }

        public static void GenerateJavascriptFile(string Path, string[] JavascriptRows)
        {
            if (System.IO.File.Exists(Path))
            {
                System.IO.File.Delete(Path);
            }
            System.IO.File.Create(Path).Dispose();
            using (TextWriter tw = new StreamWriter(Path, true, Encoding.UTF8))
            {
                var Builder = new StringBuilder();
                Builder.AppendLine("\t" + string.Join("\n\t", JavascriptRows));
                tw.WriteLine(Builder.ToString());
                tw.Close();
            }
        }
    }
}
