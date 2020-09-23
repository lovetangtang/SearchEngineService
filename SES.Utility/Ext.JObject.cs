using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SES.Utility
{
    public static partial class Ext
    {

        public static bool ContainsKey(this JObject jobj, string key) {
            var obj = jobj[key];
            return obj != null;
        }
        public static string GetStringValue(this JObject jobj, string key)
        {
            var obj = jobj[key];
            if (obj == null) return "";
            return obj.ToString();
        }
        public static bool ContainsKey(this JToken jobj, string key)
        {
            var obj = jobj[key];
            return obj != null;
        }
        public static string GetStringValue(this JToken jobj, string key)
        {
            var obj = jobj[key];
            if (obj == null) return "";
            return obj.ToString();
        }
    }
}
