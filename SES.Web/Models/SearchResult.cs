using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace SES.Web.Models
{
    public class SearchResult
    {
        public int count { get; set; }
        public List<JObject> data { get; set; }
    }
}