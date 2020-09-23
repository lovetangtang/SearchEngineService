using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SES.Utility
{
    public static class ExtString
    {

        /// <summary>
        /// 获得指定长度的字符（如果操过长度是加省略号）
        /// </summary>
        /// <param name="word"></param>
        /// <param name="num"></param>
        /// <param name="showTips">截断字符后是否显示提示信息</param>
        /// <returns></returns>
        public static string StrCut(this string word, int num, bool showTips)
        {
            word = word.Replace("&nbsp;", " ");
            if (word.Length > num)
            {
                string str = word.Substring(0, num - 1).Replace(" ", "&nbsp;") + "..";
                if (showTips) str = "<span title='" + word + "'>" + str + "</span>";
                return str;
            }
            else
            {
                return word;
            }
        }

        public static bool IsDecimal(this string str) {
            decimal d;
            return decimal.TryParse(str, out d);
        }
        public static bool IsDateTime(this string str)
        {
            DateTime d;
            return DateTime.TryParse(str, out d);
        }
        public static bool IsInt(this string str)
        {
            long d;
            return long.TryParse(str, out d);
        }
    }
}
