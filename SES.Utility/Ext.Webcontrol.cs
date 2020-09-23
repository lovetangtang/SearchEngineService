using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI.WebControls;

namespace SES.Utility
{
    public static class ExtWebControl
    {

        /// <summary>
        /// 获取CheckBoxList选中项的值，每项值之间用separator分隔
        /// </summary>
        /// <param name="chkList"></param>
        /// <param name="separator">分隔符</param>
        /// <returns></returns>
        public static string GetValue(this CheckBoxList chkList, string separator)
        {
            string valList = "";
            for (int i = 0; i < chkList.Items.Count; i++)
            {
                if (chkList.Items[i].Selected)
                {
                    if (valList != "") valList += separator;
                    valList += chkList.Items[i].Value;
                }
            }
            return valList;
        }
        /// <summary>
        /// 设置CheckBoxList选中项
        /// </summary>
        /// <param name="chkList"></param>
        /// <param name="separator"></param>
        /// <param name="valList"></param>
        public static void SetValue(this CheckBoxList chkList, string separator, string valList)
        {
            valList = separator + valList + separator;
            for (int i = 0; i < chkList.Items.Count; i++)
            {
                if (valList.IndexOf(separator + chkList.Items[i].Value + separator) > -1)
                {
                    chkList.Items[i].Selected = true;
                }
            }
        }
    }
}
