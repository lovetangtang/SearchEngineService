using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Xml;

namespace SES.Utility
{
    public class Util
    {

        /// <summary>
        /// 对字符进行HTML编码(替换空格和换行符)
        /// </summary>
        /// <param name="str">需编码的字符串</param>
        /// <returns>为空时返回null</returns>
        public static string HtmlEncode(string str)
        {
            string s = str;
            if (s == null) { return null; }
            s = HttpContext.Current.Server.HtmlEncode(s);
            s = s.Replace(" ", "&nbsp;");
            s = s.Replace("\r\n", "<br />");
            return s;
        }

        /// <summary>
        /// 对字符串进HTML解码
        /// </summary>
        /// <param name="str">需解码的字符串</param>
        /// <returns>为空时返回null</returns>
        public static string HtmlDecode(string str)
        {
            string s = str;
            if (s == null) { return null; }
            s = s.Replace("&lt;", "<");
            s = s.Replace("&gt;", ">");
            s = s.Replace("&nbsp;", " ");
            s = s.Replace("<br />", "\r\n");
            return s;
        }

        /// <summary>
        /// 验证idList是否为数字，用逗号分隔(非数字的值删除)
        /// </summary>
        /// <param name="idList">1,3,4,5,6</param>
        /// <returns>没有合格数据时返回0</returns>
        public static string ValidateIdList(string idList)
        {
            if (string.IsNullOrEmpty(idList)) return "0";

            string[] arr = idList.Split(',');
            idList = "";
            for (int i = 0; i < arr.Length; i++)
            {
                int id;
                if (int.TryParse(arr[i], out id))
                {
                    if (idList != "") idList += ",";
                    idList += id;
                }
            }
            if (idList == "")
            {
                idList = "0";
            }
            return idList;
        }

        /// <summary>
        /// 读取appSettings中的value值
        /// </summary>
        /// <param name="key"></param>
        /// <param name="defaultVal"></param>
        /// <returns></returns>
        public static string GetAppSetting(string key, string defaultVal = "")
        {
            string val = ConfigurationManager.AppSettings[key];
            if (string.IsNullOrEmpty(val))
            {
                return defaultVal;
            }
            return val;
        }

        /// <summary>
        /// 获取Session中的值 
        /// </summary>
        /// <param name="key"></param>
        /// <param name="defaultVal"></param>
        /// <returns></returns>
        public static string GetSessionVal(string key, string defaultVal = "")
        {
            if (HttpContext.Current.Session[key] == null)
            {
                return defaultVal;
            }
            return HttpContext.Current.Session[key].ToString();
        }


        /// <summary>
        /// 得到MD5加密后的字符串(不可解密)
        /// </summary>
        /// <param name="input"></param>
        /// <returns></returns>
        public static string MD5(string input)
        {
            MD5CryptoServiceProvider md5Hasher = new MD5CryptoServiceProvider();
            byte[] data = md5Hasher.ComputeHash(Encoding.Default.GetBytes(input));
            StringBuilder sBuilder = new StringBuilder();
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            return sBuilder.ToString();
        }


        /// <summary>
        /// 判断值是否为int或""或null;
        /// </summary>
        /// <param name="val"></param>
        /// <returns></returns>
        public static bool IsIntOrEmpty(string val)
        {
            if (string.IsNullOrEmpty(val)) return true;
            int num = 0;
            return int.TryParse(val, out num);
        }


        public static string GetClientInfo()
        {
            HttpBrowserCapabilities bc = HttpContext.Current.Request.Browser;
            StringBuilder sb = new StringBuilder();
            sb.Append("浏览器=" + bc.Browser + ";");
            sb.Append("型态=" + bc.Type + ";");
            sb.Append("名称=" + bc.Browser + ";");
            sb.Append("版本=" + bc.Version + ";");
            sb.Append("使用平台=" + bc.Platform + ";");
            sb.Append("是否为测试版=" + bc.Beta + ";");
            sb.Append("是否为16 位的环境=" + bc.Win16 + ";");
            sb.Append("是否为32 位的环境=" + bc.Win32 + ";");
            sb.Append("是否支持框架(frame) =" + bc.Frames + ";");
            sb.Append("是否支持cookie =" + bc.Cookies + ";");
            sb.Append("是否支持activex controls =" + bc.ActiveXControls);
            return sb.ToString();
        }


        /// <summary>
        /// 获取客户端IP地址
        /// </summary>
        /// <returns></returns>
        public static string GetClientIP()
        {
            string result = HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"];
            if (null == result || result == String.Empty)
            {
                result = HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"];
            }
            if (null == result || result == String.Empty)
            {
                result = HttpContext.Current.Request.UserHostAddress;
            }
            return result;
        }

        /// <summary>
        /// 查询DataTable中第一列的值（所有行，每行间用,分隔）
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string GetIdList(DataTable dt)
        {
            string idList = "0";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string id = dt.Rows[i][0].ToString();
                if (!string.IsNullOrEmpty(id))
                {
                    idList += "," + id;
                }
            }
            return idList;
        }

        public static string GetQueryString(string key, string dftVal = "")
        {
            string val = HttpContext.Current.Request.QueryString[key];
            if (string.IsNullOrEmpty(val))
            {
                val = dftVal;
            }
            return val;
        }
        public static string GetFormString(string key, string dftVal = "")
        {
            string val = HttpContext.Current.Request.Form[key];
            if (string.IsNullOrEmpty(val))
            {
                val = dftVal;
            }
            return val;
        }
        public static string GetParamsString(string key, string dftVal = "")
        {
            string val = HttpContext.Current.Request.Params[key];
            if (string.IsNullOrEmpty(val))
            {
                val = dftVal;
            }
            return val;
        }

        /// <summary>
        /// 发起一个请求
        /// </summary>
        /// <param name="url"></param>
        /// <param name="method"></param>
        /// <param name="postData"></param>
        /// <returns></returns>
        public static string SendRequest(string url, string method, string postData, Dictionary<string, string> dict = null)
        {
            HttpWebRequest request;
            HttpWebResponse response;
            Stream stream;
            StreamReader streamReader;
            string result = "";

            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = method;

            //add by hg 20170412 添加头信息
            if (dict != null)
            {
                foreach (var item in dict)
                {
                    request.Headers.Add(item.Key, item.Value);
                }
            }

            if (method == "POST")
            {
                byte[] data = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = data.Length;
                Stream reqStream = request.GetRequestStream();
                reqStream.Write(data, 0, data.Length);
                reqStream.Close();
            }

            request.ContentType = "application/x-www-form-urlencoded";
            request.AllowAutoRedirect = true;
            response = (HttpWebResponse)request.GetResponse();
            stream = response.GetResponseStream();
            streamReader = new StreamReader(stream, Encoding.UTF8);
            result = streamReader.ReadToEnd();
            return result;
        }

        public static string SendRequest2(string url, string method, string postData)
        {
            HttpWebRequest request;
            HttpWebResponse response;
            Stream stream;
            StreamReader streamReader;
            string result = "";

            //add by hg 20160908 解决部分服务器上"未能为 SSL/TLS 安全通道建立信任关系"的问题
            //ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback(CheckValidationResult);//验证服务器证书回调自动验证

            request = (HttpWebRequest)WebRequest.Create(url);
            request.Method = method;
            request.ContentType = "application/x-www-form-urlencoded";
            request.AllowAutoRedirect = true;
            if (method == "POST")
            {
                byte[] data = Encoding.UTF8.GetBytes(postData);
                request.ContentLength = data.Length;
                Stream reqStream = request.GetRequestStream();
                reqStream.Write(data, 0, data.Length);
                reqStream.Close();
            }

            response = (HttpWebResponse)request.GetResponse();
            stream = response.GetResponseStream();
            streamReader = new StreamReader(stream, Encoding.UTF8);
            result = streamReader.ReadToEnd();

            return result;
        }


        public static string GetXmlNodeVal(string xml, string xpath)
        {
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.LoadXml(xml);
            }
            catch (Exception)
            {
                return "";
            }

            XmlNode node = doc.SelectSingleNode(xpath);
            if (node != null)
            {
                return node.InnerText;
            }
            return "";
        }



        /// <summary>
        /// 写入错误日志
        /// </summary>
        /// <param name="content"></param>
        public static void WriteLog(string content)
        {
            string LogPath = HttpContext.Current.Server.MapPath("/Error/");
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(LogPath);
            }

            DateTime date = DateTime.Now;
            string str = date.ToString("yyyy-MM-dd HH:mm:ss");
            str += "\t" + content;
            str += "\t[错误页面]: " + HttpContext.Current.Request.RawUrl + "\r\n\r\n\r\n";

            string fPath = LogPath + date.ToString("yyyyMMdd") + ".txt";
            using (FileStream stream = new FileStream(fPath, FileMode.Append))
            {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(str);
                stream.Write(bytes, 0, bytes.Length);
            }
        }

        /// <summary>
        /// 写入接口调用日志
        /// </summary>
        /// <param name="content"></param>
        public static void WriteApiLog(string content)
        {
            string LogPath = HttpContext.Current.Server.MapPath("/apilog/");
            if (!Directory.Exists(LogPath))
            {
                Directory.CreateDirectory(LogPath);
            }

            DateTime date = DateTime.Now;
            string str = date.ToString("yyyy-MM-dd HH:mm:ss");
            str += "\t" + content;
            str += "\t[请求地址]: " + HttpContext.Current.Request.RawUrl + "\r\n\r\n\r\n";

            string fPath = LogPath + date.ToString("yyyyMMdd") + ".txt";
            using (FileStream stream = new FileStream(fPath, FileMode.Append))
            {
                byte[] bytes = System.Text.Encoding.UTF8.GetBytes(str);
                stream.Write(bytes, 0, bytes.Length);
            }
        }



        /// <summary>
        /// 运JS脚本
        /// </summary>
        /// <param name="sMessage"></param>
        public static void ExecJS(string sMessage)
        {
            HttpContext.Current.Response.Write("<script type=\"text/javascript\">" + sMessage + "</script>");
            HttpContext.Current.Response.End();
        }

        public static void ExecJS2(System.Web.UI.Page page, string js)
        {
            page.ClientScript.RegisterStartupScript(page.GetType(), "Key", js, true);
        }

        public static void ShowMsg(System.Web.UI.Page page, string msg)
        {
            page.ClientScript.RegisterStartupScript(page.GetType(), "showmsg", "alert('" + msg + "');", true);
        }

        /// <summary>
        /// 显示提示框并刷新当前页
        /// </summary>
        /// <param name="sMessage"></param>
        public static void ShowMsg(string sMessage)
        {
            HttpContext.Current.Response.Clear();
            sMessage = sMessage.Replace("'", "\"").Replace("\r\n", "");
            HttpContext.Current.Response.Write("<script>alert('" + sMessage + "');history.back();</script>");
            HttpContext.Current.Response.End();
        }

        /// <summary>
        /// 显示提示框
        /// </summary>
        /// <param name="sMessage"></param>
        public static void ShowMsgLayer(System.Web.UI.Page page, string sMessage)
        {
            sMessage = sMessage.Replace("'", "\"").Replace("\r\n", "");
            page.ClientScript.RegisterStartupScript(page.GetType(), "layer", "layer.msg('" + sMessage + "');", true);
        }

        /// <summary>
        /// 显示提示框
        /// </summary>
        /// <param name="sMessage"></param>
        public static void ShowErrMsgLayer(System.Web.UI.Page page, string sMessage)
        {
            sMessage = sMessage.Replace("'", "\"").Replace("\r\n", "");

            page.ClientScript.RegisterStartupScript(page.GetType(), "layer", "layer.open({ title: '系统提示'  ,content: '" + sMessage + "'});", true);
        }

        /// <summary>
        /// 显示提示框,2秒后跳到指定页
        /// </summary>
        /// <param name="sMessage"></param>
        public static void ShowMsgLayerRedirect(System.Web.UI.Page page, string sMessage, string url, int time = 2)
        {
            sMessage = sMessage.Replace("'", "\"").Replace("\r\n", "");
            page.ClientScript.RegisterStartupScript(page.GetType(), "layer", "layer.msg('" + sMessage + "',{time: " + time * 1000 + "},function(){window.location='" + url + "'});", true);
        }

        public static string GetJsStr(string msg)
        {
            return msg.Replace("'", "\"").Replace("\r\n", "");
        }

        /// <summary>
        /// 显示提示框并刷新当前页
        /// </summary>
        /// <param name="sMessage"></param>
        public static void ShowMsg(string sMessage, string url)
        {
            sMessage = sMessage.Replace("'", "\"").Replace("\r\n", "");
            HttpContext.Current.Response.Write("<script type='text/javascript'>alert('" + sMessage + "');window.location='" + url + "';</script>");
            HttpContext.Current.Response.End();
        }


        /// <summary>
        /// 显示提示，非JS
        /// </summary>
        /// <param name="sData"></param>
        /// <param name="url"></param>
        public static void AlertDiv(string sData, string url)
        {
            HttpContext.Current.Response.Write("<div style='width:100%;font-size:30px;color:red;text-align:center;padding-top:200px;'>" + sData + "</div><script type='text/javascript'>setTimeout(\"location.href='" + url + "'\",1000);</script>");
            HttpContext.Current.Response.End();
        }



        /// <summary>
        /// 返回声母
        /// </summary>
        /// <param name="cn"></param>
        /// <returns></returns>
        static public string GetSpell(string cn)
        {
            byte[] arrCN = Encoding.Default.GetBytes(cn);
            if (arrCN.Length > 1)
            {
                int area = (short)arrCN[0];
                int pos = (short)arrCN[1];
                int code = (area << 8) + pos;
                int[] areacode = { 45217, 45253, 45761, 46318, 46826, 47010, 47297, 47614, 48119, 48119, 49062, 49324, 49896, 50371, 50614, 50622, 50906, 51387, 51446, 52218, 52698, 52698, 52698, 52980, 53689, 54481 };
                for (int i = 0; i < 26; i++)
                {
                    int max = 55290;
                    if (i != 25) max = areacode[i + 1];
                    if (areacode[i] <= code && code < max)
                    {
                        return Encoding.Default.GetString(new byte[] { (byte)(65 + i) });
                    }
                }
                return cn;
            }
            else return cn;
        }


        /// <summary>
        /// 查询DataTable中单列的多行数据(多个值之间用,分隔)
        /// </summary>
        /// <param name="dt"></param>
        /// <param name="columnName"></param>
        /// <returns></returns>
        public static string GetColumnVal(DataTable dt, string columnName)
        {
            string valList = "";
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                string val = dt.Rows[i][columnName].ToString();
                if (!string.IsNullOrEmpty(val))
                {
                    if (valList.Length > 0) valList += ",";
                    valList += val;
                }
            }
            return valList;
        }


        /// <summary>
        /// 随机字符串类别
        /// </summary>
        public enum NonceType
        {
            Upper,//仅大写
            Lower,//仅小写
            UpperLower,//大小写混合
            Int//仅数字
        }
        /// <summary>
        /// 获取随机字符串
        /// </summary>
        /// <param name="nonceType"></param>
        /// <param name="length"></param>
        /// <returns></returns>
        public static string GetNonceStr(NonceType nonceType, int length)
        {
            List<char> arr = new List<char>();
            if (nonceType == NonceType.Upper)
            {
                for (int i = 65; i <= 90; i++)
                {
                    arr.Add((char)i);
                }
            }
            else if (nonceType == NonceType.Lower)
            {
                for (int i = 97; i <= 122; i++)
                {
                    arr.Add((char)i);
                }
            }
            else if (nonceType == NonceType.Int)
            {
                for (int i = 48; i <= 57; i++)
                {
                    arr.Add((char)i);
                }
            }
            else
            {
                for (int i = 65; i <= 90; i++)
                {
                    arr.Add((char)i);
                }
                for (int i = 97; i <= 122; i++)
                {
                    arr.Add((char)i);
                }
            }

            string nonceStr = "";
            Random rand = new Random();
            for (int i = 0; i < length; i++)
            {
                int n = rand.Next(0, arr.Count);
                nonceStr += arr[n];
            }
            return nonceStr;
        }

        /// <summary>
        /// 获取数据字典类型的设置信息
        /// </summary>
        /// <param name="key"></param>
        /// <param name="dftDict"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetAppSettingDict(string key, Dictionary<string, string> dftDict)
        {
            string val = GetAppSetting(key);
            if (val == "") return dftDict;
            //val: {QMYSSYS:42B33E5B-7AAB-44C7-A70D-4F2E25FC44D6,MYSOFTSYS:A454D305-FFE9-46CF-BA9A-9F216696753E}
            val = val.Replace("{", "{\"").Replace("}", "\"}").Replace(":", "\":\"").Replace(",", "\",\"");
            try
            {
                return JsonConvert.DeserializeObject<Dictionary<string, string>>(val);
            }
            catch (Exception)
            {
                return null;
            }

        }



        public static void ResponseStr(string str)
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.Write(str);
        }
        public static void ResponseJsonResult(JsonResult jsonResult)
        {
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.Write(jsonResult.ToJson());
        }
        public static void ResponseJsonResult(int errcode, object errmsg)
        {
            ResponseJsonResult(new JsonResult() { errcode = errcode, errmsg = errmsg });
        }

        /// <summary>
        /// 验证查询Sql语句是包含除查询外的非法字符
        /// </summary>
        /// <param name="sql">sql为""时不验证</param>
        /// <returns>不合法返回错误信息</returns>
        public static string ValidateQuerySql(string sql)
        {
            if (string.IsNullOrEmpty(sql)) return "";

            //判断sql语句中是否包含非法语句
            string pattern = "exec|create|insert|delete|update|drop|alter|insert";
            string[] arr = pattern.Split('|');
            string pattern2 = "(\\b" + string.Join("\\b)|(\\b", arr) + "\\b)";
            Regex regex = new Regex(pattern2, RegexOptions.IgnoreCase);
            if (regex.IsMatch(sql))
            {
                return "Sql语句不合法！";
            }

            return "";
        }
    }

    /// <summary>
    /// 返回json数据对象
    /// </summary>
    public class JsonResult
    {
        /// <summary>
        /// 错误编码 0：成功    否则为有错误(500)
        /// </summary>
        public int errcode { get; set; }
        /// <summary>
        /// 成功时返回成功数据，失败时返回错误信息
        /// </summary>
        public object errmsg { get; set; }
    }


    /// <summary>
    /// Layui分页控件返回数据源
    /// </summary>
    public class LayuiPagerResult
    {
        public int code { get; set; }
        public string msg { get; set; }
        public int count { get; set; }
        public object data { get; set; }
    }

    //#region Excel操作帮助类
    //public class ExcelHelper
    //{
    //    /// <summary>
    //    /// 将excel导入到datatable
    //    /// </summary>
    //    /// <param name="filePath">excel路径</param>
    //    /// <param name="isColumnName">第一行是否是列名</param>
    //    /// <returns>返回datatable</returns>
    //    public static DataTable ExcelToDataTable(string filePath, bool isColumnName)
    //    {
    //        DataTable dataTable = null;
    //        FileStream fs = null;
    //        DataColumn column = null;
    //        DataRow dataRow = null;
    //        IWorkbook workbook = null;
    //        ISheet sheet = null;
    //        IRow row = null;
    //        ICell cell = null;
    //        int startRow = 0;
    //        try
    //        {
    //            using (fs = File.OpenRead(filePath))
    //            {
    //                // 2007版本
    //                if (filePath.IndexOf(".xlsx") > 0)
    //                    workbook = new XSSFWorkbook(fs);
    //                // 2003版本
    //                else if (filePath.IndexOf(".xls") > 0)
    //                    workbook = new HSSFWorkbook(fs);

    //                if (workbook != null)
    //                {
    //                    sheet = workbook.GetSheetAt(0);//读取第一个sheet，当然也可以循环读取每个sheet
    //                    int a = workbook.NumberOfSheets;
    //                    dataTable = new DataTable();
    //                    if (sheet != null)
    //                    {
    //                        int rowCount = sheet.LastRowNum;//总行数
    //                        if (rowCount > 0)
    //                        {
    //                            IRow firstRow = sheet.GetRow(0);//第一行
    //                            int cellCount = firstRow.LastCellNum;//列数

    //                            //构建datatable的列
    //                            if (isColumnName)
    //                            {
    //                                startRow = 1;//如果第一行是列名，则从第二行开始读取
    //                                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
    //                                {
    //                                    cell = firstRow.GetCell(i);
    //                                    if (cell != null)
    //                                    {
    //                                        if (cell.StringCellValue != null)
    //                                        {
    //                                            column = new DataColumn(cell.StringCellValue);
    //                                            dataTable.Columns.Add(column);
    //                                        }
    //                                    }
    //                                }
    //                            }
    //                            else
    //                            {
    //                                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
    //                                {
    //                                    column = new DataColumn("column" + (i + 1));
    //                                    dataTable.Columns.Add(column);
    //                                }
    //                            }

    //                            //填充行
    //                            for (int i = startRow; i <= rowCount; ++i)
    //                            {
    //                                row = sheet.GetRow(i);
    //                                if (row == null) continue;

    //                                dataRow = dataTable.NewRow();
    //                                for (int j = row.FirstCellNum; j < cellCount; ++j)
    //                                {
    //                                    cell = row.GetCell(j);
    //                                    if (cell == null)
    //                                    {
    //                                        dataRow[j] = "";
    //                                    }
    //                                    else
    //                                    {
    //                                        //CellType(Unknown = -1,Numeric = 0,String = 1,Formula = 2,Blank = 3,Boolean = 4,Error = 5,)
    //                                        switch (cell.CellType)
    //                                        {
    //                                            case CellType.Blank:
    //                                                dataRow[j] = "";
    //                                                break;
    //                                            case CellType.Numeric:
    //                                                short format = cell.CellStyle.DataFormat;
    //                                                //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理
    //                                                if (format == 14 || format == 31 || format == 57 || format == 58)
    //                                                    dataRow[j] = cell.DateCellValue;
    //                                                else
    //                                                    dataRow[j] = cell.NumericCellValue;
    //                                                break;
    //                                            case CellType.String:
    //                                                dataRow[j] = cell.StringCellValue;
    //                                                break;
    //                                        }
    //                                    }
    //                                }
    //                                dataTable.Rows.Add(dataRow);
    //                            }
    //                        }
    //                    }
    //                }
    //            }
    //            return dataTable;
    //        }
    //        catch (Exception)
    //        {
    //            if (fs != null)
    //            {
    //                fs.Close();
    //            }
    //            return null;
    //        }
    //    }

    //    /// <summary>
    //    /// 将datatable导出到excel
    //    /// </summary>
    //    /// <param name="dt"></param>
    //    /// <param name="filePath"></param>
    //    /// <returns></returns>
    //    public static bool DataTableToExcel(DataTable dt, string filePath)
    //    {
    //        bool result = false;
    //        IWorkbook workbook = null;
    //        FileStream fs = null;
    //        IRow row = null;
    //        ISheet sheet = null;
    //        ICell cell = null;

    //        try
    //        {
    //            if (dt != null && dt.Rows.Count > 0)
    //            {
    //                // 2007版本
    //                if (filePath.IndexOf(".xlsx") > 0)
    //                    workbook = new XSSFWorkbook();
    //                // 2003版本
    //                else if (filePath.IndexOf(".xls") > 0)
    //                    workbook = new HSSFWorkbook();
    //                else
    //                    return false;

    //                sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
    //                int rowCount = dt.Rows.Count;//行数
    //                int columnCount = dt.Columns.Count;//列数

    //                //创建列样式（格式等）
    //                NPOI.SS.UserModel.ICellStyle dateStyle = workbook.CreateCellStyle();
    //                //设置数据显示格式
    //                NPOI.SS.UserModel.IDataFormat format = workbook.CreateDataFormat();
    //                dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

    //                //设置列头
    //                row = sheet.CreateRow(0);//excel第一行设为列头
    //                for (int c = 0; c < columnCount; c++)
    //                {
    //                    cell = row.CreateCell(c);
    //                    cell.SetCellValue(dt.Columns[c].ColumnName);
    //                }

    //                //设置每行每列的单元格,
    //                for (int i = 0; i < rowCount; i++)
    //                {
    //                    row = sheet.CreateRow(i + 1);
    //                    for (int j = 0; j < columnCount; j++)
    //                    {
    //                        cell = row.CreateCell(j);//excel第二行开始写入数据
    //                        //

    //                        cell.SetCellValue(dt.Rows[i][j].ToString());
    //                        #region 根据datatable列类型写入对应类型的数据
    //                        //根据datatable列类型写入对应类型的数据
    //                        //switch (column.DataType.ToString())
    //                        //{
    //                        //    case "System.String"://字符串类型
    //                        //        newCell.SetCellValue(drValue);
    //                        //        break;
    //                        //    case "System.DateTime"://日期类型
    //                        //        DateTime dateV;
    //                        //        DateTime.TryParse(drValue, out dateV);
    //                        //        newCell.SetCellValue(dateV);

    //                        //        newCell.CellStyle = dateStyle;//格式化显示
    //                        //        break;
    //                        //    case "System.Boolean"://布尔型
    //                        //        bool boolV = false;
    //                        //        bool.TryParse(drValue, out boolV);
    //                        //        newCell.SetCellValue(boolV);
    //                        //        break;
    //                        //    case "System.Int16"://整型
    //                        //    case "System.Int32":
    //                        //    case "System.Int64":
    //                        //    case "System.Byte":
    //                        //        int intV = 0;
    //                        //        int.TryParse(drValue, out intV);
    //                        //        newCell.SetCellValue(intV);
    //                        //        break;
    //                        //    case "System.Decimal"://浮点型
    //                        //    case "System.Double":
    //                        //        double doubV = 0;
    //                        //        double.TryParse(drValue, out doubV);
    //                        //        newCell.SetCellValue(doubV);
    //                        //        break;
    //                        //    case "System.DBNull"://空值处理
    //                        //        newCell.SetCellValue("");
    //                        //        break;
    //                        //    default:
    //                        //        newCell.SetCellValue("");
    //                        //        break;
    //                        //}
    //                        #endregion
    //                    }
    //                }
    //                using (fs = File.OpenWrite(filePath))//@"D:/myxls.xlsx"
    //                {
    //                    workbook.Write(fs);//向打开的这个xls文件中写入数据
    //                    result = true;
    //                }
    //            }
    //            return result;
    //        }
    //        catch (Exception ex)
    //        {
    //            if (fs != null)
    //            {
    //                fs.Close();
    //            }
    //            return false;
    //        }
    //    }

    //    public byte[] getExcel(DataTable dt, string sheetName)
    //    {

    //        IWorkbook book = new HSSFWorkbook();
    //        if (dt.Rows.Count < 65535)
    //            this.DataWrite2Sheet(dt, 0, dt.Rows.Count - 1, book, sheetName);
    //        else
    //        {
    //            int page = dt.Rows.Count / 65535;
    //            for (int i = 0; i < page; i++)
    //            {
    //                int start = i * 65535;
    //                int end = (i * 65535) + 65535 - 1;
    //                this.DataWrite2Sheet(dt, start, end, book, sheetName + i.ToString());
    //            }
    //            int lastPageItemCount = dt.Rows.Count % 65535;
    //            this.DataWrite2Sheet(dt, dt.Rows.Count - lastPageItemCount, lastPageItemCount, book, sheetName + page.ToString());
    //        }
    //        MemoryStream ms = new MemoryStream();
    //        book.Write(ms);
    //        return ms.ToArray();
    //    }
    //    public void DataWrite2Sheet(DataTable dt, int startRow, int endRow, IWorkbook book, string sheetName)
    //    {
    //        ISheet sheet = book.CreateSheet(sheetName);
    //        IRow header = sheet.CreateRow(0);
    //        for (int i = 0; i < dt.Columns.Count; i++)
    //        {
    //            ICell cell = header.CreateCell(i);
    //            string val = dt.Columns[i].Caption ?? dt.Columns[i].ColumnName;
    //            cell.SetCellValue(val);
    //        }
    //        int rowIndex = 1;
    //        for (int i = startRow; i <= endRow; i++)
    //        {
    //            DataRow dtRow = dt.Rows[i];
    //            IRow excelRow = sheet.CreateRow(rowIndex++);
    //            for (int j = 0; j < dtRow.ItemArray.Length; j++)
    //            {
    //                excelRow.CreateCell(j).SetCellValue(dtRow[j].ToString());
    //            }
    //        }

    //    }
    //    /// <summary>
    //    /// 将datatable导出到excel 
    //    /// </summary>
    //    /// <param name="dt"></param>
    //    /// <param name="isXlsx">是否为Xlsx（2007）版本</param>
    //    /// <returns></returns>
    //    public static MemoryStream DataTableToExcel(DataTable dt, bool isXlsx)
    //    {
    //        IWorkbook workbook = null;
    //        FileStream fs = null;
    //        IRow row = null;
    //        ISheet sheet = null;
    //        ICell cell = null;

    //        try
    //        {
    //            if (dt != null && dt.Rows.Count > 0)
    //            {
    //                // 2007版本
    //                if (isXlsx)
    //                    workbook = new XSSFWorkbook();
    //                // 2003版本
    //                else
    //                    workbook = new HSSFWorkbook();

    //                sheet = workbook.CreateSheet("Sheet0");//创建一个名称为Sheet0的表
    //                int rowCount = dt.Rows.Count;//行数
    //                int columnCount = dt.Columns.Count;//列数

    //                //创建列样式（格式等）
    //                NPOI.SS.UserModel.ICellStyle dateStyle = workbook.CreateCellStyle();
    //                //设置数据显示格式
    //                NPOI.SS.UserModel.IDataFormat format = workbook.CreateDataFormat();
    //                dateStyle.DataFormat = format.GetFormat("yyyy-mm-dd");

    //                //设置列头
    //                row = sheet.CreateRow(0);//excel第一行设为列头
    //                for (int c = 0; c < columnCount; c++)
    //                {
    //                    cell = row.CreateCell(c);
    //                    cell.SetCellValue(dt.Columns[c].ColumnName);
    //                }

    //                //设置每行每列的单元格,
    //                for (int i = 0; i < rowCount; i++)
    //                {
    //                    row = sheet.CreateRow(i + 1);
    //                    for (int j = 0; j < columnCount; j++)
    //                    {
    //                        cell = row.CreateCell(j);//excel第二行开始写入数据
    //                        //

    //                        cell.SetCellValue(dt.Rows[i][j].ToString());
    //                        #region 根据datatable列类型写入对应类型的数据
    //                        //根据datatable列类型写入对应类型的数据
    //                        //switch (column.DataType.ToString())
    //                        //{
    //                        //    case "System.String"://字符串类型
    //                        //        newCell.SetCellValue(drValue);
    //                        //        break;
    //                        //    case "System.DateTime"://日期类型
    //                        //        DateTime dateV;
    //                        //        DateTime.TryParse(drValue, out dateV);
    //                        //        newCell.SetCellValue(dateV);

    //                        //        newCell.CellStyle = dateStyle;//格式化显示
    //                        //        break;
    //                        //    case "System.Boolean"://布尔型
    //                        //        bool boolV = false;
    //                        //        bool.TryParse(drValue, out boolV);
    //                        //        newCell.SetCellValue(boolV);
    //                        //        break;
    //                        //    case "System.Int16"://整型
    //                        //    case "System.Int32":
    //                        //    case "System.Int64":
    //                        //    case "System.Byte":
    //                        //        int intV = 0;
    //                        //        int.TryParse(drValue, out intV);
    //                        //        newCell.SetCellValue(intV);
    //                        //        break;
    //                        //    case "System.Decimal"://浮点型
    //                        //    case "System.Double":
    //                        //        double doubV = 0;
    //                        //        double.TryParse(drValue, out doubV);
    //                        //        newCell.SetCellValue(doubV);
    //                        //        break;
    //                        //    case "System.DBNull"://空值处理
    //                        //        newCell.SetCellValue("");
    //                        //        break;
    //                        //    default:
    //                        //        newCell.SetCellValue("");
    //                        //        break;
    //                        //}
    //                        #endregion
    //                    }
    //                }
    //                using (MemoryStream ms = new MemoryStream())
    //                {
    //                    workbook.Write(ms);

    //                    workbook = null;
    //                    return ms;
    //                }
    //            }
    //            return null;
    //        }
    //        catch (Exception ex)
    //        {
    //            if (fs != null)
    //            {
    //                fs.Close();
    //            }
    //            return null;
    //        }
    //    }



    //    public static void ExportExcel(string fileName, DataTable dt, string columnNames, Hashtable note)
    //    {
    //        MemoryStream ms = new MemoryStream();
    //        IWorkbook workbook = new HSSFWorkbook();
    //        ISheet sheet = workbook.CreateSheet();
    //        IRow headerRow = sheet.CreateRow(0);
    //        int cellCount = dt.Columns.Count;

    //        //设置字体颜色
    //        //IFont font = workbook.CreateFont();
    //        //font.Color = 10;

    //        //header. 
    //        string[] arr = columnNames.Split('|');
    //        for (int i = 0; i < cellCount; i++)
    //        {
    //            ICell cell = headerRow.CreateCell(i);
    //            string colName = "";
    //            if (string.IsNullOrEmpty(columnNames))
    //            {
    //                colName = dt.Columns[i].ColumnName;
    //            }
    //            else
    //            {
    //                colName = arr[i];
    //            }
    //            cell.SetCellValue(colName);

    //            //表头添加备注
    //            if (note.ContainsKey(colName))
    //            {
    //                HSSFPatriarch patr = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
    //                HSSFComment comment1 = patr.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, 2, 2, 10, 10));
    //                comment1.String = new HSSFRichTextString(note[colName].ToString());
    //                comment1.Author = "";
    //                cell.CellComment = comment1;
    //                //设置字体颜色
    //                //cell.CellStyle.SetFont(font);
    //            }
    //        }

    //        // value. 
    //        int rowIndex = 1;
    //        for (int i = 0; i < dt.Rows.Count; i++)
    //        {
    //            IRow dataRow = sheet.CreateRow(rowIndex);

    //            for (int j = 0; j < cellCount; j++)
    //            {
    //                dataRow.CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
    //            }
    //            rowIndex++;
    //        }

    //        workbook.Write(ms);
    //        ms.Flush();
    //        ms.Position = 0;

    //        HttpContext.Current.Response.ClearContent();
    //        HttpContext.Current.Response.BufferOutput = true;
    //        HttpContext.Current.Response.Charset = "utf-8";
    //        HttpContext.Current.Response.ContentType = "application/excel";
    //        HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8;

    //        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;fileName=" + HttpContext.Current.Server.UrlEncode(fileName) + ".xls");
    //        HttpContext.Current.Response.BinaryWrite(ms.ToArray());
    //        HttpContext.Current.Response.End();
    //    }




    //    /// <summary>
    //    /// 导出Excel
    //    /// </summary>
    //    /// <param name="fileName">导出文件名</param>
    //    /// <param name="dt">数据源</param>
    //    /// <param name="columnNames">数据源字段名集合（|分隔）</param>
    //    /// <param name="headerNames">Excel字段名集合（|分隔）</param>
    //    /// <param name="note">字段备注</param>
    //    public static void ExportExcel(string fileName, DataTable dt, string columnNames, string headerNames, Hashtable note)
    //    {
    //        MemoryStream ms = new MemoryStream();
    //        IWorkbook workbook = new HSSFWorkbook();
    //        ISheet sheet = workbook.CreateSheet();
    //        IRow headerRow = sheet.CreateRow(0);
    //        //int cellCount = dt.Columns.Count;

    //        //设置字体颜色
    //        //IFont font = workbook.CreateFont();
    //        //font.Color = 10;

    //        //header. 
    //        string[] arrColumn = columnNames.Split('|');
    //        string[] arrHeader = headerNames.Split('|');
    //        if (arrColumn.Length != arrHeader.Length)
    //        {
    //            HttpContext.Current.Response.Write("传入字段序列有误!");
    //            HttpContext.Current.Response.End();
    //        }


    //        ICellStyle cellStyle = workbook.CreateCellStyle();
    //        //设置单元格上下左右边框线
    //        cellStyle.BorderTop = BorderStyle.Thin;
    //        cellStyle.BorderBottom = BorderStyle.Thin;
    //        cellStyle.BorderLeft = BorderStyle.Thin;
    //        cellStyle.BorderRight = BorderStyle.Thin;


    //        for (int i = 0; i < arrHeader.Length; i++)
    //        {
    //            ICell cell = headerRow.CreateCell(i);
    //            cell.CellStyle = cellStyle;
    //            string colName = arrHeader[i];
    //            cell.SetCellValue(colName);

    //            //表头添加备注
    //            if (note != null && note.ContainsKey(colName))
    //            {
    //                HSSFPatriarch patr = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
    //                HSSFComment comment1 = patr.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, 2, 2, 10, 10));
    //                comment1.String = new HSSFRichTextString(note[colName].ToString());
    //                comment1.Author = "";
    //                cell.CellComment = comment1;
    //                //设置字体颜色
    //                //cell.CellStyle.SetFont(font);
    //            }
    //        }

    //        // value. 
    //        int rowIndex = 1;
    //        for (int i = 0; i < dt.Rows.Count; i++)
    //        {
    //            IRow dataRow = sheet.CreateRow(rowIndex);

    //            for (int j = 0; j < arrColumn.Length; j++)
    //            {
    //                ICell cell = dataRow.CreateCell(j);
    //                string colName = arrColumn[j];
    //                string colVal = dt.Rows[i][colName].ToString();
    //                cell.SetCellValue(colVal);
    //                cell.CellStyle = cellStyle;
    //            }
    //            rowIndex++;
    //        }

    //        workbook.Write(ms);
    //        ms.Flush();
    //        ms.Position = 0;

    //        HttpContext.Current.Response.ClearContent();
    //        HttpContext.Current.Response.BufferOutput = true;
    //        HttpContext.Current.Response.Charset = "utf-8";
    //        HttpContext.Current.Response.ContentType = "application/excel";
    //        HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8;

    //        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;fileName=" + HttpContext.Current.Server.UrlEncode(fileName) + ".xls");
    //        HttpContext.Current.Response.BinaryWrite(ms.ToArray());
    //        HttpContext.Current.Response.End();
    //    }

    //    /// <summary>
    //    /// 导出Excel
    //    /// </summary>
    //    /// <param name="fileName">导出文件名</param>
    //    /// <param name="dt">数据源</param>
    //    /// <param name="columnNames">数据源字段名集合（|分隔）</param>
    //    /// <param name="headerNames">Excel字段名集合（|分隔）</param>
    //    /// <param name="note">字段备注</param>
    //    public static void SaveExcel(string fileName, DataTable dt, string columnNames, string headerNames, Hashtable note)
    //    {
    //        MemoryStream ms = new MemoryStream();
    //        IWorkbook workbook = new HSSFWorkbook();
    //        ISheet sheet = workbook.CreateSheet();
    //        IRow headerRow = sheet.CreateRow(0);
    //        //int cellCount = dt.Columns.Count;

    //        //设置字体颜色
    //        //IFont font = workbook.CreateFont();
    //        //font.Color = 10;

    //        //header. 
    //        string[] arrColumn = columnNames.Split('|');
    //        string[] arrHeader = headerNames.Split('|');
    //        if (arrColumn.Length != arrHeader.Length)
    //        {
    //            HttpContext.Current.Response.Write("传入字段序列有误!");
    //            HttpContext.Current.Response.End();
    //        }


    //        ICellStyle cellStyle = workbook.CreateCellStyle();
    //        //设置单元格上下左右边框线
    //        cellStyle.BorderTop = BorderStyle.Thin;
    //        cellStyle.BorderBottom = BorderStyle.Thin;
    //        cellStyle.BorderLeft = BorderStyle.Thin;
    //        cellStyle.BorderRight = BorderStyle.Thin;


    //        for (int i = 0; i < arrHeader.Length; i++)
    //        {
    //            ICell cell = headerRow.CreateCell(i);
    //            cell.CellStyle = cellStyle;
    //            string colName = arrHeader[i];
    //            cell.SetCellValue(colName);

    //            //表头添加备注
    //            if (note != null && note.ContainsKey(colName))
    //            {
    //                HSSFPatriarch patr = (HSSFPatriarch)sheet.CreateDrawingPatriarch();
    //                HSSFComment comment1 = patr.CreateComment(new HSSFClientAnchor(0, 0, 0, 0, 2, 2, 10, 10));
    //                comment1.String = new HSSFRichTextString(note[colName].ToString());
    //                comment1.Author = "";
    //                cell.CellComment = comment1;
    //                //设置字体颜色
    //                //cell.CellStyle.SetFont(font);
    //            }
    //        }

    //        // value. 
    //        int rowIndex = 1;
    //        for (int i = 0; i < dt.Rows.Count; i++)
    //        {
    //            IRow dataRow = sheet.CreateRow(rowIndex);

    //            for (int j = 0; j < arrColumn.Length; j++)
    //            {
    //                ICell cell = dataRow.CreateCell(j);
    //                string colName = arrColumn[j];
    //                cell.SetCellValue(dt.Rows[i][colName].ToString());
    //                cell.CellStyle = cellStyle;
    //            }
    //            rowIndex++;
    //        }

    //        workbook.Write(ms);
    //        ms.Flush();
    //        ms.Position = 0;

    //        byte[] data = ms.ToArray();
    //        using (FileStream fs = new FileStream(HttpContext.Current.Server.MapPath(fileName), FileMode.OpenOrCreate, FileAccess.ReadWrite))
    //        {
    //            fs.Write(data, 0, data.Length);
    //        }

    //        /*
    //        HttpContext.Current.Response.ClearContent();
    //        HttpContext.Current.Response.BufferOutput = true;
    //        HttpContext.Current.Response.Charset = "utf-8";
    //        HttpContext.Current.Response.ContentType = "application/excel";
    //        HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.UTF8;

    //        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;fileName=" + HttpContext.Current.Server.UrlEncode(fileName) + ".xls");
    //        HttpContext.Current.Response.BinaryWrite(ms.ToArray());
    //        HttpContext.Current.Response.End();
    //        */
    //    }


    //    /// <summary>
    //    /// 
    //    /// </summary>
    //    /// <param name="filePath">Excel文件绝对路 如:d:\WorkPannel\1.xls</param>
    //    /// <param name="sheetIndex">Excel中表索引（从0开始）</param>
    //    /// <param name="maxCellIndex">要读取最大列索引（从0开始）</param>
    //    /// <param name="columnNames">Table列名</param>
    //    /// <returns></returns>
    //    public static DataTable GetDtByFilePath(string filePath, int sheetIndex, int maxCellIndex, string columnNames)
    //    {
    //        HSSFWorkbook wb = new HSSFWorkbook(new FileStream(filePath, FileMode.Open));
    //        ISheet sheet = wb.GetSheetAt(sheetIndex);
    //        int totalRowNum = sheet.LastRowNum;
    //        DataTable dt = new DataTable();
    //        string[] arr = columnNames.Split('|');
    //        for (int i = 0; i <= maxCellIndex; i++)
    //        {
    //            dt.Columns.Add(new DataColumn(arr[i], typeof(string)));
    //        }


    //        for (int i = 1; i <= totalRowNum; i++)
    //        {
    //            IRow row = sheet.GetRow(i);
    //            DataRow dataRow = dt.NewRow();
    //            for (int j = 0; j <= maxCellIndex; j++)
    //            {
    //                //dataRow[arr[j]] = GetValue(row.Cells[j]).Trim();
    //                dataRow[arr[j]] = GetValue(row.GetCell(j)).Trim();
    //            }
    //            dt.Rows.Add(dataRow);
    //        }

    //        return dt;
    //    }
    //    private static string GetValue(ICell cell)
    //    {
    //        if (cell == null)
    //        {
    //            return "";
    //        }
    //        if (cell.CellType == CellType.Numeric)
    //        {
    //            if (DateUtil.IsCellDateFormatted(cell))
    //            {
    //                return cell.DateCellValue.ToString("yyyy-MM-dd");
    //            }
    //            return cell.NumericCellValue.ToString();
    //        }
    //        else
    //        {
    //            cell.SetCellType(CellType.String);
    //            return cell.StringCellValue;
    //        }
    //    }
    //}
    //#endregion
}
