using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

using SES.Utility;
using SES.Web.Models;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text;

namespace SES.Web
{
    /// <summary>
    /// api 的摘要说明
    /// </summary>
    public class api : IHttpHandler
    {

        public void ProcessRequest(HttpContext context)
        {
            context.Response.ContentType = "application/json";
            ValidateSecret();
            string action = HttpContext.Current.Request.QueryString["action"];
            if (action == "Commit")
            {
                Commit();
            }
            else if (action == "Search") {
                Search();
            }
            else if (action == "Delete") {
                Delete();
            }
            else if (action == "DeleteAll") {
                DeleteAll();
            }
            else
            {
                var result = new JsonResult() { errcode = 500, errmsg = "不支持此功能" };
                HttpContext.Current.Response.Write(result.ToJson());
                HttpContext.Current.Response.End();
            }
        }

        private void DeleteAll()
        {
            try
            {
                string resultMsg = LuceneEntity.DeleteAll();
                if (resultMsg != "")
                {
                    Util.ResponseJsonResult(500, resultMsg);
                }
                else
                {
                    Util.ResponseJsonResult(0, "删除成功");
                }
            }
            catch (Exception ex)
            {
                Util.ResponseJsonResult(500, ex.Message);
            }
        }

        private void Delete()
        {
            string id = Util.GetParamsString("id");
            try
            {
                string resultMsg = LuceneEntity.Delete(id);
                if (resultMsg != "")
                {
                    Util.ResponseJsonResult(500, resultMsg);
                }
                else
                {
                    Util.ResponseJsonResult(0, "删除成功");
                }
            }
            catch (Exception ex)
            {
                Util.ResponseJsonResult(500, ex.Message);
            }
        }

        /// <summary>
        /// 搜索内容
        /// </summary>
        private void Search()
        {
            string word = Util.GetParamsString("word");
            string modcode = Util.GetParamsString("modcode");
            int pagesize = Util.GetParamsString("pagesize").ToInt();
            int pageindex = Util.GetParamsString("pageindex").ToInt();
            string searchparam1 = Util.GetParamsString("searchparam1");
            string searchparam2 = Util.GetParamsString("searchparam2");
            string searchparam3 = Util.GetParamsString("searchparam3");
            if (pageindex == 0) pageindex = 1;
            try
            {
                SearchResult searchResulg = LuceneEntity.SearchContent(modcode, word, pagesize, pageindex, searchparam1, searchparam2, searchparam3);
                Util.ResponseJsonResult(0, searchResulg);
            }
            catch (Exception ex)
            {
                Util.ResponseJsonResult(500,ex.Message);
            }            
        }

        /// <summary>
        /// 验证调用接口权限
        /// </summary>
        private void ValidateSecret()
        {
            string secret = Util.GetParamsString("secret");
            string appSecret = Util.GetAppSetting("Secret");
            if (secret != appSecret)
            {
                var result = new JsonResult() { errcode = 500, errmsg = "没有接口调用权限" };
                HttpContext.Current.Response.Write(result.ToJson());
                HttpContext.Current.Response.End();
            }
        }

        /// <summary>
        /// 提交搜索引擎内容
        /// </summary>
        private void Commit()
        {
            string id = Util.GetParamsString("id");
            string modcode = Util.GetParamsString("modcode");
            string title = Util.GetParamsString("title");
            string content = Util.GetParamsString("content");
            string filepathList = Util.GetParamsString("filepathList");//（文件全路径）多个文件路径之间用|分隔
            DateTime? date = Util.GetParamsString("date").ToDateOrNull();
            string param = Util.GetParamsString("param");
            string searchparam1 = Util.GetParamsString("searchparam1");
            string searchparam2 = Util.GetParamsString("searchparam2");
            string searchparam3 = Util.GetParamsString("searchparam3");

            content = GetContent(content, filepathList);

            try
            {
                string resultMsg = LuceneEntity.CommitContent(id, title, content, date, param, modcode, searchparam1, searchparam2, searchparam3);
                if (resultMsg != "")
                {
                    Util.ResponseJsonResult(500, resultMsg);
                }
                else
                {
                    Util.ResponseJsonResult(0, "提交成功");
                }
            }
            catch (Exception ex)
            {
                Util.ResponseJsonResult(500, ex.Message);
            }
            
        }

        private string GetContent(string content, string filepathList)
        {
            if (string.IsNullOrEmpty(filepathList)) return content;

            StringBuilder sb = new StringBuilder(content);
            string[] filepathArr = filepathList.Split('|');
            for (int i = 0; i < filepathArr.Length; i++) {
                string filepath = filepathArr[i];
                if (File.Exists(filepath))
                {
                    try
                    {
                        sb.Append(GetFileContent(filepath));
                    }
                    catch (Exception ex)
                    {
                        Util.WriteLog("读取“" + filepath + "”文件内容失败:"+ex.Message);
                    }
                    
                }
                else {
                    Util.WriteLog("“"+filepath + "”文件不存在");
                }
            }
            return sb.ToString();
        }

        public static string GetFileContent(string filepath)
        {
            string extension = Path.GetExtension(filepath).ToLower();
            string content = "";
            if (extension == ".doc" || extension == ".docx") {
                content=AsposeUtil.GetWordContent(filepath);
            }
            else if (extension == ".xls" || extension == ".xlsx")
            {
                content = AsposeUtil.GetExcelContent(filepath);
            }
            else if (extension == ".pdf")
            {
                content = AsposeUtil.GetPdfContent(filepath);
            }
            else if (extension == ".txt" || extension == ".rtf")
            {
                content = AsposeUtil.GetTxtContent(filepath);
            }
            else {
                Util.WriteLog("不能读取“" + filepath + "”文件内容");
            }
            return content;
        }

        public bool IsReusable
        {
            get
            {
                return false;
            }
        }
    }
}