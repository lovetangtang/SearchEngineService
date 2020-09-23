using Lucene.Net.Analysis;
using Lucene.Net.Analysis.Standard;
using Lucene.Net.Documents;
using Lucene.Net.Index;
using Lucene.Net.QueryParsers;
using Lucene.Net.Search;
using Lucene.Net.Store;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using SES.Utility;
using PanGu.HighLight;
using Aspose.Pdf.Generator;

namespace SES.Web.Models
{
    public class LuceneEntity
    {
        /// <summary>
        /// 索引保存路径
        /// </summary>
        public static string indexPath;
        static LuceneEntity() {
            indexPath = HttpContext.Current.Server.MapPath("/IndexPath");
        }


        /// <summary>
        /// 提交内容
        /// </summary>
        /// <param name="id"></param>
        /// <param name="title"></param>
        /// <param name="content"></param>
        /// <param name="date"></param>
        /// <param name="param"></param>
        /// <param name="modcode"></param>
        /// <returns></returns>
        public static string CommitContent(string id, string title, string content, DateTime? date, string param, string modcode, string searchparam1, string searchparam2, string searchparam3)
        {
            #region 验证输入参数是否合格
            if (string.IsNullOrEmpty(id))
            {
                throw new Exception("参数id不能为空");
            }
            if (string.IsNullOrEmpty(title))
            {
                throw new Exception("参数title不能为空");
            }
            if (string.IsNullOrEmpty(content))
            {
                throw new Exception("参数content和filepathList不能同时为空");
            } 
            #endregion
            if (date == null) date = DateTime.Now;//日期为null时使用当前时间
            
            //Analyzer analyzer = new StandardAnalyzer(Lucene.Net.Util.Version.LUCENE_29);
            Analyzer analyzer =new PanGuAnalyzer();
            FSDirectory directory = FSDirectory.Open(new DirectoryInfo(indexPath), new NativeFSLockFactory());
            bool isUpdate = IndexReader.IndexExists(directory);
            if (isUpdate)
            {
                //如果索引目录被锁定（比如索引过程中程序异常退出），则首先解锁
                if (IndexWriter.IsLocked(directory))
                {
                    IndexWriter.Unlock(directory);
                }
            }
            IndexWriter writer = new IndexWriter(directory, analyzer, !isUpdate, IndexWriter.MaxFieldLength.UNLIMITED);
            //为避免重复索引，所以先删除id=id的记录，再重新添加
            writer.DeleteDocuments(new Term("id", id));

            string result = "";
            try
            {
                Document doc = new Document();
                doc.Add(new Field("id", id, Field.Store.YES, Field.Index.NOT_ANALYZED));//存储
                doc.Add(new Field("title", title, Field.Store.YES, Field.Index.ANALYZED, Field.TermVector.WITH_POSITIONS_OFFSETS));//分词建立索引
                doc.Add(new Field("content", content, Field.Store.YES, Field.Index.ANALYZED, Field.TermVector.WITH_POSITIONS_OFFSETS));//分词建立索引
                doc.Add(new Field("date", date.ToString(), Field.Store.YES, Field.Index.NOT_ANALYZED));//不分词建立索引
                doc.Add(new Field("param", param, Field.Store.YES, Field.Index.NO));//存储   
                doc.Add(new Field("modcode", modcode, Field.Store.YES, Field.Index.NOT_ANALYZED));//存储   

                doc.Add(new Field("searchparam1", searchparam1, Field.Store.YES, Field.Index.NOT_ANALYZED));//存储   
                doc.Add(new Field("searchparam2", searchparam2, Field.Store.YES, Field.Index.NOT_ANALYZED));//存储   
                doc.Add(new Field("searchparam3", searchparam3, Field.Store.YES, Field.Index.NOT_ANALYZED));//存储   
                
                writer.AddDocument(doc);

                //输出api调用记录
                string EnableApiLog = Util.GetAppSetting("EnableApiLog");
                if (EnableApiLog == "1") {
                    string logMsg = GetParamList(id, title, content, date, param,modcode) + "\r\n" + Utility.Util.GetClientInfo();
                    Utility.Util.WriteApiLog(logMsg);
                }                
            }
            catch (Exception ex)
            {
                result= ex.Message;
                string errMsg = ex.Message + "\r\n" + GetParamList(id, title, content, date, param, modcode) + "\r\n" + Utility.Util.GetClientInfo();
                Utility.Util.WriteLog(errMsg);
            }

            //对索引文件进行优化
            //writer.Optimize();
            analyzer.Close();
            writer.Dispose();
            directory.Dispose();
            return result;
        }


        /// <summary>
        /// 搜索内容
        /// </summary>
        /// <param name="word">搜索关键字</param>
        /// <param name="pagesize">每页显示记录数</param>
        /// <param name="pageindex">当前页码</param>
        /// <returns></returns>
        public static SearchResult SearchContent(string modcode, string word, int pagesize, int pageindex, string searchparam1, string searchparam2, string searchparam3)
        {
            SearchResult searchResult = new SearchResult();

            FSDirectory directory = FSDirectory.Open(new DirectoryInfo(indexPath), new NativeFSLockFactory());
            IndexSearcher searcher = new IndexSearcher(directory, true);
            var analyzer = new PanGuAnalyzer();
            //初始化MultiFieldQueryParser以便同时查询多列 
            Lucene.Net.QueryParsers.MultiFieldQueryParser parser = new Lucene.Net.QueryParsers.MultiFieldQueryParser(Lucene.Net.Util.Version.LUCENE_29,new string[] { "title", "content" },analyzer);
            Lucene.Net.Search.Query query = parser.Parse(word);//初始化Query 
            parser.DefaultOperator=QueryParser.AND_OPERATOR;

            Lucene.Net.Search.BooleanQuery boolQuery = new Lucene.Net.Search.BooleanQuery();
            boolQuery.Add(query, Occur.MUST);
            if (!string.IsNullOrEmpty(modcode)) {
                PhraseQuery queryModCode = new PhraseQuery();
                queryModCode.Add(new Term("modcode", modcode));
                boolQuery.Add(queryModCode, Occur.MUST);
            }

            if (!string.IsNullOrEmpty(searchparam1)) {
                WildcardQuery query1 = new WildcardQuery(new Term("searchparam1","*"+searchparam1+"*"));
                boolQuery.Add(query1, Occur.MUST);
            }
            if (!string.IsNullOrEmpty(searchparam2))
            {
                WildcardQuery query1 = new WildcardQuery(new Term("searchparam2", "*" + searchparam2 + "*"));
                boolQuery.Add(query1, Occur.MUST);
            }
            if (!string.IsNullOrEmpty(searchparam3))
            {
                WildcardQuery query1 = new WildcardQuery(new Term("searchparam3", "*" + searchparam3 + "*"));
                boolQuery.Add(query1, Occur.MUST);
            }

            Sort sort = new Sort(new SortField("date", SortField.STRING,true));
            var result = searcher.Search(boolQuery, null, 1000, sort);
            if (result.TotalHits == 0)
            {
                searchResult.count = 0;
            }
            else
            {
                searchResult.count = result.TotalHits;
                int startNum = 0, endNum = result.TotalHits;
                if (pagesize > 0) {
                    //当pagesize>0时使用分页功能
                    startNum = (pageindex-1) * pagesize;
                    endNum = startNum+pagesize;
                }
                ScoreDoc[] docs = result.ScoreDocs;
                List<JObject> dataList = new List<JObject>();
                for (int i = 0; i < docs.Length; i++)
                {
                    if (i < startNum) { continue; }
                    if (i >= endNum) { break; }

                    Document doc = searcher.Doc(docs[i].Doc);
                    string id = doc.Get("id").ToString();
                    string title = doc.Get("title").ToString();
                    string content = doc.Get("content").ToString();
                    string date = doc.Get("date").ToString();
                    string param = doc.Get("param").ToString();
                    string mcode = doc.Get("modcode").ToString();
                    string param1 = doc.Get("searchparam1").ToString();
                    string param2 = doc.Get("searchparam2").ToString();
                    string param3 = doc.Get("searchparam3").ToString();
                    JObject obj = new JObject();
                    obj["id"] = id;

                    //创建HTMLFormatter,参数为高亮单词的前后缀 
                    string highLightTag = Util.GetAppSetting("HighLightTag", "<font color=\"red\">|</font>");
                    string[] tarArr = highLightTag.Split('|');
                    var simpleHTMLFormatter = new SimpleHTMLFormatter(tarArr[0],tarArr[1]);
                    //创建 Highlighter ，输入HTMLFormatter 和 盘古分词对象Semgent 
                    var highlighter = new Highlighter(simpleHTMLFormatter, new PanGu.Segment());
                    //设置每个摘要段的字符数 
                    int highlightFragmentSize = Util.GetAppSetting("HighlightFragmentSize", "100").ToInt();
                    highlighter.FragmentSize = highlightFragmentSize;
                    //获取最匹配的摘要段 
                    String bodyPreview = highlighter.GetBestFragment(word, content);
                    string newTitle = highlighter.GetBestFragment(word, title);
                    if (!string.IsNullOrEmpty(newTitle)) title = newTitle;

                    obj["title"] = title;
                    obj["content"] = bodyPreview;
                    obj["date"] = date;
                    obj["param"] = param;
                    obj["modcode"] = mcode;
                    obj["searchparam1"] = param1;
                    obj["searchparam2"] = param2;
                    obj["searchparam3"] = param3;
                    dataList.Add(obj);
                }
                searchResult.data = dataList;
            }
            analyzer.Close();
            searcher.Dispose();
            directory.Dispose();

            return searchResult;
        }

        public static string GetParamList(string id, string title, string content, DateTime? date, string param, string modcode)
        {
            if (content.Length > 500) content = content.Substring(0, 500)+"...";
            return "{id:" + id + ",title:" + title + ",content:" + content + ",date:" + date + ",param:" + param + ",modcode:" + modcode + "}";
        }

        public static string Delete(string id)
        {
            #region 验证输入参数是否合格
            if (string.IsNullOrEmpty(id))
            {
                throw new Exception("参数id不能为空");
            }
            #endregion





            //Analyzer analyzer = new PanGuAnalyzer();
            //FSDirectory directory = FSDirectory.Open(new DirectoryInfo(indexPath), new NativeFSLockFactory());
            //bool isUpdate = IndexReader.IndexExists(directory);
            //if (isUpdate)
            //{
            //    //如果索引目录被锁定（比如索引过程中程序异常退出），则首先解锁
            //    if (IndexWriter.IsLocked(directory))
            //    {
            //        IndexWriter.Unlock(directory);
            //    }
            //}
            //IndexWriter writer = new IndexWriter(directory, analyzer, !isUpdate, IndexWriter.MaxFieldLength.UNLIMITED);
            ////为避免重复索引，所以先删除id=id的记录，再重新添加
            //writer.DeleteDocuments(new Term("id", id));
            ////对索引文件进行优化
            //writer.Optimize();
            //analyzer.Close();
            //writer.Dispose();
            //directory.Dispose();


            Analyzer analyzer = new PanGuAnalyzer();
            FSDirectory directory = FSDirectory.Open(new DirectoryInfo(indexPath), new NativeFSLockFactory());
            /**创建 索引写对象
             * 用于正式 写入索引与文档数据、删除索引与文档数据
             * */
            IndexWriter indexWriter = new IndexWriter(directory, analyzer, Lucene.Net.Index.IndexWriter.MaxFieldLength.LIMITED);

            /** 删除所有索引
             * 如果索引库中的索引已经被删除，则重复删除时无效*/
            Term term = new Term("id", id);
            TermQuery query = new TermQuery(term);
            indexWriter.DeleteDocuments(query);
            /** 虽然不 commit，也会生效，但建议做提交操作，*/
            indexWriter.Commit();
            /**  关闭流，里面会自动 flush*/
            indexWriter.Optimize();
            indexWriter.Dispose();

            return "";
        }

        public static string DeleteAll(){
            ///** 创建 IKAnalyzer 中文分词器
            // * IKAnalyzer()：默认使用最细粒度切分算法
            // * IKAnalyzer(boolean useSmart)：当为true时，分词器采用智能切分 ；当为false时，分词器迚行最细粒度切分
            // * */
            //Analyzer analyzer = new IKAnalyzer();
            ///** 指定索引和文档存储的目录
            // * 如果此目录不是 Lucene 的索引目录，则不进行任何操作*/
            //Directory directory = FSDirectory.open(indexDir);
 
            ///** 创建 索引写配置对象，传入分词器
            // * Lucene 7.4.0 版本 IndexWriterConfig 构造器不需要指定 Version.LUCENE_4_10_3
            // * Lucene 4.10.3 版本 IndexWriterConfig 构造器需要指定 Version.LUCENE_4_10_3
            // * */
            //IndexWriterConfig config = new IndexWriterConfig(Version.LUCENE_4_10_3, analyzer);
 
            ///**创建 索引写对象
            // * 用于正式 写入索引与文档数据、删除索引与文档数据
            // * */
            //IndexWriter indexWriter = new IndexWriter(directory, config);
 
            ///** 删除所有索引
            // * 如果索引库中的索引已经被删除，则重复删除时无效*/
            //indexWriter.deleteAll();
 
            ///** 虽然不 commit，也会生效，但建议做提交操作，*/
            //indexWriter.commit();
            ///**  关闭流，里面会自动 flush*/
            //indexWriter.close();


            
            Analyzer analyzer = new PanGuAnalyzer();
            FSDirectory directory = FSDirectory.Open(new DirectoryInfo(indexPath), new NativeFSLockFactory());
            /**创建 索引写对象
             * 用于正式 写入索引与文档数据、删除索引与文档数据
             * */
            IndexWriter indexWriter = new IndexWriter(directory,analyzer,Lucene.Net.Index.IndexWriter.MaxFieldLength.LIMITED);
 
            /** 删除所有索引
             * 如果索引库中的索引已经被删除，则重复删除时无效*/
            indexWriter.DeleteAll();
 
            /** 虽然不 commit，也会生效，但建议做提交操作，*/
            indexWriter.Commit();
            /**  关闭流，里面会自动 flush*/
            indexWriter.Dispose();

            return "";
        }
    }
}