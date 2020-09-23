using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SES.Utility
{
    public class AsposeUtil
    {
        private const int PDFMaxPage = 100;//PDF文件最多读取前100页
        private const int ExcelMaxCell = 2000;//Excel文件最多读取前2000个单元格


        public static string GetWordContent(string filepath)
        {
            Document doc = new Document(filepath);
            return doc.GetText();
        }

        public static string GetExcelContent(string filepath)
        {
            int numExcelMaxCell = Util.GetAppSetting("ExcelMaxCell", ExcelMaxCell.ToString()).ToInt();
            Aspose.Cells.Workbook book = new Aspose.Cells.Workbook(filepath);
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < book.Worksheets.Count; i++)
            {
                var cells = book.Worksheets[i].Cells;
                for (int j = 0; j < cells.Count; j++)
                {
                    if (j > numExcelMaxCell) break;
                    var cell = cells[j];
                    sb.Append(cell.StringValue + ",");
                }
            }
            return sb.ToString();
        }

        public static string GetPdfContent(string filepath)
        {
            int numPDFMaxPage = Util.GetAppSetting("PDFMaxPage", PDFMaxPage.ToString()).ToInt();
            Aspose.Pdf.Document doc = new Aspose.Pdf.Document(filepath);
            Aspose.Pdf.Text.TextAbsorber txt = new Aspose.Pdf.Text.TextAbsorber();
            StringBuilder sb = new StringBuilder();
            for (int i = 1; i <= doc.Pages.Count; i++)
            {
                if (i > numPDFMaxPage) break;

                doc.Pages[i].Accept(txt);
                sb.Append(txt.Text);
            }
            return sb.ToString();
        }

        public static string GetTxtContent(string filepath)
        {
            return File.ReadAllText(filepath);
        }
    }
}
