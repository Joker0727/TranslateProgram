using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Translate.Helper;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Windows.Forms;
using System.Threading;
using Aspose.Words;
using Aspose.Cells;

namespace Translate.TranslateClass
{
    public class Baidu
    {
        private string baiduApi = "http://api.fanyi.baidu.com/api/trans/vip/translate";
        public HttpHelper hh = null;
        private string sourceLanguage = string.Empty;
        private string targetLanguage = string.Empty;
        public Tool tool = null;
        private string AppId = "20180414000146313";
        private string SecretKey = "Puqmy1cQTfXuCtXdsqci";
        public Dictionary<string, string> textDic = new Dictionary<string, string>();
        public string folderPath = string.Empty;
        public Label label6 = null;
        public int totalCount = 0;

        public Baidu(string s, string t, string fp, Label label6)
        {
            this.sourceLanguage = s;
            this.targetLanguage = t;
            this.folderPath = fp;
            this.label6 = label6;
            hh = new HttpHelper();
            tool = new Tool();
            textDic = tool.ReadAllFile(fp);
            totalCount = textDic.Count();
        }
        /// <summary>
        /// 开启线程
        /// </summary>
        public void TranslateThread()
        {
            Thread th = new Thread(ToTranslate);
            th.IsBackground = true;
            th.Start();
        }
        /// <summary>
        /// 翻译文章
        /// </summary>
        public void ToTranslate()
        {
            int index = 0;
            foreach (var item in textDic)
            {
                try
                {
                    if (item.Value.Contains(".txt"))
                        DealTxt(item);
                    else if (item.Value.Contains(".doc") || item.Value.Contains(".docx"))
                        DealWord(item);
                    else if (item.Value.Contains(".xls") || item.Value.Contains(".xlsx"))
                        DealExcel(item);
                    index++;
                    this.label6.Invoke(new Action(() =>
                    {
                        this.label6.Text = index + "/" + totalCount;
                    }));
                }
                catch (Exception ex)
                {
                    tool.WriteLog(ex);
                }
            }
            MessageBox.Show($"{totalCount} 片文章翻译完毕！", "Translation Software");
        }

        public void DealTxt(KeyValuePair<string, string> item)
        {
            string content = tool.ReadTextContent(item.Value);
            string newText = Translate(content);
            tool.WriteTxt(item.Key, newText);
        }
        public void DealWord(KeyValuePair<string, string> item)
        {
            Document doc = new Document(item.Value);
            ParagraphCollection pCollection = null;
            if (doc.FirstSection.Body.Paragraphs.Count > 0)
                pCollection = doc.FirstSection.Body.Paragraphs;//word中的所有段落
            string text = string.Empty;
            foreach (Paragraph pc in pCollection)
            {
                text += pc.GetText() + "\r\n";
            }
            string newText = Translate(text);
            tool.WriteWord(item.Key, newText);
        }

        public void DealExcel(KeyValuePair<string, string> item)
        {
            Workbook workbook = new Workbook(item.Value);
            Cells cells = workbook.Worksheets[0].Cells;
            string cellText = string.Empty, newText = string.Empty;
            for (int i = 0; i < cells.MaxDataRow + 1; i++)
            {
                for (int j = 0; j < cells.MaxDataColumn + 1; j++)
                {
                    cellText = cells[i, j].StringValue.Trim();
                    if (string.IsNullOrEmpty(cellText))
                        break;
                    newText = Translate(cellText);
                    cells[i, j + 2].PutValue(newText);
                    cells[i, j + 2].SetStyle(cells[i, j].GetStyle());
                    break;
                }
            }
            tool.WriteExcel(item.Key, workbook);
        }
        public string Translate(string item)
        {
            string newContent = string.Empty;
            try
            {
                string q = item;
                string from = sourceLanguage;
                string to = targetLanguage;
                string salt = tool.DateTimeToUnix();
                string signStr = AppId + q + salt + SecretKey;
                string sign = tool.MD5Encrypt32(signStr);

                string query = $"q={q}&from={from}&to={to}&appid={AppId}&salt={salt}&sign={sign}";
                string queryUrl = baiduApi + "?" + query;
                string result = hh.HttpPost(baiduApi, query);
                var obj = JObject.Parse(result);

                foreach (var trans_result in obj["trans_result"])
                    newContent += trans_result?["dst"].ToString() + "\r\n";
            }
            catch (Exception ex)
            {
                tool.WriteLog(ex.ToString());
            }
            return newContent;
        }
    }
}
