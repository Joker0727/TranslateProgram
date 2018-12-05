using Aspose.Cells;
using Aspose.Words;
using Newtonsoft.Json.Linq;
using RestSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Translate.Helper;
using ParameterType = RestSharp.ParameterType;

namespace Translate.TranslateClass
{
    public class SouGou
    {
        public HttpHelper hh = null;
        public Tool tool = null;
        private string sourceLanguage = string.Empty;
        private string targetLanguage = string.Empty;
        private string souGouApi = "http://fanyi.sogou.com/reventondc/api/sogouTranslate";
        private string PID = "bcd9105a9a9a2ae0629e5e2bfb17ba89";
        private string Key = "323d57a3f015b8bf0bca10b8abdc63fd";
        private string sign = string.Empty;
        public Dictionary<string, string> textDic = new Dictionary<string, string>();
        public string folderPath = string.Empty;
        public Label label6 = null;
        public int totalCount = 0;

        public SouGou(string s, string t, string fp, Label label6)
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
        /// <returns></returns>
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
            string q = string.Empty, salt = string.Empty, signStr = string.Empty;
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
        /// <summary>
        /// 实现搜狗的加密规则
        /// </summary>
        /// <param name="pd">？问号后的参数字符串</param>
        /// <returns></returns>
        public string SouGouEncrypt(string pd)
        {
            string result = string.Empty;
            string[] pdArr = pd.Split('&');
            List<string> kvList = new List<string>();
            foreach (var item in pdArr)
            {
                string[] kvArr = item.Split('=');
                string key = tool.UriEncoe(kvArr[0]);
                string value = tool.UriEncoe(kvArr[1]);
                kvList.Add(key + "=" + value);
            }
            kvList.Sort();
            foreach (var kv in kvList)
            {
                result += kv + "&";
            }
            return result.Trim('&');
        }
        public void DealTxt(KeyValuePair<string, string> item)
        {
            if (!File.Exists(item.Value))
            {
                MessageBox.Show($"【{item.Value}】文件不存在", "Translation Software");
                return;
            }
            string text = tool.ReadTextContent(item.Value);
            string newText = Translate(text);
            tool.WriteTxt(item.Key, newText.Replace("\n", "\r\n"));
        }
        public void DealWord(KeyValuePair<string, string> item)
        {
            string newContent = string.Empty, text = string.Empty;
            Document doc = new Document(item.Value);
            ParagraphCollection pCollection = null;
            if (doc.FirstSection.Body.Paragraphs.Count > 0)
                pCollection = doc.FirstSection.Body.Paragraphs;//word中的所有段落
            foreach (Paragraph pc in pCollection)
            {
                text = tool.GetUtf8String(tool.StringConvert(pc.GetText())).Replace("\r", "").Replace("\f", "");
                newContent += Translate(text) + "\r\n";
            }
            tool.WriteWord(item.Key, newContent);
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
            string q = string.Empty, salt = string.Empty, signStr = string.Empty, newText = string.Empty;
            try
            {
                q = item;
                salt = tool.DateTimeToUnix();
                signStr = PID.Trim(' ') + q.Trim(' ') + salt.Trim(' ') + Key.Trim(' ');
                sign = tool.MD5Encrypt32(signStr);

                var client = new RestClient(souGouApi);
                var request = new RestRequest(Method.POST);
                request.AddHeader("cache-control", "no-cache");
                request.AddHeader("accept", "application/json");
                request.AddHeader("content-type", "application/x-www-form-urlencoded");
                string postData = $"from={sourceLanguage}&pid={PID}&q={q}&sign={sign}&salt={salt}&to={targetLanguage}";
                request.AddParameter("application/x-www-form-urlencoded", postData, ParameterType.RequestBody);
                IRestResponse response = client.Execute(request);
                string jsonStr = response.Content;
                var jsonObj = JObject.Parse(jsonStr);
                newText = jsonObj["translation"]?.ToString();
            }
            catch (Exception ex)
            {
                tool.WriteLog(ex);
            }
            return newText;
        }
    }
}
