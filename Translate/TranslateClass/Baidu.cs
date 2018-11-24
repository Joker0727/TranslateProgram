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
            textDic = tool.ReadAllText(fp);
            totalCount = textDic.Count();
        }
        /// <summary>
        /// 开启线程
        /// </summary>
        public void TranslateThread()
        {
            Thread th = new Thread(ToTranslateTxt);
            th.IsBackground = true;
            th.Start();
        }
        /// <summary>
        /// 翻译文章
        /// </summary>
        public void ToTranslateTxt()
        {
            int index = 0;
            foreach (var item in textDic)
            {
                try
                {
                    string q = tool.ReadTextContent(item.Value);
                    string from = sourceLanguage;
                    string to = targetLanguage;
                    string salt = tool.DateTimeToUnix();
                    string signStr = AppId + q + salt + SecretKey;
                    string sign = tool.MD5Encrypt32(signStr);

                    string query = $"q={q}&from={from}&to={to}&appid={AppId}&salt={salt}&sign={sign}";
                    string queryUrl = baiduApi + "?" + query;
                    string result = hh.HttpPost(baiduApi, query);
                    var obj = JObject.Parse(result);
                    string newContent = string.Empty;

                    foreach (var trans_result in obj["trans_result"])
                    {
                        newContent += trans_result?["dst"].ToString() + "\r\n";
                    }
                    tool.WriteTxt(item.Key, newContent);
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
    }
}
