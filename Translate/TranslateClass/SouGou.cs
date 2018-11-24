using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Translate.Helper;

namespace Translate.TranslateClass
{
    public class SouGou
    {
        public HttpHelper hh = null;
        public Tool tool = null;
        private string sourceLanguage = string.Empty;
        private string targetLanguage = string.Empty;
        private string souGouApi = "http://api.ai.sogou.com/pub/nlp/fanyi";
        private string ApiKey = "z9-KEwCqdindo61nAF3KfgEb";
        private string SecretKey = "5SRtgahjhAXHNcVeMldvuEpQR5ljhNqu";
        private string AppId = "15283";
        public Dictionary<string, string> textDic = new Dictionary<string, string>();
        public string folderPath = string.Empty;
        public SouGou(string s, string t)
        {
            this.sourceLanguage = s;
            this.targetLanguage = t;
            hh = new HttpHelper();
            tool = new Tool();
        }
        /// <summary>
        /// 开启线程
        /// </summary>
        /// <returns></returns>
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
            foreach (var item in textDic)
            {
                try
                {
                    string result = string.Empty;
                    string q = tool.ReadTextContent(item.Value);
                    string from = sourceLanguage;
                    string to = targetLanguage;
                    string charset = "utf-8";
                    string callback = string.Empty;
                    string secondsSinceEpoch = tool.DateTimeToUnix();
                    string expirationPeriodInSeconds = "3600";
                    string postData = $"q={q}&from={from}&to={to}&charset={charset}&callback={callback}";
                    string AuthPrefix = $"sac-auth-v1/{ApiKey}/{secondsSinceEpoch}/{expirationPeriodInSeconds}";

                    string REQUEST_METHOD = "POST";
                    string HOST = "api.ai.sogou.com";
                    string URI = "/pub/nlp/fanyi";
                    string SORTED_QUERY_STRING = SouGouEncrypt(postData);
                    string Data = REQUEST_METHOD + "\n" + HOST + "\n" + URI + "\n" + SORTED_QUERY_STRING;

                    string Signature = tool.EncodeHMACSHA256(SecretKey, AuthPrefix + "\n" + Data);
                    string Authorization = $"{AuthPrefix}/{Signature}";
                    //sac-auth-v1/bTkALtTB9x6GAxmFi9wetAGH/1491810516/3600/vuVEkzcnUeFv8FxeWS50c7S0HaYH1QKgtIV5xrxDY/s=
                    //sac-auth-v1/z9-KEwCqdindo61nAF3KfgEb/1542877943/3600/ZnkGZ8MK75CispY1bfPzai3DEUS8cjmCZRTaEkSxQDY=
                    Dictionary<string, string> headerDic = new Dictionary<string, string>();
                    headerDic.Add("Authorization", tool.EncodeBase64(Encoding.UTF8, Authorization));

                    result = hh.HttpPost(souGouApi, postData, headerDic);
                }
                catch (Exception ex)
                {
                    tool.WriteLog(ex);
                }
            }
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
    }
}
