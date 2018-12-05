using Aspose.Cells;
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace Translate.Helper
{
    public class Tool
    {
        public string basePath = System.AppDomain.CurrentDomain.BaseDirectory;
        public string resulePath = System.AppDomain.CurrentDomain.BaseDirectory + @"TranslaterResult\";

        /// <summary>  
        /// 获取当前时间戳  
        /// </summary>  
        /// <param name="bflag">为真时获取10位时间戳,为假时获取13位时间戳.bool bflag = true</param>  
        /// <returns></returns>  
        public string DateTimeToUnix(bool bflag = true)
        {
            TimeSpan ts = DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, 0);
            string ret = string.Empty;
            if (bflag)
                ret = Convert.ToInt64(ts.TotalSeconds).ToString();//10位的时间戳
            else
                ret = Convert.ToInt64(ts.TotalMilliseconds).ToString();//13位的时间戳

            return ret;
        }
        /// <summary>
        /// uri编码
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public string UriEncoe(string str)
        {
            return System.Web.HttpUtility.UrlEncode(str, System.Text.Encoding.UTF8);
        }
        /// <summary>
        /// Base64加密
        /// </summary>
        /// <param name="codeName">加密采用的编码方式</param>
        /// <param name="source">待加密的明文</param>
        /// <returns></returns>
        public string EncodeBase64(Encoding encode, string source)
        {
            byte[] bytes = encode.GetBytes(source);
            try
            {
                source = Convert.ToBase64String(bytes);
            }
            catch
            {
                source = source;
            }
            return source;
        }
        /// <summary>
        /// 基于HMACSHA256的 Base64加密
        /// </summary>
        /// <param name="message"></param>
        /// <param name="secret"></param>
        /// <returns></returns>
        public string EncodeHMACSHA256(string secret, string message)
        {
            secret = secret ?? "";
            var encoding = new System.Text.ASCIIEncoding();
            byte[] keyByte = encoding.GetBytes(secret);
            byte[] messageBytes = encoding.GetBytes(message);
            using (var hmacsha256 = new HMACSHA256(keyByte))
            {
                byte[] hashmessage = hmacsha256.ComputeHash(messageBytes);
                return Convert.ToBase64String(hashmessage);
            }
        }
        /// <summary>
        /// 32位MD5加密
        /// </summary>
        /// <param name="password"></param>
        /// <returns></returns>
        public string MD5Encrypt32(string str)
        {
            if (str == null)
            {
                return null;
            }
            MD5 md5Hash = MD5.Create();
            //将输入字符串转换为字节数组并计算哈希数据  
            byte[] data = md5Hash.ComputeHash(Encoding.UTF8.GetBytes(str));
            //创建一个 Stringbuilder 来收集字节并创建字符串  
            StringBuilder sBuilder = new StringBuilder();
            //循环遍历哈希数据的每一个字节并格式化为十六进制字符串  
            for (int i = 0; i < data.Length; i++)
            {
                sBuilder.Append(data[i].ToString("x2"));
            }
            //返回十六进制字符串  
            return sBuilder.ToString();
        }
        /// <summary>
        /// 将字符转换成utf8编码
        /// </summary>
        /// <param name="oldStr"></param>
        /// <returns></returns>
        public string GetUtf8String(string oldStr)
        {
            UTF8Encoding utf8 = new UTF8Encoding();
            Byte[] encodedBytes = utf8.GetBytes(oldStr);
            String decodedString = utf8.GetString(encodedBytes);
            return decodedString;
        }
        /// <summary>
        /// 读取所有txt
        /// </summary>
        /// <param name="folder"></param>
        /// <returns></returns>
        public Dictionary<string, string> ReadAllFile(string folder)
        {
            Dictionary<string, string> textDic = new Dictionary<string, string>();
            DirectoryInfo root = new DirectoryInfo(folder);
            string ext = string.Empty;
            foreach (FileInfo f in root.GetFiles())
            {
                ext = f.Extension.ToLower();
                if (ext == ".txt" || ext == ".doc" || ext == ".docx" || ext == ".xls" || ext == ".xlsx")
                {
                    string fullName = f.FullName;
                    if (!fullName.Contains("~$"))
                        textDic.Add(f.Name, f.FullName);
                }
            }
            return textDic;
        }
        /// <summary>
        /// 读取txt内容
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public string ReadTextContent(string filePath)
        {
            string content = string.Empty;
            if (File.Exists(filePath))
                content = File.ReadAllText(filePath, Encoding.Default);
            return GetUtf8String(content);
        }
        /// <summary>
        /// 写入文件
        /// </summary>
        /// <param name="content"></param>
        /// <returns></returns>
        public bool WriteTxt(string fileName, string content)
        {
            bool isSuccess = false;
            try
            {
                CheckFolderPath();
                string fileFullPath = resulePath + fileName;
                if (File.Exists(fileFullPath))
                    File.Delete(fileFullPath);
                File.WriteAllText(fileFullPath, content, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                WriteLog(ex);
            }
            return isSuccess;
        }
        /// <summary>
        /// 写入word
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="content"></param>
        /// <returns></returns>
        public bool WriteWord(string fileName, string content)
        {
            bool isSuccess = false;
            try
            {
                CheckFolderPath();
                string fileFullPath = resulePath + fileName;
                if (File.Exists(fileFullPath))
                    File.Delete(fileFullPath);
                Document doc = new Document();
                DocumentBuilder builder = new DocumentBuilder(doc);
                // Specify font formatting
                //Aspose.Words.Font font = builder.Font;
                //font.Size = 16;
                //font.Bold = true;
                //font.Color = System.Drawing.Color.Blue;
                //font.Name = "Arial";
                //font.Underline = Underline.Dash;
                // Specify paragraph formatting
                ParagraphFormat paragraphFormat = builder.ParagraphFormat;
                paragraphFormat.FirstLineIndent = 8;
                paragraphFormat.Alignment = ParagraphAlignment.Justify;
                paragraphFormat.KeepTogether = true;
                builder.Writeln(content);
                doc.Save(fileFullPath);
            }
            catch (Exception ex)
            {
                WriteLog(ex);
            }
            return isSuccess;
        }

        public bool WriteExcel(string fileName, Workbook workbook)
        {
            bool isSuccess = false;
            try
            {
                CheckFolderPath();
                string fileFullPath = resulePath + fileName;
                if (File.Exists(fileFullPath))
                    File.Delete(fileFullPath);
                workbook.Save(fileFullPath);
            }
            catch (Exception ex)
            {
                WriteLog(ex);
            }
            return isSuccess;
        }
        /// <summary>
        /// 判断文件目录是否存在，不存在则创建
        /// </summary>
        public void CheckFolderPath()
        {
            if (!Directory.Exists(resulePath))
                Directory.CreateDirectory(resulePath);
        }
        /// <summary>
        /// 打印日志
        /// </summary>
        /// <param name="log"></param>
        public void WriteLog(object logObj)
        {
            try
            {
                //string log = logObj.ToString();
                //string path = AppDomain.CurrentDomain.BaseDirectory + "log\\";//日志文件夹
                //DirectoryInfo dir = new DirectoryInfo(path);
                //if (!dir.Exists)//判断文件夹是否存在
                //    dir.Create();//不存在则创建

                //FileInfo[] subFiles = dir.GetFiles();//获取该文件夹下的所有文件
                //foreach (FileInfo f in subFiles)
                //{
                //    string fname = Path.GetFileNameWithoutExtension(f.FullName); //获取文件名，没有后缀a
                //    DateTime start = Convert.ToDateTime(fname);//文件名转换成时间
                //    DateTime end = Convert.ToDateTime(DateTime.Now.ToString("yyyy-MM-dd"));//获取当前日期
                //    TimeSpan sp = end.Subtract(start);//计算时间差
                //    if (sp.Days > 30)//大于30天删除
                //        f.Delete();
                //}

                //string logName = DateTime.Now.ToString("yyyy-MM-dd") + ".log";//日志文件名称，按照当天的日期命名
                //string fullPath = path + logName;//日志文件的完整路径
                //string contents = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " -> " + log + "\r\n";//日志内容

                //File.AppendAllText(fullPath, contents, Encoding.UTF8);//追加日志
            }
            catch (Exception)
            {
                throw;
            }
        }
        /// <summary>
        /// 中文繁简互转
        /// </summary>
        /// <param name="x"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public string StringConvert(string x, int type = 2)
        {
            String value = String.Empty;
            switch (type)
            {
                case 1://转繁体
                    value = Microsoft.VisualBasic.Strings.StrConv(x, Microsoft.VisualBasic.VbStrConv.TraditionalChinese, 0);
                    break;
                case 2://转简体
                    value = Microsoft.VisualBasic.Strings.StrConv(x, Microsoft.VisualBasic.VbStrConv.SimplifiedChinese, 0);
                    break;
                default:
                    break;
            }
            return value;
        }
    }
}
