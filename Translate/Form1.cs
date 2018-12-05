using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Translate.TranslateClass;

namespace Translate
{
    public partial class Form1 : Form
    {
        private string sourceLanguage = string.Empty;
        private string targetLanguage = string.Empty;
        private string folderPath = string.Empty;
        private string[] baiduLanguage = { "zh：中文", "en：英语", "jp：日语", "kor：韩语", "fra：法语", "spa：西班牙语", "th：泰语", "ara：阿拉伯语", "ru：俄语", "pt：葡萄牙语", "de：德语", "it：意大利语", "el：希腊语", "vie：越南语", "nl：荷兰语", "pl：波兰语" };
        private string[] sougouuLanguage = { "zh-CHS：中文", "en：英语", "ja：日语", "ko：韩语", "fr：法语", "es：西班牙语", "th：泰语", "ar：阿拉伯语", "ru：俄语", "pt：葡萄牙语", "de：德语", "it：意大利语", "el：希腊语", "vi：越南语", "nl：荷兰语", "pl：波兰语" };
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            this.comboBox1.SelectedIndex = 0;
            this.comboBox2.SelectedIndex = 0;
            this.comboBox3.SelectedIndex = 1;
        }
        /// <summary>
        /// 选择文件夹
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textBox1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.ShowDialog();
            this.textBox1.Text = path.SelectedPath;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            folderPath = this.textBox1.Text;
            int cb1Index = this.comboBox1.SelectedIndex;

            string cb1Text = this.comboBox1.Text;
            sourceLanguage = GetLanguageSign(this.comboBox2.Text);
            targetLanguage = GetLanguageSign(this.comboBox3.Text);
            if (cb1Index < 0 || string.IsNullOrEmpty(folderPath))
            {
                MessageBox.Show("请选择待翻译文件所在目录！", "Translation Software");
                return;
            }
            if (!IsAuthorised("ww-0010"))
            {
                MessageBox.Show("调用接口异常！", "Translation Software");
                return;
            }
            ToTranslate(cb1Text);
        }
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combobox = sender as ComboBox;
            string cst = combobox.SelectedItem.ToString();
            this.comboBox2.Items.Clear();
            this.comboBox3.Items.Clear();
            switch (cst)
            {
                case "百度翻译":
                    {
                        this.comboBox2.Items.AddRange(baiduLanguage);
                        this.comboBox3.Items.AddRange(baiduLanguage);
                        break;
                    }
                case "搜狗翻译":
                    {
                        this.comboBox2.Items.AddRange(sougouuLanguage);
                        this.comboBox3.Items.AddRange(sougouuLanguage);
                        break;
                    }
                default:
                    break;
            }
            this.comboBox2.SelectedIndex = 0;
            this.comboBox3.SelectedIndex = 1;
        }
        public void ToTranslate(string cbText)
        {
            switch (cbText)
            {
                case "百度翻译":
                    {
                        Baidu baidu = new Baidu(sourceLanguage, targetLanguage, folderPath, label6);
                        baidu.TranslateThread();
                        break;
                    }
                case "搜狗翻译":
                    {
                        SouGou sg = new SouGou(sourceLanguage, targetLanguage, folderPath, label6);
                        sg.TranslateThread();
                        break;
                    }
                default:
                    {
                        MessageBox.Show("不支持的翻译！", "Translation Software");
                        break;
                    }
            }
        }
        /// <summary>
        /// 获取语言标志
        /// </summary>
        /// <param name="cbText"></param>
        /// <returns></returns>
        private string GetLanguageSign(string cbText)
        {
            string sign = string.Empty;
            string[] sbTextArr = cbText.Split('：');
            sign = sbTextArr?[0];
            return sign;
        }
        /// <summary>
        /// 授权
        /// </summary>
        /// <param name="workId"></param>
        /// <returns></returns>
        public bool IsAuthorised(string workId)
        {
            string conStr = "Server=111.230.149.80;DataBase=MyDB;uid=sa;pwd=1add1&one";
            using (SqlConnection con = new SqlConnection(conStr))
            {
                string sql = string.Format("select count(*) from MyWork Where WorkId ='{0}'", workId);
                using (SqlCommand cmd = new SqlCommand(sql, con))
                {
                    con.Open();
                    int count = int.Parse(cmd.ExecuteScalar().ToString());
                    if (count > 0)
                        return true;
                }
            }
            return false;
        }
    }
}
