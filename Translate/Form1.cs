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

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.MaximizeBox = false;
            this.comboBox1.SelectedIndex = 0;
            this.comboBox2.SelectedIndex = 0;
            this.comboBox3.SelectedIndex = 0;
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
                MessageBox.Show("请补全信息！", "Translation Software");
                return;
            }
            if (!IsAuthorised("ww-0010"))
            {
                MessageBox.Show("调用接口异常！", "Translation Software");
                return;
            }
            ToTranslate(cb1Text);
        }
        public void ToTranslate(string cbText)
        {
            switch (cbText)
            {
                case "百度翻译":
                    {
                        Baidu baidu = new Baidu(sourceLanguage, targetLanguage, folderPath,label6);
                        baidu.TranslateThread();
                        break;
                    }
                case "搜狗翻译":
                    {
                        SouGou sg = new SouGou(sourceLanguage, targetLanguage);
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
            sign = sbTextArr?[1];
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
