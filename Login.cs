using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace JiShi_FileUpload
{
    public partial class Login : Form
    {
        private string name;
        private string id;
        public Login()
        {
            InitializeComponent();
        }
        public string NAME
        {
            get
            {
                return name;
            }
            set
            {
                name = value;
            }
        }
        public string ID
        {
            get
            {
                return id;
            }
            set
            {
                id = value;
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(textBox1.Text) && !string.IsNullOrEmpty(textBox2.Text))
            {
                OracleConnection conn = new OracleConnection(ConfigurationManager.AppSettings["strconn"]);
                conn.Open();
                string sql = "select * from Sys_User where NAME='" + textBox1.Text + "' and  PASSWORD='" + ToSHA1(textBox1.Text) + "'";
                OracleCommand com = new OracleCommand(sql, conn);
                OracleDataAdapter da = new OracleDataAdapter(com);
                DataSet ds = new DataSet();
                da.Fill(ds);
                conn.Close();
                if (ds.Tables[0].Rows.Count > 0)
                {
                    FileMonitor fm = new FileMonitor();
                    fm.Owner = this;
                    this.NAME = ds.Tables[0].Rows[0]["REALNAME"].ToString();
                    this.ID = ds.Tables[0].Rows[0]["ID"].ToString();
                    fm.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("用户名或密码错误！");
                }
            }
        }
        private string ToSHA1(string value)
        {
            string result = string.Empty;
            SHA1 sha1 = new SHA1CryptoServiceProvider();
            byte[] array = sha1.ComputeHash(Encoding.Unicode.GetBytes(value));
            for (int i = 0; i < array.Length; i++)
            {
                result += array[i].ToString("x2");
            }
            return result;
        }
    }
}
