using ReportMailer;
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

namespace DEPO_KİMYASAL_ERP
{
    public partial class Form1 : Form
    {
        private SqlConnection baglanti;
        public Form1()
        {
            InitializeComponent();
        }

        private void button11_Click(object sender, EventArgs e)
        {

            Posta.Default.SQL_Server = this.SQL_server.Text;
            Posta.Default.SQL_Database = this.SQL_database.Text;
            Posta.Default.SQL_User = this.SQL_Kullanici.Text;
            Posta.Default.SQL_Password = this.SQL_sifre.Text;
            Posta.Default.Save();
            MessageBox.Show("Server Ayarları Kaydedildi");
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (SQL_database.Text == "")
            {
                MessageBox.Show("lütfen veri tabanı kullanıcı adı giriniz");
            }
            baglanti = new SqlConnection("Server=" + this.SQL_server.Text + ";Database=" + this.SQL_database.Text + ";User Id=" + this.SQL_Kullanici.Text + ";Password=" + this.SQL_sifre.Text + ";");

            baglanti.Open();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.SQL_server.Text = Posta.Default.SQL_Server.ToString();
            this.SQL_database.Text = Posta.Default.SQL_Database.ToString();
            this.SQL_Kullanici.Text = Posta.Default.SQL_User.ToString();
            this.SQL_sifre.Text = Posta.Default.SQL_Password.ToString();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string BaglantiAdresi = string.Concat(new string[]
           {
                "Server=",
                this.SQL_server.Text,
                ";Database=",
                this.SQL_database.Text,
                ";User Id=",
                this.SQL_Kullanici.Text,
                ";Password=",
                this.SQL_sifre.Text,
                ";"
           });
            SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Baglanti;
            Baglanti.Open();
            try
            {
                cmd.CommandText = "CREATE TABLE DEPO_KİMYASAL ( SID int IDENTITY(1,1) PRIMARY KEY, Tarih  varchar(150), Yil  varchar(150), Ay  varchar(150), Hafta varchar(150), Kimyasal_Ad varchar(150),Bolum varchar(150), Islev varchar(100), Islem varchar(100), Cıkıs_miktar int, Kalan_miktar int, Giris int,Cıkıs int, Lot_no varchar(150), Barcode varchar(150) )";
                cmd.ExecuteNonQuery();
            }
            catch { }
            try
            {
                cmd.CommandText = "CREATE TABLE Depo_Malzeme ( SID int IDENTITY(1,1) PRIMARY KEY, Malzeme_Adi  varchar(150), Malzeme_Kodu  varchar(150), Kalan_Miktar  varchar(150) )";
                cmd.ExecuteNonQuery();
            }
            catch
            {

            }
        }
    }
}
