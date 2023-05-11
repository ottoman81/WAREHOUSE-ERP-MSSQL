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
using System.Data.OleDb;
using ExcelDataReader;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;
using System.Security.Cryptography.X509Certificates;

namespace DEPO_KİMYASAL_ERP
{
    public partial class ANA_FORM : Form
    {
        public string _path;

        public ANA_FORM()
        {
            InitializeComponent();
        }
        BackgroundWorker bw = new BackgroundWorker
        {
            WorkerReportsProgress = true,
            WorkerSupportsCancellation = true
        };
        public void cmbvericek()
        {
            try
            {
                string BaglantiAdresi = string.Concat(new string[]
              {
                "Server=",
                Posta.Default.SQL_Server.ToString(),
                ";Database=",
                Posta.Default.SQL_Database.ToString(),
                ";User Id=",
                Posta.Default.SQL_User.ToString(),
                ";Password=",
                Posta.Default.SQL_Password.ToString(),
                ";"
              });
                SqlConnection baglanti = new SqlConnection(BaglantiAdresi);

                SqlCommand komut = new SqlCommand();
                komut.CommandText = "SELECT *FROM Depo_Malzeme";
                komut.Connection = baglanti;
                komut.CommandType = CommandType.Text;

                SqlDataReader dr;
                baglanti.Open();
                dr = komut.ExecuteReader();
                while (dr.Read())
                {
                    comboBox1.Items.Add(dr["Malzeme_Adi"]);
                }
                baglanti.Close();
            }
            catch { }
        }
      
        private void ANA_FORM_Load(object sender, EventArgs e)
        {

            try {
                cmbvericek();
                Grid_Doldur1();
                Grid_Doldur();
            }
            catch { Form1 frm = new Form1(); ANA_FORM frm1 = new ANA_FORM(); frm1.Hide(); frm.Show();  }

        }
        public void Grid_Doldur()
        {
            string BaglantiAdresi = string.Concat(new string[]
            {
                "Server=",
                Posta.Default.SQL_Server.ToString(),
                ";Database=",
                Posta.Default.SQL_Database.ToString(),
                ";User Id=",
                Posta.Default.SQL_User.ToString(),
                ";Password=",
                Posta.Default.SQL_Password.ToString(),
                ";"
            });
            SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);//"Data Source='"+ Posta.Default.SQL_Server.ToString() + "';Initial Catalog=\'" + Posta.Default.SQL_Database.ToString()+ "';Integrated Security=True"
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Baglanti;
            Baglanti.Open();
            DataTable Hareket1 = new DataTable();
            SqlDataAdapter da1 = new SqlDataAdapter("SELECT * From Depo_Malzeme", Baglanti);
            da1.Fill(Hareket1);
            this.dataGridView2.DataSource = Hareket1;
            this.dataGridView2.Refresh();
        }
        public void Grid_Doldur1()
        {
            string BaglantiAdresi = string.Concat(new string[]
            {
                "Server=",
                Posta.Default.SQL_Server.ToString(),
                ";Database=",
                Posta.Default.SQL_Database.ToString(),
                ";User Id=",
                Posta.Default.SQL_User.ToString(),
                ";Password=",
                Posta.Default.SQL_Password.ToString(),
                ";"
            });
            SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);//"Data Source='"+ Posta.Default.SQL_Server.ToString() + "';Initial Catalog=\'" + Posta.Default.SQL_Database.ToString()+ "';Integrated Security=True"
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Baglanti;
            Baglanti.Open();
            DataTable Hareket = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter("SELECT * From DEPO_KİMYASAL", Baglanti);
            da.Fill(Hareket);
            this.dataGridView1.DataSource = Hareket;
            this.dataGridView1.Refresh();
        }
        private void InsertExcelRecords()
        {

            try
            {
                string BaglantiAdresi = string.Concat(new string[]
           {
                "Server=",
                Posta.Default.SQL_Server.ToString(),
                ";Database=",
                Posta.Default.SQL_Database.ToString(),
                ";User Id=",
                Posta.Default.SQL_User.ToString(),
                ";Password=",
                Posta.Default.SQL_Password.ToString(),
                ";"
           });
                SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);//"Data Source='"+ Posta.Default.SQL_Server.ToString() + "';Initial Catalog=\'" + Posta.Default.SQL_Database.ToString()+ "';Integrated Security=True"
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Baglanti;
                

                //  ExcelConn(_path);
                string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|DEPO KİMYASAL.xlsx;Extended Properties=Excel 12.0", _path);
                OleDbConnection Econ = new OleDbConnection(constr);
                string Query = string.Format("Select [TARİH],[YIL],[AY],[HAFTA],[Kimyasal Adı],[Çıkış Yapılan Bölüm],[Kimyasal İşlevi],[İŞLEM],[ÇIKIŞ MİKTARI],[KALAN MİKTAR],[GİRİŞ],[ÇIKIŞ],[LOT NO] FROM [{0}]", "DEPO TAKİP$");
                OleDbCommand Ecom = new OleDbCommand(Query, Econ);
                Econ.Open();

                DataSet ds = new DataSet();
                OleDbDataAdapter oda = new OleDbDataAdapter(Query, Econ);
                Econ.Close();
                oda.Fill(ds);
                DataTable Exceldt = ds.Tables[0];

                for (int i = Exceldt.Rows.Count - 1; i >= 0; i--)
                {
                    if (Exceldt.Rows[i]["TARİH"] == DBNull.Value || Exceldt.Rows[i]["LOT NO"] == DBNull.Value)
                    {
                        Exceldt.Rows[i].Delete();
                    }
                }
                Exceldt.AcceptChanges();
                //creating object of SqlBulkCopy
                SqlBulkCopy objbulk = new SqlBulkCopy(Baglanti);
                //assigning Destination table name
                objbulk.DestinationTableName = "DEPO_KİMYASAL";
                //Mapping Table column
                objbulk.ColumnMappings.Add("TARİH", "[Tarih]");
                objbulk.ColumnMappings.Add("YIL", "[Yil]");
                objbulk.ColumnMappings.Add("AY", "[Ay]");
                objbulk.ColumnMappings.Add("HAFTA", "[Hafta]");
                objbulk.ColumnMappings.Add("Kimyasal Adı", "[Kimyasal_Ad]");
                objbulk.ColumnMappings.Add("Çıkış Yapılan Bölüm", "[Bolum]");
                objbulk.ColumnMappings.Add("Kimyasal İşlevi", "[Islev]");
                objbulk.ColumnMappings.Add("İŞLEM", "[Islem]");
                objbulk.ColumnMappings.Add("ÇIKIŞ MİKTARI", "[Cıkıs_miktar]");
                objbulk.ColumnMappings.Add("KALAN MİKTAR", "[Kalan_miktar]");
                objbulk.ColumnMappings.Add("GİRİŞ", "[Giris]");
                objbulk.ColumnMappings.Add("ÇIKIŞ", "[Cıkıs]");
                objbulk.ColumnMappings.Add("LOT NO", "[Lot_no]");

                //inserting Datatable Records to DataBase
                SqlConnection sqlConnection = new SqlConnection();
                sqlConnection.ConnectionString = BaglantiAdresi; //Connection Details
                Baglanti.Open();
                objbulk.WriteToServer(Exceldt);
                Baglanti.Close();
                MessageBox.Show("Data has been Imported successfully.", "Imported", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            catch (Exception ex)
            {
                MessageBox.Show(string.Format("Data has not been Imported due to :{0}", ex.Message), "Not Imported", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                
              
                
            }
        }
        private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "Excell|*.xls;*.xlsx;";
            od.FileName = "DEPO KİMYASAL.xlsx";
            DialogResult dr = od.ShowDialog();
            if (dr == DialogResult.Abort)
                return;
            if (dr == DialogResult.Cancel)
                return;
            txtpath.Text = od.FileName.ToString();
            
            _path = txtpath.Text;
            if (txtpath.Text == "" || !txtpath.Text.Contains("DEPO KİMYASAL.xlsx"))
            {
                MessageBox.Show("Please Browse EmployeeList.xlsx to upload", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                txtpath.Text = "";
                
                return;
            }
            if (bw.IsBusy)
            {
                return;
            }

            System.Diagnostics.Stopwatch sWatch = new System.Diagnostics.Stopwatch();
            bw.DoWork += (bwSender, bwArg) =>
            {
                //what happens here must not touch the form
                //as it's in a different thread
                sWatch.Start();
                InsertExcelRecords();
            };

            bw.ProgressChanged += (bwSender, bwArg) =>
            {
                //update progress bars here
            };

            bw.RunWorkerCompleted += (bwSender, bwArg) =>
            {
                //now you're back in the UI thread you can update the form
                //remember to dispose of bw now

                sWatch.Stop();

                //work is done, no need for the stop button now...
               
                txtpath.Text = "";
                
                
                bw.Dispose();
            };

            //lets allow the user to click stop
            
         
            MessageBox.Show("Uploading has been started !.\nyou are free to do any other tasks in this application,if you wish to close this screen  you can do it.but please don't close this application until upload message popups.", "Upload processing..", MessageBoxButtons.OK, MessageBoxIcon.Information);

           

            //Starts the actual work - triggerrs the "DoWork" event
            bw.RunWorkerAsync();

            //InsertExcelRecords();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Add(1);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
            // changing the name of active sheet
            ws.Name = "DEPO KİMYASAL";

            ws.Rows.HorizontalAlignment = HorizontalAlignment.Center;
            // storing header part in Excel
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                ws.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }


            // storing Each row and column value to excel sheet
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    ws.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }

            // sizing the columns
            ws.Cells.EntireColumn.AutoFit();

            // save the application
            wb.SaveAs("Depo kimyasal giriş çıkış" + "'.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Exit from the application
            app.Quit();
        }

        private void btn_sil_Click(object sender, EventArgs e)
        {
            DialogResult Silelim = MessageBox.Show("Silmek istiyormusunuz ?", "Uyarı", MessageBoxButtons.YesNo);
            bool flag = Silelim == DialogResult.Yes;
            if (flag)
            {
                string KartKodu = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
                string BaglantiAdresi = string.Concat(new string[]
                {
                    "Server=",
                    Posta.Default.SQL_Server.ToString(),
                    ";Database=",
                    Posta.Default.SQL_Database.ToString(),
                    ";User Id=",
                    Posta.Default.SQL_User.ToString(),
                    ";Password=",
                    Posta.Default.SQL_Password.ToString(),
                    ";"
                });
                SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);//"Data Source='" + Posta.Default.SQL_Server.ToString() + "';Initial Catalog=\'" + Posta.Default.SQL_Database.ToString() + "';Integrated Security=True"
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Baglanti;
                Baglanti.Open();
                string silsorgu = "DELETE from DEPO_KİMYASAL Where SID='" + KartKodu + "'";
                cmd.CommandText = silsorgu;
                cmd.ExecuteNonQuery();
                Baglanti.Close();
                MessageBox.Show("Kart Silindi.");
                this.Grid_Doldur1();
            }
        }

        private void btn_guncelle_Click(object sender, EventArgs e)
        {
            int deger = int.Parse(textBox1.Text);
            DialogResult gun = MessageBox.Show("Guncellemek istiyormusunuz ?", "Uyarı", MessageBoxButtons.YesNo);
            bool flag = gun == DialogResult.Yes;
            if (flag)
            {
                string KartKodu = this.dataGridView1.CurrentRow.Cells[0].Value.ToString();
                string BaglantiAdresi = string.Concat(new string[]
                {
                    "Server=",
                    Posta.Default.SQL_Server.ToString(),
                    ";Database=",
                    Posta.Default.SQL_Database.ToString(),
                    ";User Id=",
                    Posta.Default.SQL_User.ToString(),
                    ";Password=",
                    Posta.Default.SQL_Password.ToString(),
                    ";"
                });
                SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Baglanti;
                Baglanti.Open();
                string silsorgu = "UPDATE DEPO_KİMYASAL  SET Tarih = '" + this.dateTimePicker1.Value.ToString() + "',Kimyasal_Ad ='" + this.comboBox1.Text + "',[Bolum] ='" + this.comboBox2.Text + "',[Islev] ='" + this.comboBox3.Text + "',[Islem] ='" + this.comboBox4.Text + "',[Cıkıs_miktar] ='" + deger.ToString() + "',[Lot_no] ='" + this.textBox2.Text + "',[Barcode] ='" + this.textBox3.Text + "'Where SID='" + KartKodu + "'";
                cmd.CommandText = silsorgu;
                cmd.ExecuteNonQuery();
                Baglanti.Close();
                MessageBox.Show("Güncellendi");
                this.Grid_Doldur1();
            }
            bool flag2 = gun == DialogResult.No;
            if (flag2)
            {
                MessageBox.Show("İşlem İptal Edildi");
            }
        }

        private void btn_ekle_Click(object sender, EventArgs e)
        {
            int deger = int.Parse(textBox1.Text);

            string BaglantiAdresi = string.Concat(new string[]
            {
                "Server=",
                Posta.Default.SQL_Server.ToString(),
                ";Database=",
                Posta.Default.SQL_Database.ToString(),
                ";User Id=",
                Posta.Default.SQL_User.ToString(),
                ";Password=",
                Posta.Default.SQL_Password.ToString(),
                ";"
            });
            SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Baglanti;
            Baglanti.Open();
            string SorguKayit = "INSERT into DEPO_KİMYASAL(";
            SorguKayit += "[Tarih],[Kimyasal_Ad],[Bolum],[Islev],[Islem],[Cıkıs_miktar],[Lot_no],[Barcode]) Values(";
            SorguKayit = string.Concat(new string[]
            {
                SorguKayit,
                "'",
                this.dateTimePicker1.Value.ToString(),
                "','",//BURADA KALDIM....
                this.comboBox1.Text,
                "','",
                this.comboBox2.Text,
                "','",
                this.comboBox3.Text,
                "','",
                this.comboBox4.Text,
                "','",
                deger.ToString(),
                "','",
                this.textBox2.Text,
                "','",
                this.textBox3.Text,
                "')"
            });
            cmd.CommandText = SorguKayit;
            cmd.ExecuteNonQuery();
            MessageBox.Show("Kayıt Edildi!", "DEPO KİMYASAL", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            this.Grid_Doldur1();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            sid.Text = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            dateTimePicker1.Text= dataGridView1.CurrentRow.Cells[1].Value.ToString();
            comboBox1.Text= dataGridView1.CurrentRow.Cells[5].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells[6].Value.ToString();
            comboBox3.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
            comboBox4.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
            textBox1.Text= comboBox1.Text = dataGridView1.CurrentRow.Cells[9].Value.ToString();
            textBox2.Text= comboBox1.Text = dataGridView1.CurrentRow.Cells[13].Value.ToString();
            textBox3.Text=comboBox1.Text = dataGridView1.CurrentRow.Cells[14].Value.ToString();
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
          
            dataGridView2.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            sıd2.Text = dataGridView2.CurrentRow.Cells[0].Value.ToString();
            textBox4.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
            textBox5.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
           
        
        }
        public void arttir()
        {
            
                int sayac;
                int b = dataGridView2.RowCount;
                string parca = dataGridView2.CurrentRow.Cells[2].Value.ToString();
                string deger = parca.Substring(1, 3);
                int a = int.Parse(deger);

                a++;
                textBox5.Text = "M" + "00" + b.ToString();
                sayac = b;
                Console.WriteLine(sayac);
                Console.WriteLine(b);
                sayac++;

           
        }
        private void Mal_ekle_Click(object sender, EventArgs e)
        {

            ekle();
           
        }

        private void Mal_sil_Click(object sender, EventArgs e)
        {
            DialogResult Silelim = MessageBox.Show("Silmek istiyormusunuz ?", "Uyarı", MessageBoxButtons.YesNo);
            bool flag = Silelim == DialogResult.Yes;
            if (flag)
            {
                string KartKodu = this.dataGridView2.CurrentRow.Cells[0].Value.ToString();
                string BaglantiAdresi = string.Concat(new string[]
                {
                    "Server=",
                    Posta.Default.SQL_Server.ToString(),
                    ";Database=",
                    Posta.Default.SQL_Database.ToString(),
                    ";User Id=",
                    Posta.Default.SQL_User.ToString(),
                    ";Password=",
                    Posta.Default.SQL_Password.ToString(),
                    ";"
                });
                SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);//"Data Source='" + Posta.Default.SQL_Server.ToString() + "';Initial Catalog=\'" + Posta.Default.SQL_Database.ToString() + "';Integrated Security=True"
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Baglanti;
                Baglanti.Open();
                string silsorgu = "DELETE from Depo_Malzeme Where SID='" + KartKodu + "'";
                cmd.CommandText = silsorgu;
                cmd.ExecuteNonQuery();
                Baglanti.Close();
                MessageBox.Show("Kart Silindi.");
                this.Grid_Doldur();
            }

        }

        private void Mal_guncelle_Click(object sender, EventArgs e)
        {
            DialogResult gun = MessageBox.Show("Guncellemek istiyormusunuz ?", "Uyarı", MessageBoxButtons.YesNo);
            bool flag = gun == DialogResult.Yes;
            if (flag)
            {
                string KartKodu = this.dataGridView2.CurrentRow.Cells[0].Value.ToString();
                string BaglantiAdresi = string.Concat(new string[]
                {
                    "Server=",
                    Posta.Default.SQL_Server.ToString(),
                    ";Database=",
                    Posta.Default.SQL_Database.ToString(),
                    ";User Id=",
                    Posta.Default.SQL_User.ToString(),
                    ";Password=",
                    Posta.Default.SQL_Password.ToString(),
                    ";"
                });
                SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = Baglanti;
                Baglanti.Open();
                string silsorgu = "UPDATE Depo_Malzeme  SET Malzeme_Adi = '" + textBox4.Text + "',Malzeme_Kodu ='" + textBox5.Text+ "'Where SID='" + KartKodu + "'";
                cmd.CommandText = silsorgu;
                cmd.ExecuteNonQuery();
                Baglanti.Close();
                MessageBox.Show("Güncellendi");
                //buraya kalan miktar eklenecek
                this.Grid_Doldur();
            }
            bool flag2 = gun == DialogResult.No;
            if (flag2)
            {
                MessageBox.Show("İşlem İptal Edildi");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(timer1.Enabled==false)
            {
                
                timer1.Enabled = true; 
                

            }
           else
               { timer1.Enabled = false; }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Grid_Doldur();
            string secili = listBox1.Items[listBox1.SelectedIndex].ToString();

            textBox4.Text = secili;
            listBox1.SelectedIndex = listBox1.SelectedIndex + 1;
            arttir();



            
            if (listBox1.SelectedIndex == listBox1.Items.Count - 1)
            {
                timer1.Enabled = false;
                MessageBox.Show("BİTTİ :))");
            }
            ekle();
        }
        public void ekle()
        {
            string BaglantiAdresi = string.Concat(new string[]
           {
                "Server=",
                Posta.Default.SQL_Server.ToString(),
                ";Database=",
                Posta.Default.SQL_Database.ToString(),
                ";User Id=",
                Posta.Default.SQL_User.ToString(),
                ";Password=",
                Posta.Default.SQL_Password.ToString(),
                ";"
           });
            SqlConnection Baglanti = new SqlConnection(BaglantiAdresi);
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = Baglanti;
            Baglanti.Open();
            string SorguKayit = "INSERT into Depo_Malzeme(";
            SorguKayit += "[Malzeme_Adi],[Malzeme_Kodu]) Values(";
            SorguKayit = string.Concat(new string[]
            {
                SorguKayit,
                "'",
                this.textBox4.Text,
                "','",
                this.textBox5.Text,
                "')"
            });
            cmd.CommandText = SorguKayit;
            cmd.ExecuteNonQuery();
            //MessageBox.Show("Kayıt Edildi!", "DEPO KİMYASAL", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            //buraya kalan miktar eklenecek

            this.Grid_Doldur();
        }

        private void textBox4_MouseClick(object sender, MouseEventArgs e)
        {
           // arttir();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel.Application app = new Excel.Application();
            app.Visible = true;
            Excel.Workbook wb = app.Workbooks.Add(1);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];
            // changing the name of active sheet
            ws.Name = "DEPO KİMYASAL";

            ws.Rows.HorizontalAlignment = HorizontalAlignment.Center;
            // storing header part in Excel
            for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
            {
                ws.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
            }


            // storing Each row and column value to excel sheet
            for (int i = 0; i < dataGridView2.Rows.Count - 1; i++)
            {
                for (int j = 0; j < dataGridView2.Columns.Count; j++)
                {
                    ws.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                }
            }

            // sizing the columns
            ws.Cells.EntireColumn.AutoFit();

            // save the application
            wb.SaveAs("Depo kimyasal lıstesi" + "'.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            // Exit from the application
            app.Quit();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                StreamReader oku = new StreamReader(openFileDialog1.FileName);
                string satir = oku.ReadLine();
                while (satir != null)
                {
                    listBox1.Items.Add(satir);
                    satir = oku.ReadLine();
                }
            }
        }
    }
}
