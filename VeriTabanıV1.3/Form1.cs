using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace VeriTabanıV1._3
{
    public partial class Form1 : Form
    {
      //  string b = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
        //string a = b + "\\Ek_dosya\\Yeni Microsoft Excel Çalışma Sayfası.xlsx";
        public string adresFonksiyon(string isim)//istenilen dosyanın adresini çeken fonk.
        {
            string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string a = path + "\\Ek_dosya\\"+isim;
            return a;
        }
        
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            pictureBox1.Image = Image.FromFile(adresFonksiyon("AmasyaLogo.png"));//pictureBox'ın resminin atanması
            label3.Text = "Amasya Üniversitesi\nÖgrenci Bilgi Sistemi\n      Girisi";
        }

        private void button2Temizle_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            int girilenAscıı = Convert.ToInt32(e.KeyChar);
            //sayılar ve backspace dışındaki kalvye girişlerinde uyarı veriyor.
            if (!(girilenAscıı >= 48 && girilenAscıı <= 57 || girilenAscıı == Convert.ToInt32(Keys.Back)))
            {
                e.Handled = true;  //hatalı klavye girişinin textBox'a eklenmesini engeller.
                MessageBox.Show("ID bölümüne sadece sayı girilebilir.");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.PasswordChar = '*';
        }

        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        private void buttonGiris_Click(object sender, EventArgs e)
        {
            
               //Excel dosyasını açıp, adminlerin bilgilerini kontrol eden kısım
               Excel.Application app = new Excel.Application();
            // var workbook = app.Workbooks.Open("E:\\VSC ÇIKTI\\ödev\\Excel\\Yeni Microsoft Excel Çalışma Sayfası.xlsx",0,false);
           /* string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string a = path + "\\Ek_dosya\\Yeni Microsoft Excel Çalışma Sayfası.xlsx";*/
           
            var workbook = app.Workbooks.Open(adresFonksiyon("Yeni Microsoft Excel Çalışma Sayfası"),0, false);
            var sheetAdmin = workbook.Worksheets["Admins"];
               var range = (Excel.Range)sheetAdmin.Columns["A:A"];

               /* var result = range.Find("İsim", LookAt: Excel.XlLookAt.xlWhole);
                var address = result.Address;//cell address
                var value = result.Value2;//cell value
                label4.Text = value;*/
               var range2 = (Excel.Range)sheetAdmin.Columns["B:B"];
          /*     try
               {
                var tryFind = range.Find(textBox1.Text, LookAt: Excel.XlLookAt.xlWhole);
                var try_address=tryFind.Address;
                var try_value=tryFind.Value;
                var try_sifreRow=tryFind.Row;
                var try_sifre = (string)(sheetAdmin.Cells[try_sifreRow, 2] as Excel.Range).Value;
                if(try_sifre==textBox2.Text)
                {
                        
                }
               }
               catch(NullReferenceException err)
               {
                MessageBox.Show("ID veya şifre geçersiz");
                //MessageBox.Show(err.Message);
                }*/
          

               var idtemp = range.Find(textBox1.Text, LookAt: Excel.XlLookAt.xlWhole);

               if (idtemp != null)
               {

                //  var address = idtemp.Address;
                //  var value = idtemp.Value;
                var sifreRow = idtemp.Row;
                   var şifre = (string)(sheetAdmin.Cells[sifreRow, 2] as Excel.Range).Value;
                   if (şifre == textBox2.Text)
                   {
                       if ( 1==(sheetAdmin.Cells[sifreRow,3] as Excel.Range).Value)
                       {
                        this.Hide();
                        Form2ogisleri formisler=new Form2ogisleri();
                           formisler.ShowDialog();
                         this.Close();
                        
                       }
                       else
                       {
                        this.Hide();
                        Form2akademi formakademi = new Form2akademi();
                           formakademi.ShowDialog();
                           this.Close();
                       }
                      // MessageBox.Show("Giriş doğru");
                   }

                   else
                       MessageBox.Show("Hatalı giriş");


               }
               else MessageBox.Show("ID veya şifre geçersiz");


            workbook.Close(false);

               app.Quit();
               sheetAdmin = null;
               workbook = null;
            
               GC.Collect();
               GC.WaitForPendingFinalizers(); // wait for memory to be released
            
               
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
