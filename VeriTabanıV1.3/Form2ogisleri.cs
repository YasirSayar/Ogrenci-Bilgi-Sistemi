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
using System.Runtime.InteropServices;
using System.Data.OleDb;

namespace VeriTabanıV1._3
{
    public partial class Form2ogisleri : Form
    {
        
        
        //KALITIMLA DA GETİREBİLİRİZ 
        public string adresFonksiyon(string isim)//istenilen dosyanın adresini çeken fonk.
        {
            string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName;
            string a = path + "\\Ek_dosya\\" + isim;
            return a;
        }
        OleDbConnection baglanti = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=E:\VSC ÇIKTI\ödev\1.4\VeriTabanıV1.3\Ek_dosya\Yeni Microsoft Excel Çalışma Sayfası.xlsx;Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=0;MODE=SHARE;READONLY=FALSE""");
        void tablo_olustur()
        {
            OleDbDataAdapter da = new OleDbDataAdapter("Select * From [Ogrenciler$]",baglanti);//sorgu için nesne oluşturan kısım
            DataTable tablo = new DataTable();//tablo oluşturma
            da.Fill(tablo);//da nesnesini tablo ile doldurma
            dataGridView1.DataSource = tablo;
            

        }
        public Form2ogisleri()
        {
            InitializeComponent();
            
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void Form2ogisleri_Load(object sender, EventArgs e)
        {
          
            tablo_olustur();
        }

        private void buttonKayıt_Click(object sender, EventArgs e)
        {
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("INSERT INTO [Ogrenciler$] (ID,İsim,Soyisim,Sınıf) values (@p1,@p2,@p3,@p4)", baglanti);
            komut.Parameters.AddWithValue("@p1",txt_id.Text);
            komut.Parameters.AddWithValue("@p2", txt_isim.Text);
            komut.Parameters.AddWithValue("@p3", txt_soyisim.Text);
            komut.Parameters.AddWithValue("@p4",txt_sinif.Text);
            komut.ExecuteNonQuery();//Bu metod geriye int olarak update, insert, delete olaylarından etkilenen satır sayısı döndürüyor. // işlemleri gerçekleştir demek gibi 
            baglanti.Close();
            MessageBox.Show("Ogrenci Bilgisi sisteme kaydedildi");
            tablo_olustur();
        }
    }
}
