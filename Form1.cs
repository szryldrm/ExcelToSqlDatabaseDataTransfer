using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data.SqlClient;

namespace ExcelToSqlDatabaseDataTransfer
{
    public partial class Form1 : Form
    {
        string[] iller = {"","ADANA", "ADIYAMAN", "AFYON", "AGRI", "AMASYA", "ANKARA", "ANTALYA", "ARTVIN",
"AYDIN", "BALIKESIR", "BILECIK", "BINGOL", "BITLIS", "BOLU", "BURDUR", "BURSA", "CANAKKALE",
"CANKIRI", "CORUM", "DENIZLI", "DIYARBAKIR", "EDIRNE", "ELAZIG", "ERZINCAN", "ERZURUM", "ESKISEHIR",
"GAZIANTEP", "GIRESUN", "GUMUSHANE", "HAKKARI", "HATAY", "ISPARTA", "MERSIN", "ISTANBUL", "IZMIR",
"KARS", "KASTAMONU", "KAYSERI", "KIRKLARELI", "KIRSEHIR", "KOCAELI", "KONYA", "KUTAHYA", "MALATYA",
"MANISA", "KAHRAMANMARAS", "MARDIN", "MUGLA", "MUS", "NEVSEHIR", "NIGDE", "ORDU", "RIZE", "SAKARYA",
"SAMSUN", "SIIRT", "SINOP", "SIVAS", "TEKIRDAG", "TOKAT", "TRABZON", "TUNCELI", "ŞANLIURFA", "USAK",
"VAN", "YOZGAT", "ZONGULDAK", "AKSARAY", "BAYBURT", "KARAMAN", "KIRIKKALE", "BATMAN", "SIRNAK",
"BARTIN", "ARDAHAN", "IGDıR", "YALOVA", "KARABUK", "KILIS", "OSMANIYE", "DUZCE"};
        string[] group = {"","FUNGISIT", "B.G.D.", "INSEKTISIT", "FUMIGANT", "HERBISIT", "RODENTISIT", "AKARASIT", "NEMATISIT",
"MOLLUSSISIT", "DIGERLERI"};

        string[] formula = {"","WG", "WP", "SP", "EC", "TB", "SC", "GR", "SL",
"TOZ", "DUST", "SG", "FS", "DP", "DS", "SIVI", "EO", "DF", "SE"};

        string[] pm = { "", "KG", "LT" };

        public Form1()
        {
            InitializeComponent();
        }

        SqlConnection baglanti = new SqlConnection("Data Source=.; Initial Catalog=TarkimDB; Integrated Security=True");

        OleDbConnection xlsxbaglanti = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=excel_dosya.xlsx; Extended Properties='Excel 12.0 Xml;HDR=YES'");
        DataTable tablo = new DataTable();
        static string tarihParse(string tarih)
        {
            string newTarih = "";
            newTarih += tarih.Substring(2, 2) + ".";
            newTarih += tarih.Substring(0, 2) + ".";
            newTarih += "2017";
            return newTarih;
        }
        private void VerileriCekButton_Click(object sender, EventArgs e)
        {
            try
            {
                xlsxbaglanti.Open();
                tablo.Clear();
                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [YeniSayfa$]", xlsxbaglanti);
                da.Fill(tablo);
                dataGridView1.DataSource = tablo;
                xlsxbaglanti.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Programda Hata Meydana Geldi." + Environment.NewLine + "Hata : " + ex.Message, "Hata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void VerileriTekTekCekButton_Click(object sender, EventArgs e)
        {
            int kayitsay = 0;
            xlsxbaglanti.Open();
            baglanti.Open();
            OleDbCommand komut = new OleDbCommand("SELECT * FROM [YeniSayfa$]", xlsxbaglanti);
            OleDbDataReader oku = komut.ExecuteReader();
            while (oku.Read())
            {
                string StokKod = oku["StokKod"].ToString(); //Customers Table
                string BarkodKod = oku["BarkodKod"].ToString();//Customers Table
                string StokName = oku["StokName"].ToString(); //Get id States Table
                string TeknikMadde = oku["TeknikMadde"].ToString(); //Phone Table
                string Grup = karakterArındırma(oku["Grup"].ToString()); //Customers Table
                string FormTip = karakterArındırma(oku["FormTip"].ToString()); //Cuscode Table
                string AmbOlcu = oku["AmbOlcu"].ToString();//Customers Table
                string KoliIciAdet = oku["KoliIciAdet"].ToString(); //Cuscode Table
                string KoliKG = oku["KoliKG"].ToString();//Customers Table
                SqlCommand kmt = new SqlCommand("Insert into Products(StockCode, BarcodeCode, StockName, TechMaterial, NumberOfParcel, ParcelMeasure, forID, gID, pmID) VALUES  ('" + StokKod + "', '" + BarkodKod + "', '" + StokName + "', '" + TeknikMadde + "', '" + Convert.ToInt32(KoliIciAdet) + "', '" + Convert.ToInt32(KoliKG) + "', '" + Convert.ToInt32(Array.IndexOf(formula, FormTip)) + "', '" + Convert.ToInt32(Array.IndexOf(group, Grup)) + "','" + Convert.ToInt32(Array.IndexOf(pm, AmbOlcu)) + "')", baglanti);
                kmt.ExecuteNonQuery();
                kayitsay++;
            }
            baglanti.Close();
            xlsxbaglanti.Close();
            MessageBox.Show("Toplam " + kayitsay + " Tane Kayıt Başarı ile Excelden Alındı", "Başarılı", MessageBoxButtons.OK, MessageBoxIcon.Information);
            kayitsay = 0;
        }

        static string karakterArındırma(string metin)
        {

            char[] türkcekarakterler = { 'ı', 'ğ', 'İ', 'Ğ', 'ç', 'Ç', 'ş', 'Ş', 'ö', 'Ö', 'ü', 'Ü' };
            char[] ingilizce = { 'i', 'g', 'I', 'G', 'c', 'C', 's', 'S', 'o', 'O', 'u', 'U' };//karakterler sırayla ingilizce karakter karşılıklarıyla yazıldı
            for (int i = 0; i < türkcekarakterler.Length; i++)
            {

                metin = metin.Replace(türkcekarakterler[i], ingilizce[i]);

            }
            return metin;

            
        }
    }
}
