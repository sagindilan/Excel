using System;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;

namespace Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void excelBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = " |*.xlsx";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string excelYolu = openFileDialog.FileName;
               
                OleDbConnection baglanti = new OleDbConnection($@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelYolu};Extended Properties='Excel 12.0 Xml;HDR=YES'");
                baglanti.Open();

                    DataTable dtExcel = baglanti.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    string sayfaAdi = dtExcel.Rows[0]["TABLE_NAME"].ToString();
                    OleDbCommand komut = new OleDbCommand($"Select * From [{sayfaAdi}]", baglanti);
                    OleDbDataAdapter da = new OleDbDataAdapter(komut);

                    DataTable data = new DataTable();
                    da.Fill(data); 
                    dataGridView1.DataSource = data;
                    
                
            }
        }
    }
}
