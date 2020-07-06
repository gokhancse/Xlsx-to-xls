using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace ExcelTransfer
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            try
            {
                string str = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\20.02.18 beIN SPORTS HD 1 YAYIN AKISI.xlsx";
                string str2 = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + "\\myfile.xml";

                //OpenFileDialog file = new OpenFileDialog();
                //file.Filter = "Excel Dosyası |*.xlsx| Excel Dosyası|*.xls";
                //file.FilterIndex = 2;
                //file.RestoreDirectory = true;
                //file.CheckFileExists = false;
                //file.Title = "Excel Dosyası Seçiniz..";
                //file.Multiselect = false;
                //string DosyaYolu = "";
                //if (file.ShowDialog() == DialogResult.OK)
                //{
                //    DosyaYolu = file.FileName;
                //}
                System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + str + "'; Extended Properties='Excel 12.0 xml';"); // COnnection
                conn.Open();
                System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(@"SELECT * FROM [HD 1 AKIS$] WHERE [beIN SPORTS HD 1] LIKE '%KAMU SPOTU%'", conn);// SQL Query for requirement
                var ds = new DataSet();
                da.Fill(ds);//Transfer XLSX data to dataset

                Extensions xmlString = new Extensions();
                var xml = xmlString.ToXml(ds); // Serialize XLSX file to XML

                XmlDocument xdoc = new XmlDocument();
                xdoc.LoadXml(xml);
                xdoc.Save("myfilename.xml"); // Save the xml variable as XML file

                XmlTextReader xmlReader = new XmlTextReader("myfilename.xml"); //Then read XML file

                //var document = XDocument.Parse(xml);
                //document.Save("myfile.xml");

                ds.ReadXml(str); // Transfer to dataset XML file


                System.Data.DataTable dt = new System.Data.DataTable();

                List<string> columns = new List<string>()
                {
                    "_x0023_",
                    "_x0020_",
                    "F3",
                   "beIN_x0020_SPORTS_x0020_HD_x0020_1",
                    "F5",
                    "F6",
                    "F7",
                    "F8",
                };

                dt = ds.Tables[0];

                for (int i = 0; i < columns.Count; i++)
                {
                    dt.Columns[i].ColumnName = columns[i].ToString();
                }
                conn.Close();



                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                object Missing = Type.Missing;
                Workbook workbook = excel.Workbooks.Add(Missing);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

                // Column Headings
                int iColumn = 0;

                foreach (DataColumn c in dt.Columns)
                {
                    iColumn++;
                    excel.Cells[1, iColumn] = c.ColumnName;
                }

                // Row Data
                int iRow = sheet1.UsedRange.Rows.Count - 1;

                foreach (DataRow dr in dt.Rows)
                {
                    iRow++;
                    // Row's Cell Data
                    iColumn = 0;
                    foreach (DataColumn c in dt.Columns)
                    {
                        iColumn++;
                        excel.Cells[iRow + 1, iColumn] = dr[c.ColumnName];
                    }
                }
                workbook.Save();

            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Button2_Click(object sender, EventArgs e)
        {

            //XDocument xdoc = XDocument.Load(@"deneme.xml");//node that can contain other nodes
            //xdoc.Element("cars").

        }
    }
}

