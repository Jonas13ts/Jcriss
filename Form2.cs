using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Threading;
using excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;

namespace Test_selenium
{
    public partial class Form2 : Form
    {
        //static IWebDriver driverff;
        static IWebDriver driverGC;
        Form3 use = new Form3();
        private string a;
        private string b;
        excel.Application xlApp = new excel.Application();
        public Form2()
        {
        }

        public Form2(string a, string b)
        {
            // TODO: Complete member initialization
            InitializeComponent();
            this.a = a;
            this.b = b;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            when_clicked();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            driverGC = new ChromeDriver();
            driverGC.Navigate().GoToUrl(@"F:\Project\Ramy MAM\web\mainpage.html");
            excel.Workbook xlWorkbook1 = xlApp.Workbooks.Open(textBox1.Text);
            int x = xlWorkbook1.Sheets.Count;

            String[] b = GetExcelSheetNames(textBox1.Text);
            string[] a = new string[10];



            for (int i = 0; i < x; i++)
            {
                int p = i + 1;
                excel.Worksheet xlWorksheet = xlWorkbook1.Worksheets.get_Item(p);

                excel.Range attper = xlWorksheet.UsedRange.Columns[4];
                System.Array myvalues = (System.Array)attper.Cells.Value;
                string[] strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();

                excel.Range intmark = xlWorksheet.UsedRange.Columns[6];
                System.Array myvalues1 = (System.Array)intmark.Cells.Value;
                string[] strArray1 = myvalues1.OfType<object>().Select(o => o.ToString()).ToArray();

                String[] f = new String[10];
                excel.Range xlRange = xlWorksheet.UsedRange;
                String[] w = new String[10];
                w[i] = xlWorksheet.Cells[1, 9].Value.ToString();
                f[i] = w[i];
                a[i] = b[i].Replace("$", string.Empty);

                new SelectElement(driverGC.FindElement(By.Id("Subj_list"))).SelectByIndex(p);
                driverGC.FindElement(By.Id("fill_to_all_tot")).SendKeys(w[i]);
                for (int j = 0; j < xlRange.Rows.Count; j++)
                {
                    int y = j + 1;
                    driverGC.FindElement(By.Id("totalhrs_" + y)).SendKeys(w[i]);
                    driverGC.FindElement(By.Id("atthrs_" + y)).SendKeys(strArray[j]);
                    driverGC.FindElement(By.Id("mark_" + y)).SendKeys(strArray1[j]);
                }
            }



        }
        public string[] GetExcelSheetNames(string excelFileName)
        {
            OleDbConnection con = null;
            DataTable dt = null;
            String conStr = "Provider=Microsoft.Jet.OLEDB.4.0;" +
            "Data Source=" + excelFileName + ";Extended Properties=Excel 8.0;";
            con = new OleDbConnection(conStr);
            con.Open();
            dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            if (dt == null)
            {
                return null;
            }
            String[] excelSheetNames = new String[dt.Rows.Count];
            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                excelSheetNames[i] = row["TABLE_NAME"].ToString();
                i++;
            }
            return excelSheetNames;
        }
        public void when_clicked()
        {
            OpenFileDialog ofd1 = new OpenFileDialog
            {
                InitialDirectory = @"D:\",
                Title = "Browse Text Files",
                DefaultExt = "xls",
                Filter = "Excel files (*.xls)|*.xls|(*.xlsx)|*.xlsx",
                FilterIndex = 2,
                ShowReadOnly = true
            };
            if (ofd1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = ofd1.FileName;
                excel.Workbook xlWorkbook = xlApp.Workbooks.Open(textBox1.Text);
                int x = xlWorkbook.Sheets.Count;
                excel.Worksheet xlWorksheet = xlWorkbook.Worksheets.get_Item(x);
                int c = xlWorkbook.Sheets.Count;
                List<Label> list1 = new List<Label>();
                String[] sarray = GetExcelSheetNames(textBox1.Text);
                String[] f = new String[10];
                for (int j = 0; j < c; j++)
                {
                    f[j] = sarray[j].Replace("$", string.Empty);
                    list1.Add(new Label());
                    list1[j].Location = new System.Drawing.Point(92, (j + 187 + j * 40));
                    list1[j].Text = f[j];
                    this.Controls.Add(list1[j]);

                }
                List<Label> list = new List<Label>();
                String[] r = new string[10];
                for (int i = 0; i < c; i++)
                {
                    int p = i + 1;
                    excel.Worksheet xlWorksheet1 = xlWorkbook.Worksheets.get_Item(p);
                    excel.Range xlRange = xlWorksheet1.UsedRange;
                    String[] w = new String[10];
                    w[i] = xlWorksheet1.Cells[1, 9].Value.ToString();
                    f[i] = w[i];
                    list.Add(new Label());
                    list[i].Location = new System.Drawing.Point(329, (i + 187 + i * 40));
                    list[i].Text = f[i];
                    this.Controls.Add(list[i]);
                }

            }

        }

        private void Form2_Load(object sender, EventArgs e)
        {

        }
    }
}
