using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace MakeExcelChartsFromCSV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        List<List<double>> li = new List<List<double>>();
        List<string> fileName = new List<string>();

        /// <summary>
        /// Reading csv data
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                SetDataToList(ofd.FileNames);
            }
            label3.Visible = true;
        }

        /// <summary>
        /// set csv data to list
        /// </summary>
        /// <param name="fileNames"></param>
        private void SetDataToList(string[] fileNames)
        {
            string temp;
            foreach (string fn in fileNames)
            {
                temp = Path.GetFileName(fn);
                textBox1.Text += temp + Environment.NewLine;
                string[] s = temp.Split('.');
                fileName.Add(s[0]);
                ReadCSV(fn);
            }
        }

        /// <summary>
        /// read CSV data
        /// </summary>
        /// <param name="fileName"></param>
        private void ReadCSV(string fileName)
        {
            StreamReader sr = new StreamReader(fileName, Encoding.GetEncoding("Shift_JIS"));
            li.Add(new List<double>());
            while (sr.EndOfStream == false)
            {
                string line = sr.ReadLine();
                li[li.Count - 1].Add(double.Parse(line));
            }
        }

        /// <summary>
        /// Create Excel sheet
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, EventArgs e)
        {
            Console.WriteLine(li.Count);
            for (int i = 0; i < li.Count; i++)
            {
                Console.WriteLine("li[" + i + "]:" + li[i].Count);
            }

            makeGraph();
            label4.Visible = true;
        }

        void makeGraph()
        {
            // Open Excel API
            using (var excelApplication = new NetOffice.ExcelApi.Application())
            {
                // Add workbook
                var workBook = excelApplication.Workbooks.Add();
                workBook.Worksheets.Add();
                // set worksheet to set data
                var dataSheet = (NetOffice.ExcelApi.Worksheet)workBook.Worksheets[1];
                dataSheet.Name = "Data";
                // set worksheet to set chart
                var chartSheet = (NetOffice.ExcelApi.Worksheet)workBook.Worksheets[2];
                chartSheet.Name = "Charts";

                // input data
                progressBar1.Maximum = li.Count;
                for (int i = 0; i < li.Count; i++)
                {
                    dataSheet.Cells[1, i + 1].Value = fileName[i];
                    for (int j = 0; j < li[i].Count; j++)
                    {
                        dataSheet.Cells[j + 2, i + 1].Value = li[i][j];
                    }

                    var chart = ((NetOffice.ExcelApi.ChartObjects)chartSheet.ChartObjects()).Add(0, 200*i, 350, 200);
                    chart.Chart.ChartType = NetOffice.ExcelApi.Enums.XlChartType.xlLine;
                    chart.Chart.SetSourceData(dataSheet.Range(dataSheet.Cells[1, i + 1], dataSheet.Cells[li[i].Count, i + 1]));

                    progressBar1.PerformStep();
                }

                workBook.SaveAs(sfd());
                excelApplication.Quit();
            }
        }

        string sfd()
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName = "*.xlsx";
            sfd.Filter = "Excel Book(*.xlsx;*.xls)|*.xlsx;*.xls|All Type(*.*)|*.*";
            sfd.FilterIndex = 1;
            sfd.Title = "Select directory to save file";

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                return sfd.FileName;
            }

            return "";
        }
    }
}
