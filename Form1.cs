using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows.Forms;
using System.Net;
using OfficeOpenXml.Table;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WildberriesComparisonTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //for top most
            #if DEBUG
            TopMost = false;
            #endif
            this.button1.Click += button1_Click1;
        }

        private void button1_Click(object sender, EventArgs e) { }
        private async void button1_Click1(object sender, EventArgs e)
        {
            bool close = false;

            await Task.Run(() =>
            {
                if (checkBox1.Checked == true)
                {
                    /*
                    this.BasicMethod(close);
                    */
                }
                else
                {
                    this.BasicMethod(true);
                }
            });

        }
        public void BasicMethod(bool close)
        {
            //path for excel
            var location = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var path_this_project = System.IO.Path.GetDirectoryName(location);
            var path_excel = path_this_project + @"\Comparing_table.xlsx";
            //read
            var excel_read = new ReadAndWriteExcel();
            //project Excel packgage
            excel_read.MyExcelTable = excel_read.CreateExcelPackage(path_excel);
            //Excel worksheet project of definite number
            var my_work_exc = excel_read.CreateExcelWorksheet(1);

            //Compatitor comparing
            string response_product_json_competitor = String.Empty;
            string response_product_json_client = String.Empty;
            //Get string jsonn from excel doc
            excel_read.ReadExcelAndGetJson(my_work_exc, out response_product_json_competitor, out response_product_json_client);
            //Write data excel
            var comparing = new Comparing();
            comparing.CompetitorComparing(response_product_json_client, response_product_json_competitor, excel_read.MyExcelTable, my_work_exc);
            //Is done
            this.progressBar1.Maximum = 100;
            this.progressBar1.BeginInvoke((MethodInvoker)(() => 
            {
                for (int i = 0; i < progressBar1.Maximum; i++)
                {
                    this.progressBar1.Value++;
                    System.Threading.Thread.Sleep(100);
                }
            }));
            //"В процессе..." => "Готово!"
            this.label5.BackColor = System.Drawing.Color.GreenYellow;
            this.label5.BeginInvoke((MethodInvoker)(() =>this.label5.Text = "Готово!"));
            //Close if checkBox == false
            if (close) Invoke((MethodInvoker)(() => Close()));
        }
    }
}
