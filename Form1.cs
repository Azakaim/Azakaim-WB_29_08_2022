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

namespace WildberriesComparisonTable
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            await Task.Run(() =>
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
            });
        }
    }
}
