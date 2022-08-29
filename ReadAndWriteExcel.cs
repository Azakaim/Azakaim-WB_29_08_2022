using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace WildberriesComparisonTable
{
    public class ReadAndWriteExcel
    {
        
        public ExcelPackage MyExcelTable { get; set; }
        
        public ExcelPackage CreateExcelPackage(string path_excel)
        { 
            if (System.IO.File.Exists(path_excel))
            return MyExcelTable = new ExcelPackage(new System.IO.FileInfo(path_excel));
            else
            {
                MessageBox.Show(@"Please create ""Comparing_table.xlsx"" !");
                throw new Exception(@"Please create ""Comparing_table.xlsx"" !");
            }

        }
        public ExcelWorksheet CreateExcelWorksheet(int workSheet)
        {
            ExcelWorksheet variable = this.MyExcelTable.Workbook.Worksheets[workSheet] ;
            return variable ;
        }
        public void ReadExcelAndGetJson(ExcelWorksheet MyWorkSheet, out string response_product_json_competitor, out string response_product_json_client)
        {
            //product
            List<string> articul_product_client = new List<string>();
            List<string> articul_product_competitor = new List<string>();
            string art_prod_client = String.Empty;
            string art_prod_competitor = String.Empty;
            string url_all = @"https://card.wb.ru/cards/detail?spp=0&regions=64,58,83,4,38,80,33,70,82,86,30,69,22,66,31,40,1,48&pricemarginCoeff=1.0&reg=0&appType=1&emp=0&locale=ru&lang=ru&curr=rub&couponsGeo=2,12,7,3,6,13,21&dest=-1113276,-79379,-1104258,-5803327&nm=";
            string request_relevant_url_client = String.Empty;
            string request_relevant_url_competitor = String.Empty;
            int count_prod = 0;
            string url_cl = String.Empty;
            string url_cmpttr = String.Empty;
            for (int i = 2; i <= MyWorkSheet.Dimension.Rows; i++)
            {
                try
                {
                    art_prod_client = MyWorkSheet.GetValue(i, 1).ToString().Trim(' ');
                    art_prod_competitor = MyWorkSheet.GetValue(i, 4).ToString().Trim(' ');
                }
                catch (Exception)
                {
                    System.Windows.Forms.MessageBox.Show("Null is not article");
                    throw new Exception("Null is not article");
                }
                
                if (art_prod_client != "")
                {
                    articul_product_client.Add(art_prod_client.Trim(' '));
                }
                if (art_prod_competitor != "")
                {
                    articul_product_competitor.Add(art_prod_competitor.Trim(' '));
                }
            }
            //HTTP WB_client
            foreach (var i in articul_product_client)
            {
                url_cl += $"{i};";
                count_prod++;
                if (count_prod == articul_product_client.Count)
                {
                    request_relevant_url_client = url_all + url_cl;
                    count_prod = 0;
                }
            }
            new HttpWildberrise(request_relevant_url_client, out response_product_json_client, out HttpStatusCode httpStatusCode_client);
            //HTTP WB_competitor
            foreach (var i in articul_product_competitor)
            {
                url_cmpttr += $"{i};";
                count_prod++;
                if (count_prod == articul_product_competitor.Count)
                {
                    request_relevant_url_competitor = url_all + url_cmpttr;
                    count_prod = 0;
                }
            }
            new HttpWildberrise(request_relevant_url_competitor, out response_product_json_competitor, out HttpStatusCode httpStatusCode_competitor);
        }
        /// <summary>
        /// ?Who seller ;client == true ;competitor == false ; 
        /// </summary>
        /// <param name="who_saler"></param>
        /// <param name="dict"></param>
        public void WriteDataExcel(bool who_saler ,Dictionary<string, string> dict, ExcelPackage myexctable, ExcelWorksheet myworksheet)
        {
            int i = 2;
            int j = 2;
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"(?<=/id:).*");
            string id_product = String.Empty;
            string value_id_excel = String.Empty;
            switch (who_saler)
            {
                case true:
                    foreach (var item in dict)
                    {
                        for(int x = 0;x < myworksheet.Dimension.Rows;x++)
                        {
                            //Comparison of the identifier from the table and the received identifier from wildberry
                            id_product = regex.Match(item.Key).Value;
                            value_id_excel = myworksheet.GetValue(i, 1).ToString();
                            if (value_id_excel == id_product)
                            {
                                //set name
                                myworksheet.SetValue(i, 2, item.Key);
                                //set price
                                myworksheet.SetValue(i, 3, item.Value);
                                i = 1;
                                break;
                            }
                            else i++;
                        }
                        i++;
                    }
                    myexctable.Save();
                    break;

                case false:
                    foreach (var item2 in dict)
                    {
                        for (int x = 0; x < myworksheet.Dimension.Rows; x++)
                        {
                            //Comparison of the identifier from the table and the received identifier from wildberry
                            id_product = regex.Match(item2.Key).Value;
                            value_id_excel = myworksheet.GetValue(j, 4).ToString();
                            if (value_id_excel == id_product)
                            {
                                //set name
                                myworksheet.SetValue(j, 5, item2.Key);
                                //set price
                                myworksheet.SetValue(j, 6, item2.Value);
                                j = 1;
                                break;
                            }
                            else j++;
                        }
                        j++;
                    }
                    myexctable.Save(); 
                    break;

            }
        }
    }
}
