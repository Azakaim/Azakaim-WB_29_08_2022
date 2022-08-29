using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using Newtonsoft;
using Newtonsoft.Json;

namespace WildberriesComparisonTable
{
    public class Comparing
        {/// <summary>
        /// 
        /// </summary>
        /// <param name="json_data_client_product"></param>
        /// <param name="json_data_competitor_product"></param>
        /// <param name="myexctable"></param>
        /// <param name="myworksheet"></param>
        public void CompetitorComparing(string json_data_client_product,string json_data_competitor_product,ExcelPackage myexctable,ExcelWorksheet myworksheet)
        {
            //for calculation difference
            int prc = 0;
            List<string> price_client_all = new List<string>();
            List<string> price_competitor_all = new List<string>();
            var calulation = new Comparing();
            //Client
            var data_client = JsonConvert.DeserializeObject<Root>(json_data_client_product);
            string price_cl = String.Empty;
            string base_price_client = String.Empty;
            string name_prod_client = String.Empty;
            Dictionary<string, string> str_name_and_base_price = new Dictionary<string, string>();
            var set_data_client = new ReadAndWriteExcel();
            foreach (var i in data_client.data.products)
            {
                price_cl = Convert.ToString(i.salePriceU);
                base_price_client = price_cl.Remove(price_cl.Length - 2);
                name_prod_client = i.name;
                str_name_and_base_price.Add(name_prod_client +"/id:"+ i.id, base_price_client);
            }
            //Competitor
            var data_competitor = JsonConvert.DeserializeObject<Root>(json_data_competitor_product);
            string price_cmpttr = String.Empty;
            string base_price_competitor =String.Empty;
            string name_prod_competitor = String.Empty;
            Dictionary<string, string> str_name_and_base_price_competitor = new Dictionary<string, string>();
            var set_data_competitor = new ReadAndWriteExcel();
            foreach (var i in data_competitor.data.products)
            {
                price_cmpttr = Convert.ToString(i.salePriceU);
                base_price_competitor = price_cmpttr.Remove(price_cmpttr.Length - 2);
                name_prod_competitor = i.name;
                str_name_and_base_price_competitor.Add(name_prod_competitor + "/id:" + i.id, base_price_competitor);
            }

            //set data for client
            set_data_client.WriteDataExcel(true, str_name_and_base_price, myexctable, myworksheet);
            //set data for Competitor
            set_data_competitor.WriteDataExcel(false, str_name_and_base_price_competitor, myexctable, myworksheet);

            //Calculate difference
            #region Write common difference price

            //sort price by article
            List<string> sort_price_client_for_articul = new List<string>();
            List<string> sort_price_competitor_for_articul = new List<string>();
            List<string> price_diff_procent = new List<string>();
            List<string> price_diff_rub = new List<string>();
            List<string> link_client = new List<string>();
            List<string> link_competitor = new List<string>();
            this.SortPriceByArticle(str_name_and_base_price, str_name_and_base_price_competitor, myworksheet, ref sort_price_client_for_articul, ref sort_price_competitor_for_articul,ref link_client, ref link_competitor);

            //call calculation method
            calulation.CalculationDifference(sort_price_client_for_articul, sort_price_competitor_for_articul,ref price_diff_procent,ref price_diff_rub);
            
            //Set differet % % in excel
            for (int f = 2; f <= myworksheet.Dimension.Rows; f++)
            {
                //set procent & price rub
                myworksheet.SetValue(f, 7, price_diff_procent.ElementAt(prc));
                myworksheet.SetValue(f, 8, price_diff_rub.ElementAt(prc));
                //set link
                myworksheet.SetValue(f, 9, link_client.ElementAt(prc));
                myworksheet.SetValue(f, 10, link_competitor.ElementAt(prc));
                prc++;
            }
            myexctable.Save();
            #endregion
        }
        public void SortPriceByArticle(Dictionary<string, string> price_client_all, Dictionary<string, string> price_competitor_all, ExcelWorksheet myworksheet, ref List<string> sort_price_client_for_articul, ref List<string> sort_price_competitor_for_articul,ref List<string> link_client, ref List<string> link_competitor)
        {
            //For link 
            var Link = new Common_Code();
            string link_cl = String.Empty;
            string link_comp = String.Empty;
            //Regex for id articul
            System.Text.RegularExpressions.Regex regex = new System.Text.RegularExpressions.Regex(@"(?<=/id:).*");
            //For comparing id articuls
            string id_product_client = String.Empty;
            string value_id_excel_client = String.Empty;
            string id_product_competitor = String.Empty;
            string value_id_excel_competitor = String.Empty;
            //List for a diffence the price client
            string price_client_num = String.Empty;
            string price_competitor_num = String.Empty;
            for (int p = 2; p <= myworksheet.Dimension.Rows; p++)
            {
                //create link by article
                link_cl = Link.CreateLinkProduct(myworksheet.GetValue(p, 1).ToString());
                link_comp = Link.CreateLinkProduct(myworksheet.GetValue(p, 4).ToString());
                link_client.Add(link_cl);
                link_competitor.Add(link_comp);
                //sort price by article
                sort_price_client_for_articul.Add(myworksheet.GetValue(p, 3).ToString());
                sort_price_competitor_for_articul.Add(myworksheet.GetValue(p, 6).ToString());
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="price_client"></param>
        /// <param name="price_competitor"></param>
        /// <param name="price_diff_procent"></param>
        /// <param name="price_diff_rub"></param>
        /// <exception cref="Exception"></exception>
        public void CalculationDifference(List<string> price_client, List<string> price_competitor , ref List<string> price_diff_procent,ref List<string> price_diff_rub)
        {
            double price_client_num = 0;
            double price_competitor_num = 0;
            double procent = 0;
            double diff_price_rub = 0;
            string common_str_diff_price_procent = String.Empty;
            string common_str_diff_price_rub = String.Empty;
            bool wrong_table = (price_client.Count > price_competitor.Count ) || (price_competitor.Count > price_client.Count) ? true : false;
            if (wrong_table) throw new Exception("Wrong_Table:price_client.Count > price_competitor.Count || price_competitor.Count > price_client.Count");
            for (int i = 0; i < price_client.Count; i++)
            {
                price_client_num = Convert.ToInt32(price_client[i]);
                price_competitor_num = Convert.ToInt32(price_competitor[i]);
                if ( price_client_num > price_competitor_num)
                {
                    procent = 100 + ((price_client_num - price_competitor_num) / price_client_num) * 100;
                    diff_price_rub = price_client_num - price_competitor_num;
                    common_str_diff_price_procent = $"{Math.Round(procent,2)}% ";
                    common_str_diff_price_rub = $"Цена выше на {diff_price_rub}руб.";
                }
                else if (price_client_num < price_competitor_num)
                {
                    procent = ((price_client_num - price_competitor_num)/ price_client_num) *100;
                    diff_price_rub = price_competitor_num - price_client_num;
                    common_str_diff_price_procent = $"{100 + Math.Round(procent,2)}% ";
                    common_str_diff_price_rub = $"Цена ниже на {diff_price_rub}руб.";
                }
                price_diff_procent.Add(common_str_diff_price_procent);
                price_diff_rub.Add(common_str_diff_price_rub);
            }
        }
    }

    // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
    public class Color
    {
        public string name { get; set; }
        public int id { get; set; }
    }

    public class Data
    {
        public List<Product> products { get; set; }
    }

    public class Extended
    {
        public int basicSale { get; set; }
        public int basicPriceU { get; set; }
    }

    public class Product
    {
        public int id { get; set; }
        public int root { get; set; }
        public int kindId { get; set; }
        public int subjectId { get; set; }
        public int subjectParentId { get; set; }
        public string name { get; set; }
        public string brand { get; set; }
        public int brandId { get; set; }
        public int siteBrandId { get; set; }
        public int supplierId { get; set; }
        public int priceU { get; set; }
        public int sale { get; set; }
        public int salePriceU { get; set; }
        public Extended extended { get; set; }
        public int averagePrice { get; set; }
        public int benefit { get; set; }
        public int pics { get; set; }
        public int rating { get; set; }
        public int feedbacks { get; set; }
        public List<Color> colors { get; set; }
        public List<Size> sizes { get; set; }
        public bool diffPrice { get; set; }
        public int time1 { get; set; }
        public int time2 { get; set; }
        public int wh { get; set; }
        public int? panelPromoId { get; set; }
        public string promoTextCard { get; set; }
        public string promoTextCat { get; set; }
    }

    public class Root
    {
        public int state { get; set; }
        public Data data { get; set; }
    }

    public class Size
    {
        public string name { get; set; }
        public string origName { get; set; }
        public int rank { get; set; }
        public int optionId { get; set; }
        public List<Stock> stocks { get; set; }
        public int time1 { get; set; }
        public int time2 { get; set; }
        public int wh { get; set; }
    }

    public class Stock
    {
        public int wh { get; set; }
        public int qty { get; set; }
    }
}
