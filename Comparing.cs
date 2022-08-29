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
    {
        public void CompetitorComparing(string json_data_client_product,string json_data_competitor_product,ExcelPackage myexctable,ExcelWorksheet myworksheet)
        {
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
