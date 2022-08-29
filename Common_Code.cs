using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WildberriesComparisonTable
{
    internal class Common_Code
    {
        public string CreateLinkProduct(string article)
        {
           return $@"https://www.wildberries.ru/catalog/{article}/detail.aspx?targetUrl=EX";
        }
    }
}
