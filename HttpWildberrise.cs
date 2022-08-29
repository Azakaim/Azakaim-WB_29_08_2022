using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace WildberriesComparisonTable
{
    public class HttpWildberrise
    {
        public HttpWildberrise(string url,out string response,out HttpStatusCode httpStatusCode)
        {
            HttpWebResponse resp = null ;
            //GOTO X
            X:
            var req = WebRequest.Create(url);
            try
            {
                resp = (HttpWebResponse)req.GetResponse();
            }
            catch 
            { 
                resp = null ; 
                System.Threading.Thread.Sleep(5000);
                System.Windows.Forms.MessageBox.Show("Error:404");
                goto X;
            }
            httpStatusCode = resp.StatusCode;
            var str = resp.GetResponseStream();
            var sr = new System.IO.StreamReader(str);
            response = sr.ReadToEnd();
            
        }
    }
}
