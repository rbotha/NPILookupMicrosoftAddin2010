using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Threading;

namespace NPILookup2010
{

    public partial class Ribbon1

    {

        static Microsoft.Office.Interop.Excel.Application ExApp;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            ExApp = Globals.ThisAddIn.Application as Microsoft.Office.Interop.Excel.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Thread thread = new Thread(new ThreadStart(()=>GetTheData()));
            thread.Start();
            


        }

        static void GetTheData()
        {
            try
            {
                Range SelectedRange = ExApp.Selection as Microsoft.Office.Interop.Excel.Range;
                foreach (Range c in SelectedRange)
                {

                    dynamic json = JsonConvert.DeserializeObject(GET("https://npiregistry.cms.hhs.gov/api/?number=" + c.Text));
                    JArray results = json.results;




                    foreach (dynamic item in results)
                    {
                        foreach (dynamic primary in item.taxonomies)
                        {
                            bool isPrimary = (bool)primary.primary;
                            if (isPrimary)
                            {
                                ExApp.Cells[c.Row, c.Column + 2] = primary.desc;
                            }
                        }

                        //MessageBox.Show(stuff);
                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        // Returns JSON string
        static string GET(string url)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            try
            {
                WebResponse response = request.GetResponse();
                using (Stream responseStream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    return reader.ReadToEnd();
                }
            }
            catch (WebException ex)
            {
                WebResponse errorResponse = ex.Response;
                using (Stream responseStream = errorResponse.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(responseStream, Encoding.GetEncoding("utf-8"));
                    String errorText = reader.ReadToEnd();
                    // log errorText
                }
                throw;
            }
        }


    }
}
