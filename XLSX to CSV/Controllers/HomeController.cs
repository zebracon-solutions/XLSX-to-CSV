using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;

namespace XLSX_to_CSV.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
        public void DownloadXLSX()
        {
            using (var client = new WebClient())
            {
                try
                {
                    byte[] filedata = client.DownloadData("https://bakerhughesrigcount.gcs-web.com/static-files/b562fc2a-b229-41eb-8407-54dda5dc7295");

                    using (Stream ms = new MemoryStream(filedata))
                    {
                        using (var localStream = System.IO.File.Create("C:\\rigCounts.xlsx"))
                        {
                            byte[] buffer = new byte[1024];
                            int bytesRead;

                            do
                            {
                                bytesRead = ms.Read(buffer, 0, buffer.Length);
                                localStream.Write(buffer, 0, bytesRead);

                            } while (bytesRead > 0);

                            localStream.Flush();
                        }
                    }
                }
                catch (System.Net.WebException)
                {
                    Console.WriteLine("Ne moze da kreira SSL tunel ka podacima.");
                }
            }
        }
    }
}