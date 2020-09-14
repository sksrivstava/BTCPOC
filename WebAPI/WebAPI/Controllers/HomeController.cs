using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Syncfusion.EJ2.Spreadsheet;
using Syncfusion.XlsIO;

namespace WebAPI.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ViewBag.Title = "Home Page";

            return View();
        }

        public ActionResult Open(OpenRequest openRequest)
        {
            return Content(Workbook.Open(openRequest));
        }

        public string Save(SaveSettings saveSettings)
        {
            ExcelEngine excelEngine = new ExcelEngine();
            IApplication application = excelEngine.Excel;
            try
            {
                // Convert Spreadsheet data as Stream
                Stream fileStream = Workbook.Save<Stream>(saveSettings);
                IWorkbook workbook = application.Workbooks.Open(fileStream);
               // var filePath = HttpContext.Server.MapPath("~/Files/") + saveSettings.FileName + ".xlsx";
               var filePath = HttpContext.Server.MapPath("~/Files/")+saveSettings.FileName + ".xlsx";
                workbook.SaveAs(filePath);
                return "Success";
            }
            catch (Exception ex)
            {
                return "Failure";
            }
        }
    }
}
