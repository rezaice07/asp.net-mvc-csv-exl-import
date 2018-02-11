using System;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using Microsoft.AspNet.Identity;
using Microsoft.AspNet.Identity.Owin;
using Microsoft.Owin.Security;
using CSVImport.Web.Models;
using System.IO;
using System.Data;
using Utility;
using System.Collections.Generic;

namespace CSVImport.Web.Controllers
{
    public class CSVImportController : Controller
    {
        public ActionResult Upload()
        {
            return View();
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Upload(HttpPostedFileBase upload)
        {
            if (ModelState.IsValid)
            {

                if (upload != null && upload.ContentLength > 0)
                {

                    var savepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "WebShared/Uploads");

                    if (!Directory.Exists(savepath))
                        Directory.CreateDirectory(savepath);

                    string savefile = Path.Combine(savepath, upload.FileName);
                    upload.SaveAs(savefile);

                    var dt = new DataTable();

                    string fileLocation = Server.MapPath("~/WebShared/Uploads/") + upload.FileName;

                    string connString = string.Empty;


                    var csvOrXlsPath = string.Format("{0}/{1}", Server.MapPath("~/WebShared/Uploads/"), upload.FileName);

                    var extensionIndex = upload.FileName.LastIndexOf(".");

                    var extension = Path.GetExtension(upload.FileName);

                    if (extension == ".csv")
                    {
                        dt = UtilityCSV.ConvertCSVtoDataTable(csvOrXlsPath);
                    }
                    //Connection String to Excel Workbook  
                    else if (extension.Trim() == "xls")
                    {
                        connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + csvOrXlsPath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        dt = UtilityCSV.ConvertXSLXtoDataTable(csvOrXlsPath, connString);
                    }
                    else if (extension.Trim() == "xlsx")
                    {
                        connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + csvOrXlsPath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        dt = UtilityCSV.ConvertXSLXtoDataTable(csvOrXlsPath, connString);
                    }
                    
                    var list = new List<CSVImportModel>();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        var newCSVImport = new CSVImportModel {
                            StudentId= dt.Rows[i][0].ToString()
                        };

                        list.Add(newCSVImport);
                    }


                    //var model = new CSVImportViewModel {
                    //    CSVImportList= list
                    //};

                    return View();
                }
                else
                {
                    ModelState.AddModelError("File", "Please Upload Your file");
                }
            }
            return View();
        }
    }
}