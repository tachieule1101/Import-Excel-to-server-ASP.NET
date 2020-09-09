using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
//lien ket lieu:
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Configuration;
using System.IO;

namespace WebSiteReport.Controllers
{
    public class HomeController : Controller
    {
        //
        
        

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
        //GET: /Home/
        SqlConnection MyDB = new SqlConnection(ConfigurationManager.ConnectionStrings["MyDB"].ConnectionString);
        OleDbConnection Econ;

        //end
        

        
        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            string filename = Guid.NewGuid() + Path.GetExtension(file.FileName);
            string filepath = "/excelfolder/" + filename;
            file.SaveAs(Path.Combine(Server.MapPath("/excelfolder"), filename));
            InsertExceldata(filepath, filename);

            return View();
        }
        private void ExcelConn(string filepath)
        {
            string constr = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=""Excel 12.0 Xml;HDR=YES;""", filepath);
            Econ = new OleDbConnection(constr);
        }
        
        private void InsertExceldata(string filepath,string filename)
        {
            string fullpath = Server.MapPath("/excelfolder/") + filename;
            ExcelConn(fullpath);
            string query = string.Format("SELECT * FROM [{0}]", "TienDoCBQL$"); //get data to excel in sheet = TienDoCBQL
            OleDbCommand Ecom = new OleDbCommand(query, Econ);
            Econ.Open();

            DataSet ds = new DataSet();
            OleDbDataAdapter oda = new OleDbDataAdapter(query, Econ);
            Econ.Close();
            oda.Fill(ds);

            DataTable dt = ds.Tables[0];

            SqlBulkCopy objbulk = new SqlBulkCopy(MyDB);
            objbulk.DestinationTableName = "BaoCao"; // bang BaoCao SQL
            objbulk.ColumnMappings.Add("TT", "TT");
            objbulk.ColumnMappings.Add("MA", "MA");
            objbulk.ColumnMappings.Add("TEN", "TEN");
            objbulk.ColumnMappings.Add("THANG", "THANG");
            objbulk.ColumnMappings.Add("QUY", "QUY");
            objbulk.ColumnMappings.Add("NAM", "NAM");
            objbulk.ColumnMappings.Add("KHQUY", "KHQUY");
            objbulk.ColumnMappings.Add("KHNAM", "KHNAM");
            objbulk.ColumnMappings.Add("TLQUY", "TLQUY");
            objbulk.ColumnMappings.Add("TLNAM", "TLNAM");
            objbulk.ColumnMappings.Add("NGAY", "NGAY");
            MyDB.Open();
            try
            {
                objbulk.WriteToServer(dt);
            }
            catch (Exception)
            {
                throw;
            }  
            MyDB.Close();
        }
        //end 

        /* ----------------- xuat file excel ------------- */
        
    }
}