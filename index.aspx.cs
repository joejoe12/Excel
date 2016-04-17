using DashBoard.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Excel
{
    public partial class index : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnLoad_Click(object sender, EventArgs e)
        {
            if (FileUpload1.HasFile)
            {
                string Extension = Path.GetExtension(FileUpload1.PostedFile.FileName);
                if (Extension.StartsWith(".xls"))
                {
                    string FileName = Path.GetFileName(FileUpload1.PostedFile.FileName);
                    string FolderPath = ConfigurationManager.AppSettings["FolderPath"];
                    string FilePath = Server.MapPath(FolderPath + FileName);
                    FileUpload1.SaveAs(FilePath);
                    ImportExcel impExcel = new ImportExcel();

                    using (DataTable dt = impExcel.LoadExcelToDataBase(FilePath,""))
                    {
                        grd1.DataSource = dt;
                        grd1.DataBind();
                    }
                } 
            }
        }
    }
}