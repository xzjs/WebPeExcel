using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {

    }

    static string savePath;

    protected void Button1_Click(object sender, EventArgs e)
    {
        string saveDir = @"Uploads\";

        string appPath = Request.PhysicalApplicationPath;

        if (FileUpload1.HasFile)
        {
            savePath = appPath + saveDir + Server.HtmlEncode(FileUpload1.FileName);

            FileUpload1.SaveAs(savePath);

            Label1.Text = "Your file was uploaded successfully.";
        }
        else
        {
            Label1.Text = "You did not specify a file to upload.";
        }

        excelInsertDataForInsert(GetExcelFileData(savePath));
        Label1.Text = "导入" + FileUpload1.FileName + "成功";
    }

    public static DataTable GetExcelFileData(string filePath)
    {
        OleDbDataAdapter oleAdp = new OleDbDataAdapter();
        OleDbConnection oleCon = new OleDbConnection();
        string strCon = @"Provider=Microsoft.ACE.OLEDB.12.0;data source=" + filePath + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1'";
        try
        {
            DataTable dt = new DataTable();
            oleCon.ConnectionString = strCon;
            oleCon.Open();
            DataTable table = oleCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,null);
            string sheetName = table.Rows[1][2].ToString();
            string sqlStr = "Select * From [" + sheetName + "]";
            oleAdp = new OleDbDataAdapter(sqlStr, oleCon);
            oleAdp.Fill(dt);
            oleCon.Close();
            return dt;
        }
        catch (Exception ex)
        {
            throw ex;
        }
        finally
        {
            oleAdp = null;
            oleCon = null;
        }
    }

    public static void excelInsertDataForInsert(DataTable dtExcel)
    {
        using (WPEDataContext db = new WPEDataContext())
        {
            //Dictionary<string, string> listPN_Lavel = getPN_Lave();
            var query = from q in dtExcel.AsEnumerable()
                        where 
                              !string.IsNullOrEmpty(q["学号"].ToString().Trim()) 
                         
                        select new
                        {
                            Grade=q["年级"].ToString().Trim(),
                            Num=q["学号"].ToString().Trim(),
                            Name=q["姓名"].ToString().Trim(),
                            Sex=q["性别"].ToString().Trim(),
                            College=q["学院"].ToString().Trim(),
                            Department=q["系"].ToString().Trim(),
                            Major=q["专业"].ToString().Trim(),
                            ClassNum=q["学号"].ToString().Trim().Substring(0,9),
                            Nation=q["民族"].ToString().Trim(),
                            Address=q["家庭地址"].ToString().Trim(),
                            IDNO=q["身份证号"].ToString().Trim()
                        };
            List<Total> listEntity = new List<Total>();
            foreach (var q in query)
            {
                Total Entity=new Total();
                Entity.Grade=Convert.ToInt32(q.Grade);
                Entity.Num=q.Num;
                Entity.Name=q.Name;
                Entity.Sex=q.Sex;
                Entity.College=q.College;
                Entity.Department=q.Department;
                Entity.Major=q.Major;
                Entity.ClassNum=q.ClassNum;
                Entity.Nation=q.Nation;
                Entity.Address=q.Address;
                Entity.IDNO=q.IDNO;
                listEntity.Add(Entity);
            }
            db.Total.InsertAllOnSubmit(listEntity);
            db.SubmitChanges();
        }
    }
}