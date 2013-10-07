using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class _Default : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {


    }

    static string savePath;

    /// <summary>
    /// 上传文件
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="e"></param>
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

    /// <summary>
    /// 将excel中的数据导入datatable中
    /// </summary>
    /// <param name="filePath"></param>
    /// <returns></returns>
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
            DataTable table = oleCon.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
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


    /// <summary>
    /// 将datatable中的数据用linq插入数据库
    /// </summary>
    /// <param name="dtExcel"></param>
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
                            Grade = q["年级"].ToString().Trim(),
                            Num = q["学号"].ToString().Trim(),
                            Name = q["姓名"].ToString().Trim(),
                            Sex = q["性别"].ToString().Trim(),
                            College = q["学院"].ToString().Trim(),
                            Department = q["系"].ToString().Trim(),
                            Major = q["专业"].ToString().Trim(),
                            ClassNum = q["学号"].ToString().Trim().Substring(0, 9),
                            Nation = q["民族"].ToString().Trim(),
                            Address = q["家庭地址"].ToString().Trim(),
                            IDNO = q["身份证号"].ToString().Trim()
                        };
            List<Total> listEntity = new List<Total>();
            foreach (var q in query)
            {
                Total Entity = new Total();
                Entity.Grade = Convert.ToInt32(q.Grade);
                Entity.Num = q.Num;
                Entity.Name = q.Name;
                Entity.Sex = q.Sex;
                Entity.College = q.College;
                Entity.Department = q.Department;
                Entity.Major = q.Major;
                Entity.ClassNum = q.ClassNum;
                Entity.Nation = q.Nation;
                Entity.Address = q.Address;
                Entity.IDNO = q.IDNO;
                listEntity.Add(Entity);
            }
            db.Total.InsertAllOnSubmit(listEntity);
            db.SubmitChanges();
        }
    }

    protected void Button2_Click(object sender, EventArgs e)
    {
        WPEDataContext w = new WPEDataContext();
        var q = from a in w.Total
                where a.College == DropDownList1.Text
                select a;
        var qarry = q.ToArray();
        int sum = qarry.Length;
        int peoplenum = Convert.ToInt32(TextBox3.Text);
        int averge = sum / peoplenum;
        int remander = sum % peoplenum;
        System.IO.StringWriter writer = new System.IO.StringWriter();
        for (int i = 0; i < peoplenum; i++)
        {
            Response.AppendHeader("Content-Disposition", "attachment;filename=" + "崂山2013-" + TextBox2.Text + "//" + replace(i) + "+" + TextBox1.Text + ".xlsx");
            Response.ContentType = "application/ms-excel";
            Response.Charset = "utf-8";
            Response.ContentEncoding = Encoding.GetEncoding("utf-8");

            int flag = 0;
            int k = remander > 0 ? averge + 1 : averge;
            for (int j = 0; j < k; j++)
            {
                writer.Write(qarry[j + flag].Grade);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Num);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Name);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Sex);
                writer.Write("\t");
                writer.Write(qarry[j + flag].College);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Department);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Major);
                writer.Write("\t");
                writer.Write(qarry[j + flag].ClassNum);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Nation);
                writer.Write("\t");
                writer.Write(qarry[j + flag].Address);
                writer.Write("\t");
                writer.Write(qarry[j + flag].IDNO);

                writer.WriteLine();
            }
            Response.Write(writer.ToString());
            Response.End();

        }


    }

    public string replace(int i)
    {
        switch (i)
        {
            case 0:
                return "A班";
            case 2:
                return "B班";
            case 3:
                return "C班";
            case 4:
                return "D班";
            case 5:
                return "E班";
            case 6:
                return "F班";
            case 7:
                return "G班";
            case 8:
                return "H班";
            case 9:
                return "I班";
            case 10:
                return "J班";
            default:
                return "ERROR";
        }
    }

}