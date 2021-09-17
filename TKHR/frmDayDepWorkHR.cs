using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using TKITDLL;

namespace TKHR
{
    public partial class frmDayDepWorkHR : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet dsYear = new DataSet();
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;
        string NowDB = "test";
        int rownum = 0;
        string NowTable = null;

        public frmDayDepWorkHR()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {

                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString()))
                {
                    //20210902密
                    Class1 TKID = new Class1();//用new 建立類別實體
                    SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                    //資料庫使用者密碼解密
                    sqlsb.Password = TKID.Decryption(sqlsb.Password);
                    sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                    String connectionString;
                    sqlConn = new SqlConnection(sqlsb.ConnectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                   
                    sbSql.Append(@" SELECT CONVERT(varchar(8),[Date],112) AS '日期',[DepCode] AS '代號',DEP AS '部門' ,SUM(THR) AS '班別時數',SUM(OTHR) AS '加班時數'  ,SUM(NOHR) AS '請假時數'  ");
                    sbSql.Append(@" FROM (");
                    sbSql.Append(@" SELECT [Department].[Code] AS DepCode,[Department].[Name] AS DEP");
                    sbSql.Append(@" ,[AttendanceRollcall].[Date]");
                    sbSql.Append(@" ,[Employee].[CnName],[AttendanceType].[Code]");
                    sbSql.Append(@" ,[AttendanceType].[Name]");
                    sbSql.Append(@" ,[AttendanceRollcall].[Hours]");
                    sbSql.Append(@" ,[AttendanceRollcall].[QuartersHoursUnit] ");
                    sbSql.Append(@" ,(CASE WHEN [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_003' THEN [AttendanceRollcall].[Hours]/60  WHEN  [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [AttendanceRollcall].[Hours] END)+(CASE WHEN [AttendanceRank].[JobHours]<>[AttendanceRank].[WorkHours] THEN 0.5 ELSE 0 END ) AS HR ");
                    sbSql.Append(@" ,[HRAttendanceType].CType ");
                    sbSql.Append(@" ,(CASE WHEN [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_003' THEN [HRAttendanceType].CType*[AttendanceRollcall].[Hours]/60  WHEN  [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [HRAttendanceType].CType*[AttendanceRollcall].[Hours] END) +(CASE WHEN [AttendanceRank].[JobHours]<>[AttendanceRank].[WorkHours] THEN 0.5 ELSE 0 END ) AS THR  ");
                    sbSql.Append(@" ,0 AS OTHR");
                    sbSql.Append(@" ,0 AS NOHR");
                    sbSql.Append(@" FROM [HRMDB].[dbo].[AttendanceRollcall],[HRMDB].[dbo].[AttendanceType],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department],[HRMDB].[dbo].[AttendanceRank],[TKHR].[dbo].[HRAttendanceType] ");
                    sbSql.Append(@" WHERE [AttendanceRollcall].[AttendanceTypeId]=[AttendanceType].[AttendanceTypeId] ");
                    sbSql.Append(@" AND [AttendanceRollcall].[EmployeeId]=[Employee].[EmployeeId] ");
                    sbSql.Append(@" AND [Department].[DepartmentId]=[Employee].[DepartmentId] ");
                    sbSql.AppendFormat(@" AND CONVERT(varchar(8),[AttendanceRollcall].[Date],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" AND [HRAttendanceType].Code=[AttendanceType].Code COLLATE Chinese_PRC_CI_AS ");
                    sbSql.Append(@" AND [AttendanceRollcall].[AttendanceRankId]=[AttendanceRank].[AttendanceRankId]");
                    sbSql.Append(@" AND [HRAttendanceType].CType='1' ");
                    sbSql.Append(@" AND [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_003'");
                    sbSql.Append(@" UNION ALL");
                    sbSql.Append(@" SELECT [Department].[Code] AS DepCode,[Department].[Name] AS DEP");
                    sbSql.Append(@" ,[AttendanceRollcall].[Date]");
                    sbSql.Append(@" ,[Employee].[CnName],[AttendanceType].[Code]");
                    sbSql.Append(@" ,[AttendanceType].[Name]");
                    sbSql.Append(@" ,[AttendanceRollcall].[Hours]");
                    sbSql.Append(@" ,[AttendanceRollcall].[QuartersHoursUnit] ");
                    sbSql.Append(@" ,CASE WHEN [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_003' THEN [AttendanceRollcall].[Hours]/60  WHEN  [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [AttendanceRollcall].[Hours] END AS HR ");
                    sbSql.Append(@" ,[HRAttendanceType].CType");
                    sbSql.Append(@" ,0 AS THR  ");
                    sbSql.Append(@" ,0 AS OTHR");
                    sbSql.Append(@" ,CASE WHEN [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_003' THEN [HRAttendanceType].CType*[AttendanceRollcall].[Hours]/60  WHEN  [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [HRAttendanceType].CType*[AttendanceRollcall].[Hours] END AS NOHR");
                    sbSql.Append(@" FROM [HRMDB].[dbo].[AttendanceRollcall],[HRMDB].[dbo].[AttendanceType],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department],[HRMDB].[dbo].[AttendanceRank],[TKHR].[dbo].[HRAttendanceType] ");
                    sbSql.Append(@" WHERE [AttendanceRollcall].[AttendanceTypeId]=[AttendanceType].[AttendanceTypeId] ");
                    sbSql.Append(@" AND [AttendanceRollcall].[EmployeeId]=[Employee].[EmployeeId] ");
                    sbSql.Append(@" AND [Department].[DepartmentId]=[Employee].[DepartmentId] ");
                    sbSql.AppendFormat(@" AND CONVERT(varchar(8),[AttendanceRollcall].[Date],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" AND [HRAttendanceType].Code=[AttendanceType].Code COLLATE Chinese_PRC_CI_AS ");
                    sbSql.Append(@" AND [AttendanceRollcall].[AttendanceRankId]=[AttendanceRank].[AttendanceRankId]");
                    sbSql.Append(@" AND [HRAttendanceType].CType='-1' ");
                    sbSql.Append(@" UNION ALL");
                    sbSql.Append(@" SELECT [Department].[Code] AS DepCode,[Department].[Name] AS DEP");
                    sbSql.Append(@" ,[AttendanceRollcall].[Date]");
                    sbSql.Append(@" ,[Employee].[CnName],[AttendanceType].[Code]");
                    sbSql.Append(@" ,[AttendanceType].[Name]");
                    sbSql.Append(@" ,[AttendanceRollcall].[Hours]");
                    sbSql.Append(@" ,[AttendanceRollcall].[QuartersHoursUnit]");
                    sbSql.Append(@" ,CASE WHEN [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [AttendanceRollcall].[Hours]/60  WHEN  [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [AttendanceRollcall].[Hours] END AS HR ");
                    sbSql.Append(@" ,[HRAttendanceType].CType ");
                    sbSql.Append(@" ,0 AS THR  ");
                    sbSql.Append(@" ,[AttendanceRollcall].[Hours] AS OTHR");
                    sbSql.Append(@" ,0 AS NOHR");
                    sbSql.Append(@" FROM [HRMDB].[dbo].[AttendanceRollcall],[HRMDB].[dbo].[AttendanceType],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department],[HRMDB].[dbo].[AttendanceRank],[TKHR].[dbo].[HRAttendanceType]  ");
                    sbSql.Append(@" WHERE [AttendanceRollcall].[AttendanceTypeId]=[AttendanceType].[AttendanceTypeId]");
                    sbSql.Append(@" AND [AttendanceRollcall].[EmployeeId]=[Employee].[EmployeeId]");
                    sbSql.Append(@" AND [Department].[DepartmentId]=[Employee].[DepartmentId]");
                    sbSql.AppendFormat(@" AND CONVERT(varchar(8),[AttendanceRollcall].[Date],112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" AND [HRAttendanceType].Code=[AttendanceType].Code COLLATE Chinese_PRC_CI_AS");
                    sbSql.Append(@" AND [AttendanceRollcall].[AttendanceRankId]=[AttendanceRank].[AttendanceRankId] ");
                    sbSql.Append(@" AND [HRAttendanceType].CType='1'");
                    sbSql.Append(@" AND [QuartersHoursUnit]  ='AttendanceUnit_002' ");
                    sbSql.Append(@" ) AS TEMP ");
                    sbSql.Append(@" GROUP BY CONVERT(varchar(8),[Date],112),DepCode ,DEP  ");
                    sbSql.Append(@" ORDER BY CONVERT(varchar(8),[Date],112),DepCode ,DEP ");
                    sbSql.Append(@" ");


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "TEMPds");
                    sqlConn.Close();


                    if (ds.Tables["TEMPds"].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView1.DataSource = ds.Tables["TEMPds"];
                        dataGridView1.AutoResizeColumns();
                        dataGridView1.CurrentCell = dataGridView1[3, rownum];
                    }
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {

            }


        }


        public void ExcelExport()
        {
            Search();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables["TEMPds"];
            if (dt.TableName != string.Empty)
            {
                ws = wb.CreateSheet(dt.TableName);
            }
            else
            {
                ws = wb.CreateSheet("Sheet1");
            }

            ws.CreateRow(0);//第一行為欄位名稱
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
            }


            int j = 0;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                
                


                j++;
            }

            if (Directory.Exists(@"c:\temp\"))
            {
                //資料夾存在
            }
            else
            {
                //新增資料夾
                Directory.CreateDirectory(@"c:\temp\");
            }
            StringBuilder filename = new StringBuilder();
            filename.AppendFormat(@"c:\temp\每日各部門上班時數明細表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

            FileStream file = new FileStream(filename.ToString(), FileMode.Create);//產生檔案
            wb.Write(file);
            file.Close();

            MessageBox.Show("匯出完成-EXCEL放在-" + filename.ToString());
            FileInfo fi = new FileInfo(filename.ToString());
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(filename.ToString());
            }
            else
            {
                //file doesn't exist
            }


        }


        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion


    }
}
