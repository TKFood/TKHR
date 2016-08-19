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

namespace TKHR
{
    public partial class frmDayDepLostCard : Form
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

        public frmDayDepLostCard()
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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();
                    sbSql.Append(@" SELECT 日期,部門,姓名,上班刷卡,下班刷卡,狀況");
                    sbSql.Append(@" FROM (");
                    sbSql.Append(@" SELECT CONVERT(varchar(100),[AttendanceRollcall].[Date],112) AS '日期',[Department].Name AS '部門',CnName  AS '姓名',( CASE WHEN ISNULL([CollectBegin],'')<>''and ISNULL([CollectEnd],'')<>'' AND datepart(HH,[CollectBegin])>=13 THEN  NULL ELSE [CollectBegin] END) AS '上班刷卡',( CASE WHEN ISNULL([CollectBegin],'')<>''and ISNULL([CollectEnd],'')<>'' AND datepart(HH,[CollectBegin])<13 THEN  NULL ELSE [CollectEnd] END)   AS '下班刷卡'");
                    sbSql.Append(@" ,CASE WHEN ISNULL([CollectBegin],'')='' THEN '上班未刷' WHEN ISNULL([CollectEnd],'')='' THEN '下班未刷'   WHEN ISNULL([CollectBegin],'')<>''and ISNULL([CollectEnd],'')<>'' AND datepart(HH,[CollectBegin])>=13 THEN '下班重複刷'  WHEN ISNULL([CollectBegin],'')<>''and ISNULL([CollectEnd],'')<>'' AND datepart(HH,[CollectBegin])<=13 THEN '上班重複刷'  END AS '狀況'");
                    sbSql.Append(@" FROM [HRMDB].[dbo].[AttendanceRollcall],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department]");
                    sbSql.Append(@" WHERE [AttendanceRollcall].[EmployeeId]=[Employee].[EmployeeId]");
                    sbSql.Append(@" AND [Employee].[DepartmentId]= [Department].[DepartmentId]");
                    sbSql.Append(@" AND ((ISNULL([CollectBegin],'')='' AND  ISNULL([CollectEnd],'')<>'') OR (ISNULL([CollectBegin],'')<>'' AND  ISNULL([CollectEnd],'')='') OR (Datediff(Minute,[CollectBegin],[CollectEnd])<=5))");
                    sbSql.AppendFormat(@" AND CONVERT(varchar(100),[AttendanceRollcall].[Date],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" UNION ALL ");
                    sbSql.AppendFormat(@" SELECT '{0}',[Department].Name AS '部門',CnName  AS '姓名' ,NULL,NULL,'可能曠工'", dateTimePicker1.Value.ToString("yyyyMMdd")); ;
                    sbSql.Append(@" FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] ");
                    sbSql.Append(@" WHERE [Employee].[DepartmentId]= [Department].[DepartmentId] ");
                    sbSql.AppendFormat(@" AND ([EmployeeId] IN (SELECT [EmployeeId] FROM [HRMDB].[dbo].[AttendanceEmpRank] WHERE  [AttendanceEmpRank].[Date]='{0}') OR [EmployeeId]  IN (SELECT [EmployeeId] FROM [HRMDB].[dbo].AttendanceRankChange WHERE [Date]='{0}'))", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" AND [EmployeeId]<>'6FBF39F6-4666-4941-9FAF-A9CBBC8B1E0B'");
                    sbSql.AppendFormat(@" AND [EmployeeId] IN (SELECT [EmployeeId] FROM [HRMDB].[dbo].[AttendanceRollcall] WHERE [AttendanceRollcall].[Date]='{0}' AND ISNULL([CollectBegin],'')='' AND ISNULL([CollectEnd],'')='' AND ISNULL([EmpRankCards],'')='')", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [EmployeeId] NOT IN (SELECT [EmployeeId] FROM [HRMDB].[dbo].[AttendanceLeave] WHERE [BeginDate]>='{0}' AND [EndDate]<='{0}')", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [EmployeeId] NOT IN (SELECT [EmployeeId] FROM [HRMDB].[dbo].[TWALReg] WHERE [BeginDate]>='{0}' AND [EndDate]<='{0}') ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [EmployeeId] NOT IN (SELECT [EmployeeId] FROM [HRMDB].[dbo].[AttendanceOTRest] WHERE [BeginDate]>='{0}' AND [EndDate]<='{0}')", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [EmployeeId] NOT IN (SELECT [EmployeeId]  FROM [HRMDB].[dbo].[BusinessRegister] WHERE [BeginDate]>='{0}' AND [EndDate]<='{0}')", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [EmployeeId] NOT IN  (SELECT [EmployeeId] FROM [HRMDB].[dbo].[AttendanceEmpRank]WHERE [DATE]='{0}'AND [AttendanceRankId] IN (SELECT [AttendanceRankId]FROM [HRMDB].[dbo].[AttendanceRank]WHERE [Name] LIKE '%休息%'))", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" AND [Employee].EmployTypeId<>'EmployType_002'");
                    sbSql.Append(@" AND [Department].[DepartmentId] <>'48047BE6-156F-4364-A439-B6EE907CF87E'");
                    sbSql.AppendFormat(@" AND [EmployeeId] NOT IN (SELECT EmployeeId FROM [HRMDB].[dbo].AttendanceRollcall WHERE [DATE]='{0}' AND [AttendanceTypeId]='408')", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" ) AS TEMP");                   
                    sbSql.Append(@" ORDER BY 日期,狀況,部門 ");
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
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());



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
            filename.AppendFormat(@"c:\temp\每日各部門人員未刷卡明細表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
