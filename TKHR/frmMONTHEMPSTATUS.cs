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
    public partial class frmMONTHEMPSTATUS : Form
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
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();        
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        int rownum = 0;
        string NowTable = null;

        public frmMONTHEMPSTATUS()
        {
            DateTime dt = DateTime.Now.AddMonths(-1);
            DateTime dt2 = DateTime.Now;

            InitializeComponent();
            dateTimePicker1.Value = Convert.ToDateTime(dt.ToString("yyyy/MM/")+26);
            dateTimePicker2.Value = Convert.ToDateTime(dt2.ToString("yyyy/MM/") +25);
        }


        #region FUNCTION
        public void SearchEmployee()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                
                    sbSql.AppendFormat(@" SELECT [Department].[Code] AS '部門編碼'");
                    sbSql.AppendFormat(@" ,[Department].[Name] AS '部門名稱'");
                    sbSql.AppendFormat(@" ,[Employee].[Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[CnName] AS '中文名'");
                    sbSql.AppendFormat(@" ,CASE WHEN [GenderId]='Gender_001' THEN '男' ELSE '女' END AS '性別'");
                    //sbSql.AppendFormat(@" ,CASE WHEN DateDiff(MONTH,[Date],'{0}')=0 THEN CONVERT(varchar(100),[Date],111)  END AS '到職日期'", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                    sbSql.AppendFormat(@"  , CONVERT(varchar(100),[Employee].[Date],111) AS '到職日期'");
                    sbSql.AppendFormat(@" ,CASE WHEN [LastWorkDate]='9999/12/31' THEN NULL ELSE CONVERT(varchar(100),[LastWorkDate],111)  END AS '離職日期'");
                    sbSql.AppendFormat(@" ,CASE WHEN [LastWorkDate]='9999/12/31' THEN '到職' ELSE '離職' END AS STATUS");
                    //sbSql.AppendFormat(@" ,[EmployeeId]");
                    sbSql.AppendFormat(@" FROM [HRMDB].[dbo].[Employee], [HRMDB].[dbo].[Department]");
                    sbSql.AppendFormat(@" WHERE [Employee].[DepartmentId]=[Department].[DepartmentId]");
                    sbSql.AppendFormat(@" AND (([Date]>='{0}' AND [Date]<='{1}')OR ([LastWorkDate]>='{2}' AND [LastWorkDate]<='{3}'))",dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"), dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                    sbSql.AppendFormat(@" ORDER BY [Department].[Code],[LastWorkDate]");
                    sbSql.AppendFormat(@" ");

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
                        dataGridView1.CurrentCell = dataGridView1[0, rownum];
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

        public void SearchSalaryFixedDetail()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();

                    sbSql.AppendFormat(@" SELECT [Department].[Name] AS '部門名稱'");
                    sbSql.AppendFormat(@" ,[Employee].[Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[Employee].[CnName] AS '中文名'");
                    sbSql.AppendFormat(@" ,CONVERT(INT,SUM([KeyValue])) AS '調薪資'");
                    sbSql.AppendFormat(@" ,CONVERT(varchar(100),[BeginDate],111) AS '生效日'");
                    //sbSql.AppendFormat(@" ,[Employee].[EmployeeId] ");
                    sbSql.AppendFormat(@" FROM [HRMDB].[dbo].[SalaryFixedDetail],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department]");
                    sbSql.AppendFormat(@" WHERE [Employee].[EmployeeId]=[SalaryFixedDetail].[EmployeeId]");
                    sbSql.AppendFormat(@" AND [Employee].[DepartmentId]=[Department].[DepartmentId]");
                    sbSql.AppendFormat(@" AND [SalaryFixedDetail].[BeginDate]='{0}'", dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                    sbSql.AppendFormat(@" AND [SalaryFixedDetail].[KeyValue]>0");
                    sbSql.AppendFormat(@" GROUP BY [Department].[Name] ,[Employee].[Code], [Employee].[CnName],[Employee].[EmployeeId],CONVERT(varchar(100),[BeginDate],111)");
                    sbSql.AppendFormat(@" ");


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds2.Clear();
                    adapter.Fill(ds2, "TEMPds2");
                    sqlConn.Close();


                    if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView2.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView2.AutoResizeColumns();
                        dataGridView2.CurrentCell = dataGridView2[0, rownum];
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

        public void SearchEmployeeTranslation()
        {
            try
            {
                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();

                    sbSql.AppendFormat(@" SELECT  DEP1.[Name] AS '部門名稱'");
                    sbSql.AppendFormat(@" ,[Employee].[Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[Employee].[CnName] AS '中文名'");
                    sbSql.AppendFormat(@" ,DEP2.[Name] AS '調任單位'");
                    sbSql.AppendFormat(@" ,CONVERT(varchar(100),[ApproveDate],111)  AS '生效日'");
                    sbSql.AppendFormat(@" ,Job1.[Name] AS '舊職務'");
                    sbSql.AppendFormat(@" ,Job2.[Name] AS '新職務'");
                    //sbSql.AppendFormat(@" ,[EmployeeTranslation].[EmployeeId] ");
                    sbSql.AppendFormat(@" FROM  [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[EmployeeTranslation]");
                    sbSql.AppendFormat(@" LEFT JOIN [HRMDB].[dbo].[Department] DEP1 ON DEP1.[DepartmentId]=[EmployeeTranslation].[OldDepartmentId]");
                    sbSql.AppendFormat(@" LEFT JOIN [HRMDB].[dbo].[Department] DEP2 ON DEP2.[DepartmentId]=[EmployeeTranslation].[NewDepartmentId]");
                    sbSql.AppendFormat(@" LEFT JOIN [HRMDB].[dbo].[Job] Job1 ON Job1.[JobId]=[EmployeeTranslation].[OldJobId]");
                    sbSql.AppendFormat(@" LEFT JOIN [HRMDB].[dbo].[Job] Job2 ON Job2.[JobId]=[EmployeeTranslation].[NewJobId]");
                    sbSql.AppendFormat(@" WHERE [EmployeeTranslation].[EmployeeId]=[Employee].[EmployeeId]");
                    sbSql.AppendFormat(@" AND ([ApproveDate]>='{0}' AND [ApproveDate]<='{1}')", dateTimePicker1.Value.ToString("yyyy/MM/dd"), dateTimePicker2.Value.ToString("yyyy/MM/dd"));
                    sbSql.AppendFormat(@" ");


                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds3.Clear();
                    adapter.Fill(ds3, "TEMPds3");
                    sqlConn.Close();


                    if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView3.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView3.AutoResizeColumns();
                        dataGridView3.CurrentCell = dataGridView3[0, rownum];
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
        #endregion


        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SearchEmployee();
            SearchSalaryFixedDetail();
            SearchEmployeeTranslation();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        #endregion

    }
}
