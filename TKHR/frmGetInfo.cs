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
using System.Text.RegularExpressions;

namespace TKHR
{
    public partial class frmGetInfo : Form
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
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        string mdate;

        public frmGetInfo()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {
                sbSql.Clear();
                sbSql = SETsbSql();

                if (!string.IsNullOrEmpty(sbSql.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);
                    adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, tablename);
                    sqlConn.Close();

                    labelget.Text = "資料筆數:" + ds.Tables[tablename].Rows.Count.ToString();

                    if (ds.Tables[tablename].Rows.Count == 0)
                    {

                    }
                    else
                    {
                        dataGridView1.DataSource = ds.Tables[tablename];
                        dataGridView1.AutoResizeColumns();
                        //rownum = ds.Tables[talbename].Rows.Count - 1;
                        dataGridView1.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                        //dataGridView1.CurrentCell = dataGridView1[0, 2];

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

        public StringBuilder SETsbSql()
        {
            StringBuilder STR = new StringBuilder();

            if (comboBox1.Text.ToString().Equals("部門請假率"))
            {
                mdate = dateTimePicker1.Value.ToString("yyyyMM") + "01";

                STR.AppendFormat(@" SELECT FDATE AS '該月第一天',LDATE  AS '該月最後一天',[Code]  AS '部門',[Name]  AS '部門名稱',PEONUM*8*MDAYS AS '總工時',HRS  AS '總請假時數',CONVERT(DECIMAL(18,2),ROUND(HRS/(PEONUM*8*MDAYS)*100,2)) AS '請假率' ");
                STR.AppendFormat(@" FROM ( ");
                STR.AppendFormat(@"  SELECT CONVERT(varchar(100),'{0}',112) AS 'FDATE'", mdate);
                STR.AppendFormat(@"  ,CONVERT(varchar(100),dateadd(dd,-datepart(dd,'{0}') ,dateadd(mm,1,'{0}')),112) AS 'LDATE'", mdate);
                STR.AppendFormat(@"  ,day(dateadd(m,1,'{0}')-day('{0}')) AS 'MDAYS' ", mdate);
                STR.AppendFormat(@"  ,[Department].[Code]");
                STR.AppendFormat(@"  ,[Department].[Name]");
                STR.AppendFormat(@"  ,[Department].[DepartmentId]");
                STR.AppendFormat(@"  ,(SELECT COUNT(*) FROM [HRMDB].[dbo].[Employee] WHERE [Employee].[DepartmentId]=[Department].[DepartmentId] AND CONVERT(varchar(100),[LastWorkDate],112) LIKE '9999%') AS'PEONUM'");
                STR.AppendFormat(@"  ,ISNULL((SELECT SUM(Hours) FROM [HRMDB].[dbo].AttendanceLeaveInfo WHERE EmployeeId IN (SELECT EmployeeId FROM [HRMDB].[dbo].[Employee] WHERE [Employee].[DepartmentId]=[Department].[DepartmentId] AND CONVERT(varchar(100),[LastWorkDate],112) LIKE '9999%') AND (CONVERT(varchar(6),[BeginDate],112)='{0}' OR  CONVERT(varchar(6),[EndDate],112) ='{0}')),0) AS 'HRS' ", mdate.Substring(0,6).ToString());
                STR.AppendFormat(@"  FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department]");
                STR.AppendFormat(@"  WHERE [Employee].[DepartmentId]=[Department].[DepartmentId]");
                STR.AppendFormat(@"  AND CONVERT(varchar(100),[LastWorkDate],112) LIKE '9999%'");
                STR.AppendFormat(@"  GROUP BY [Department].[Code],[Department].[Name],[Department].[DepartmentId]");
                STR.AppendFormat(@"  ) AS TEMP");
                STR.AppendFormat(@"  ORDER BY [Code] ");
                STR.AppendFormat(@"  ");
                tablename = "TEMPds1";
            }

            return STR;
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            int rowindex;
            // MessageBox.Show(dataGridView1.CurrentRow.Index.ToString());
            if (dataGridView1.CurrentRow != null)
            {
                rowindex = dataGridView1.CurrentRow.Index;
                if (comboBox1.Text.ToString().Equals("部門請假率"))
                {
                    row = this.dataGridView1.Rows[rowindex];
                    SEARCHDETAILHRS();
                }
            }
        }

        public void SEARCHDETAILHRS()
        {
            sbSql.Clear();
            sbSql.AppendFormat(@"  SELECT Employee.Code AS '工號',Employee.CnName AS '姓名',Hours AS '請假時數'");
            sbSql.AppendFormat(@"  FROM [HRMDB].[dbo].AttendanceLeaveInfo ,[HRMDB].[dbo].Employee");
            sbSql.AppendFormat(@"  WHERE AttendanceLeaveInfo.EmployeeId=Employee.EmployeeId");
            sbSql.AppendFormat(@"  AND AttendanceLeaveInfo.EmployeeId IN");
            sbSql.AppendFormat(@"  (SELECT EmployeeId  ");
            sbSql.AppendFormat(@"  FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department] ");
            sbSql.AppendFormat(@"  WHERE [Employee].[DepartmentId]=[Department].[DepartmentId]");
            sbSql.AppendFormat(@"  AND [Department].[Code]='{0}'", row.Cells["部門"].Value.ToString());
            sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[LastWorkDate],112) LIKE '9999%') ");
            sbSql.AppendFormat(@"  AND( CONVERT(varchar(6),[BeginDate],112)='{0}' OR  CONVERT(varchar(6),[EndDate],112) ='{0}')", dateTimePicker1.Value.ToString("yyyyMM"));
          
            sbSql.AppendFormat(@"  ");

            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            adapter = new SqlDataAdapter(sbSql.ToString(), sqlConn);
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
                //rownum = ds.Tables[talbename].Rows.Count - 1;
                dataGridView2.CurrentCell = dataGridView1.Rows[rownum].Cells[0];

                //dataGridView1.CurrentCell = dataGridView1[0, 2];

            }
        }
        public void ExcelExport()
        {
            Search();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables[tablename];
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
            if (tablename.Equals("TEMPds1"))
            {
                foreach (DataGridViewRow dr in this.dataGridView1.Rows)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));

                    j++;
                }
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
            filename.AppendFormat(@"c:\temp\查詢{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        private void button2_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion

        
    }
}
