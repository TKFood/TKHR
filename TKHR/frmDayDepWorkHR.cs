﻿using System;
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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.Append(@" SELECT CONVERT(varchar(8),[Date],112) AS '日期',DEP AS '部門' ,[Name] AS '類別',SUM(HR) AS '時數'");
                    sbSql.Append(@" FROM (");
                    sbSql.Append(@" SELECT [Department].[Name] AS DEP,[AttendanceRollcall].[Date],[Employee].[CnName],[AttendanceType].[Name],[AttendanceRollcall].[Hours],[AttendanceRollcall].[QuartersHoursUnit]");
                    sbSql.Append(@" ,CASE WHEN [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_003' THEN [AttendanceRollcall].[Hours]/60  WHEN  [AttendanceRollcall].[QuartersHoursUnit]='AttendanceUnit_002' THEN [AttendanceRollcall].[Hours] END AS HR");
                    sbSql.Append(@" FROM [HRMDB].[dbo].[AttendanceRollcall],[HRMDB].[dbo].[AttendanceType],[HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Department]");
                    sbSql.Append(@" WHERE [AttendanceRollcall].[AttendanceTypeId]=[AttendanceType].[AttendanceTypeId]");
                    sbSql.Append(@" AND [AttendanceRollcall].[EmployeeId]=[Employee].[EmployeeId]");
                    sbSql.Append(@" AND [Department].[DepartmentId]=[Employee].[DepartmentId]");
                    sbSql.AppendFormat(@" AND CONVERT(varchar(8),[AttendanceRollcall].[Date],112)='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                    sbSql.Append(@" AND [AttendanceRollcall].[QuartersHoursUnit]<>'AttendanceUnit_004' ) AS TEMP");
                    sbSql.Append(@" GROUP BY CONVERT(varchar(8),[Date],112),DEP,[Name]");
                    sbSql.Append(@" ORDER BY CONVERT(varchar(8),[Date],112),DEP,[Name]");
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