using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using System.Reflection;
using System.Threading;
using System.Text.RegularExpressions;

namespace TKHR
{
    public partial class frmSALARYOVERTIME : Form
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
        DataTable dt = new DataTable();    
        int result;
        int rownum = 0;
        Thread TD;
     

        public frmSALARYOVERTIME()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {          
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();

              
                sbSql.AppendFormat(@"  SELECT [OTDATE] AS '日期',[Code] AS '工號',[NAME] AS '姓名'");
                sbSql.AppendFormat(@"  ,[STIME] AS '打卡起'  ,[ETIME] AS '打卡迄',[SHOURS] AS '打卡時數'");
                sbSql.AppendFormat(@"  ,[AHOURS] AS '核可時數',[SUNITMONEY] AS '時薪',[AUNITMONEY] AS '核可金額'  ");
                sbSql.AppendFormat(@"  FROM [TKHR].[dbo].[SALARYOVERTIME]");
                sbSql.AppendFormat(@"  WHERE [OTDATE]='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


               

                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    labelget.Text = "找不到資料";
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        labelget.Text = "有 " + ds.Tables["TEMPds1"].Rows.Count.ToString() + " 筆";
                        //dataGridView1.Rows.Clear();
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();
                        dataGridView1.CurrentCell = dataGridView1[0, rownum];

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        
        public void CHECKSALARYOVERTIME()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  SELECT [OTDATE] AS '日期',[Code] AS '工號',[NAME] AS '姓名',[STIME] AS '打卡起'");
                sbSql.AppendFormat(@"  ,[ETIME] AS '打卡迄',[SHOURS] AS '打卡時數',[AHOURS] AS '核可時數' ");
                sbSql.AppendFormat(@"  ,[SUNITMONEY] AS '時薪',[AUNITMONEY] AS '核可金額' ");
                sbSql.AppendFormat(@"  FROM [TKHR].[dbo].[SALARYOVERTIME]");
                sbSql.AppendFormat(@"  WHERE [OTDATE]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();

                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    ADDSALARYOVERTIME();
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        DialogResult dialogResult = MessageBox.Show("已有時數了，重新帶入會清空喔!", "確認?", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            ADDSALARYOVERTIME();
                        }
                        else if (dialogResult == DialogResult.No)
                        {
                            //do something else
                        }

                    }

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void ADDSALARYOVERTIME()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(@"  DELETE [TKHR].[dbo].[SALARYOVERTIME]");
                sbSql.AppendFormat(@"  WHERE [OTDATE]='{0}' ", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  INSERT INTO [TKHR].[dbo].[SALARYOVERTIME]");
                sbSql.AppendFormat(@"  ([OTDATE],[Code],[NAME],[STIME],[ETIME],[SHOURS],[AHOURS],[SUNITMONEY],[AUNITMONEY])");
                sbSql.AppendFormat(@"  SELECT DISTINCT CONVERT(varchar(100),[DateTime], 112) AS [OTDATE]");
                sbSql.AppendFormat(@"  ,[DoorLog].[EmployeeID]  AS [Code]");
                sbSql.AppendFormat(@"  ,[Employee].[CnName] AS [NAME]");
                sbSql.AppendFormat(@"  ,(SELECT TOP 1 [DateTime] FROM [SQL102].[Chiyu].[dbo].[DoorLog] DR2 WHERE DR2.[EmployeeID]=[DoorLog].[EmployeeID] AND CONVERT(varchar(100),DR2.[DateTime], 112)=CONVERT(varchar(100),[DoorLog].[DateTime], 112) ORDER BY [DateTime]) AS [STIME]");
                sbSql.AppendFormat(@"  ,(SELECT TOP 1 [DateTime] FROM [SQL102].[Chiyu].[dbo].[DoorLog] DR2 WHERE DR2.[EmployeeID]=[DoorLog].[EmployeeID] AND CONVERT(varchar(100),DR2.[DateTime], 112)=CONVERT(varchar(100),[DoorLog].[DateTime], 112) ORDER BY [DateTime] DESC) AS [ETIME]");
                sbSql.AppendFormat(@"  ,DATEDIFF(HOUR, (SELECT TOP 1 [DateTime] FROM [SQL102].[Chiyu].[dbo].[DoorLog] DR2 WHERE DR2.[EmployeeID]=[DoorLog].[EmployeeID] AND CONVERT(varchar(100),DR2.[DateTime], 112)=CONVERT(varchar(100),[DoorLog].[DateTime], 112) ORDER BY [DateTime]),(SELECT TOP 1 [DateTime] FROM [SQL102].[Chiyu].[dbo].[DoorLog] DR2 WHERE DR2.[EmployeeID]=[DoorLog].[EmployeeID] AND CONVERT(varchar(100),DR2.[DateTime], 112)=CONVERT(varchar(100),[DoorLog].[DateTime], 112) ORDER BY [DateTime] DESC)  ) AS [SHOURS]");
                sbSql.AppendFormat(@"  ,8 AS [AHOURS],0 AS[SUNITMONEY],0 AS [AUNITMONEY]");
                sbSql.AppendFormat(@"  FROM [SQL102].[Chiyu].[dbo].[DoorLog], [HRMDB].[dbo].[Employee]");
                sbSql.AppendFormat(@"  WHERE [DoorLog].[EmployeeID]=[Employee].[Code]");
                sbSql.AppendFormat(@"  AND [TerminalID] IN ('2','3')");
                sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[DateTime], 112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                //sbSql.AppendFormat(@"  AND [DoorLog].[EmployeeID]='160131'");
                sbSql.AppendFormat(@"  ");
                sbSql.AppendFormat(@"  ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox1.Text = row.Cells["工號"].Value.ToString();
                    textBox2.Text = row.Cells["姓名"].Value.ToString();
                    textBox3.Text = row.Cells["打卡時數"].Value.ToString();
                    textBox4.Text = row.Cells["核可時數"].Value.ToString();
                    textBox5.Text = row.Cells["時薪"].Value.ToString();
                    textBox6.Text = row.Cells["核可金額"].Value.ToString();
                    dateTimePicker2.Value = Convert.ToDateTime(row.Cells["日期"].Value.ToString().Substring(0,4)+"/"+ row.Cells["日期"].Value.ToString().Substring(4, 2)+"/" + row.Cells["日期"].Value.ToString().Substring(6, 2));
                    dateTimePicker3.Value = Convert.ToDateTime(row.Cells["打卡起"].Value.ToString());
                    dateTimePicker4.Value = Convert.ToDateTime(row.Cells["打卡迄"].Value.ToString());

                }
            }
        }

        public void SETUPDATE()
        {
            textBox4.ReadOnly = false;
            textBox6.ReadOnly = false;
            textBox4.Select();
        }
        public void SETFINISH()
        {
            textBox4.ReadOnly = true;
            textBox6.ReadOnly = true;
        }

        public void UPDATE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKHR].[dbo].[SALARYOVERTIME]  ");
                sbSql.AppendFormat(" SET [AHOURS]='{0}'  ",textBox4.Text);
                sbSql.AppendFormat(" WHERE [OTDATE]='{0}' AND [Code]='{1}'  ",dateTimePicker2.Value.ToString("yyyyMMdd"),textBox1.Text);
                sbSql.AppendFormat("   ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }

        public void SETSUNITMONEY()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("   UPDATE [TKHR].[dbo].[SALARYOVERTIME] ");
                sbSql.AppendFormat("   SET [SUNITMONEY]=[ItemValue],[AUNITMONEY]=[ItemValue]*[AHOURS]");
                sbSql.AppendFormat("   FROM [HRMDB].[dbo].[SalaryResultDetail],[HRMDB].[dbo].[SalaryResult], [HRMDB].[dbo].[Employee]");
                sbSql.AppendFormat("   WHERE [SalaryResultDetail].[SalaryResultId]=[SalaryResult].[SalaryResultId]");
                sbSql.AppendFormat("   AND [SalaryResult].[EmployeeId]=[Employee].[EmployeeId]");
                sbSql.AppendFormat("   AND ItemName='加班時薪'");
                sbSql.AppendFormat("   AND [SALARYOVERTIME].[Code] =[Employee].[Code] COLLATE Chinese_PRC_CI_AS");
                sbSql.AppendFormat("   AND [YearMonth]='{0}'",dateTimePicker1.Value.ToString("yyyyMM"));
                sbSql.AppendFormat("   ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void ExcelExport()
        {

            string NowDB = "TK";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            XSSFCellStyle cs = (XSSFCellStyle)wb.CreateCellStyle();
            //框線樣式及顏色
            cs.BorderBottom = NPOI.SS.UserModel.BorderStyle.Double;
            cs.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin;
            cs.BottomBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.LeftBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.RightBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;
            cs.TopBorderColor = NPOI.HSSF.Util.HSSFColor.Grey50Percent.Index;

            //Search();            
            dt = ds.Tables["TEMPds1"];

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
            int k = dt.Rows.Count - 1;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {

                if (j <= k)
                {
                    ws.CreateRow(j + 1);
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                    ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                    ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                    ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                    ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                    ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                    ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                    ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                    ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
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
            filename.AppendFormat(@"c:\temp\加班時數{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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

        public void SETADDHORSMONEY()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat("  UPDATE  [TKHR].[dbo].[SALARYOVERTIME] SET [AUNITMONEY]=(8*[SUNITMONEY])+(([AHOURS]-8)*[SUNITMONEY]*1.3333) ");
                sbSql.AppendFormat("  WHERE [AHOURS]>8 AND [AHOURS]<=10");
                sbSql.AppendFormat("  AND [OTDATE]='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat("  UPDATE  [TKHR].[dbo].[SALARYOVERTIME] SET [AUNITMONEY]=(8*[SUNITMONEY])+(2*[SUNITMONEY]*1.3333)+(([AHOURS]-10)*[SUNITMONEY]*1.6666)");
                sbSql.AppendFormat("  WHERE [AHOURS]>10");
                sbSql.AppendFormat("  AND [OTDATE]='{0}'", dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat("   ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  


                }
            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
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
            CHECKSALARYOVERTIME();
            Search();
            MessageBox.Show("完成");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SETUPDATE();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            UPDATE();
            if (ds.Tables["TEMPds1"].Rows.Count >= 1)
            {
                rownum = dataGridView1.CurrentCell.RowIndex;
            }
            SETFINISH();
            Search();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETSUNITMONEY();
            SETADDHORSMONEY();
            Search();
            MessageBox.Show("完成");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }
        #endregion


    }
}
