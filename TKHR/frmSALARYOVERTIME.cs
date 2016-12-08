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

              
                sbSql.AppendFormat(@"  SELECT [OTDATE] AS '日期',[Code] AS '工號',[NAME] AS '姓名',[STIME] AS '打卡起'");
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
                sbSql.AppendFormat(@"  ,0 AS [AHOURS],0 AS[SUNITMONEY],0 AS [AUNITMONEY]");
                sbSql.AppendFormat(@"  FROM [SQL102].[Chiyu].[dbo].[DoorLog], [HRMDB].[dbo].[Employee]");
                sbSql.AppendFormat(@"  WHERE [DoorLog].[EmployeeID]=[Employee].[Code]");
                sbSql.AppendFormat(@"  AND [TerminalID] IN ('2','3')");
                sbSql.AppendFormat(@"  AND CONVERT(varchar(100),[DateTime], 112)='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"));
                sbSql.AppendFormat(@"  AND [DoorLog].[EmployeeID]='160131'");
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
        }

        #endregion

       
    }
}
