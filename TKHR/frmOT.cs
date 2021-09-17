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
    public partial class frmOT : Form
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
        string NowDay;

        string ID;
        string Code;
        string CnName;
        DateTime OtDate;
        decimal OtHours;
        decimal OtADJHours;

        public frmOT()
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
                    sbSql.AppendFormat(@" SELECT ");
                    sbSql.AppendFormat(@" [AttendanceOTResult_Employee_EmployeeId].[Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId].[CnName] AS '姓名'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_CodeInfo_OvertimeKindId].[ScName] AS '班次'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[Hours] AS '加班時數'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[Date],111) AS '加班日'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[BeginDate],111) AS '加班開始日'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[BeginTime] AS '加班開始時間'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[EndDate],111) AS '加班結束日'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[EndTime] AS '加班結束時間'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId].[Name] AS '部門'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId_Department_DirectDeptId].[Name] AS '課'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId_Job_JobId].[Name] AS '職稱'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_AttendanceOverTimePlan_AttendanceOTPlanId].[Name] AS '班別'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_AttendanceRank_AttendanceRankId].[Name] AS '加班別'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_AttendanceType_AttendanceTypeId].[Name] AS '加班'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[AttendanceOTResultId] AS 'ID'");
                    sbSql.AppendFormat(@" FROM [HRMDB].dbo.[AttendanceOTResult] AS [AttendanceOTResult]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Employee] AS [AttendanceOTResult_Employee_EmployeeId] ON [AttendanceOTResult].[EmployeeId]=[AttendanceOTResult_Employee_EmployeeId].[EmployeeId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Department] AS [AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId] ON [AttendanceOTResult_Employee_EmployeeId].[DepartmentId]=[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId].[DepartmentId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Department] AS [AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId_Department_DirectDeptId] ON [AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId].[DirectDeptId]=[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId_Department_DirectDeptId].[DepartmentId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Job] AS [AttendanceOTResult_Employee_EmployeeId_Job_JobId] ON [AttendanceOTResult_Employee_EmployeeId].[JobId]=[AttendanceOTResult_Employee_EmployeeId_Job_JobId].[JobId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[AttendanceRank] AS [AttendanceOTResult_AttendanceRank_AttendanceRankId] ON [AttendanceOTResult].[AttendanceRankId]=[AttendanceOTResult_AttendanceRank_AttendanceRankId].[AttendanceRankId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[AttendanceOTPlan] AS [AttendanceOTResult_AttendanceOverTimePlan_AttendanceOTPlanId] ON [AttendanceOTResult].[AttendanceOTPlanId]=[AttendanceOTResult_AttendanceOverTimePlan_AttendanceOTPlanId].[AttendanceOverTimePlanId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[AttendanceType] AS [AttendanceOTResult_AttendanceType_AttendanceTypeId] ON [AttendanceOTResult].[AttendanceTypeId]=[AttendanceOTResult_AttendanceType_AttendanceTypeId].[AttendanceTypeId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[CodeInfo] AS [AttendanceOTResult_CodeInfo_OvertimeKindId] ON [AttendanceOTResult].[OvertimeKindId]=[AttendanceOTResult_CodeInfo_OvertimeKindId].[CodeInfoId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Employee] AS [AttendanceOTResult_Employee_ApproveEmployeeId] ON [AttendanceOTResult].[ApproveEmployeeId]=[AttendanceOTResult_Employee_ApproveEmployeeId].[EmployeeId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[CodeInfo] AS [AttendanceOTResult_CodeInfo_ApproveResultId] ON [AttendanceOTResult].[ApproveResultId]=[AttendanceOTResult_CodeInfo_ApproveResultId].[CodeInfoId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[User] AS [AttendanceOTResult_User_CreateBy] ON [AttendanceOTResult].[CreateBy]=[AttendanceOTResult_User_CreateBy].[UserId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[User] AS [AttendanceOTResult_User_LastModifiedBy] ON [AttendanceOTResult].[LastModifiedBy]=[AttendanceOTResult_User_LastModifiedBy].[UserId]  ");
                    sbSql.AppendFormat(@" WHERE [AttendanceOTResult].[Date]>='{0}' AND [AttendanceOTResult].[Date]<='{1}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [AttendanceOTResult_CodeInfo_OvertimeKindId].[ScName]='{0}'",comboBox1.Text.ToString());
                    sbSql.AppendFormat(@" ORDER BY [AttendanceOTResult].[Date],[AttendanceOTResult].[AttendanceOTResultId]");                    
                    sbSql.AppendFormat(@" ");
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
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(comboBox1.Text.ToString().Equals("平日早出"))
            {
                numericUpDown1.Value = 4;
            }
            else if (comboBox1.Text.ToString().Equals("平日延後"))
            {
                numericUpDown1.Value = 4;
            }
            else if (comboBox1.Text.ToString().Equals("休息日"))
            {
                numericUpDown1.Value = 12;
            }
            else if (comboBox1.Text.ToString().Equals("節日"))
            {
                numericUpDown1.Value = 12;
            }
            else if (comboBox1.Text.ToString().Equals("假日"))
            {
                numericUpDown1.Value = 0;
            }
        }

        public void SearchV2()
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
                    sbSql.AppendFormat(@" SELECT ");
                    sbSql.AppendFormat(@" [AttendanceOTResult_Employee_EmployeeId].[Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId].[CnName] AS '姓名'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_CodeInfo_OvertimeKindId].[ScName] AS '班次'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[Hours] AS '加班時數'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[Date],111) AS '加班日'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[BeginDate],111) AS '加班開始日'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[BeginTime] AS '加班開始時間'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[EndDate],111) AS '加班結束日'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[EndTime] AS '加班結束時間'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId].[Name] AS '部門'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId_Department_DirectDeptId].[Name] AS '課'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId_Job_JobId].[Name] AS '職稱'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_AttendanceOverTimePlan_AttendanceOTPlanId].[Name] AS '班別'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_AttendanceRank_AttendanceRankId].[Name] AS '加班別'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_AttendanceType_AttendanceTypeId].[Name] AS '加班'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[AttendanceOTResultId] AS 'ID'");
                    sbSql.AppendFormat(@" FROM [HRMDB].dbo.[AttendanceOTResult] AS [AttendanceOTResult]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Employee] AS [AttendanceOTResult_Employee_EmployeeId] ON [AttendanceOTResult].[EmployeeId]=[AttendanceOTResult_Employee_EmployeeId].[EmployeeId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Department] AS [AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId] ON [AttendanceOTResult_Employee_EmployeeId].[DepartmentId]=[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId].[DepartmentId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Department] AS [AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId_Department_DirectDeptId] ON [AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId].[DirectDeptId]=[AttendanceOTResult_Employee_EmployeeId_Department_DepartmentId_Department_DirectDeptId].[DepartmentId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Job] AS [AttendanceOTResult_Employee_EmployeeId_Job_JobId] ON [AttendanceOTResult_Employee_EmployeeId].[JobId]=[AttendanceOTResult_Employee_EmployeeId_Job_JobId].[JobId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[AttendanceRank] AS [AttendanceOTResult_AttendanceRank_AttendanceRankId] ON [AttendanceOTResult].[AttendanceRankId]=[AttendanceOTResult_AttendanceRank_AttendanceRankId].[AttendanceRankId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[AttendanceOTPlan] AS [AttendanceOTResult_AttendanceOverTimePlan_AttendanceOTPlanId] ON [AttendanceOTResult].[AttendanceOTPlanId]=[AttendanceOTResult_AttendanceOverTimePlan_AttendanceOTPlanId].[AttendanceOverTimePlanId] ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[AttendanceType] AS [AttendanceOTResult_AttendanceType_AttendanceTypeId] ON [AttendanceOTResult].[AttendanceTypeId]=[AttendanceOTResult_AttendanceType_AttendanceTypeId].[AttendanceTypeId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[CodeInfo] AS [AttendanceOTResult_CodeInfo_OvertimeKindId] ON [AttendanceOTResult].[OvertimeKindId]=[AttendanceOTResult_CodeInfo_OvertimeKindId].[CodeInfoId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[Employee] AS [AttendanceOTResult_Employee_ApproveEmployeeId] ON [AttendanceOTResult].[ApproveEmployeeId]=[AttendanceOTResult_Employee_ApproveEmployeeId].[EmployeeId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[CodeInfo] AS [AttendanceOTResult_CodeInfo_ApproveResultId] ON [AttendanceOTResult].[ApproveResultId]=[AttendanceOTResult_CodeInfo_ApproveResultId].[CodeInfoId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[User] AS [AttendanceOTResult_User_CreateBy] ON [AttendanceOTResult].[CreateBy]=[AttendanceOTResult_User_CreateBy].[UserId]  ");
                    sbSql.AppendFormat(@" LEFT  JOIN [HRMDB].dbo.[User] AS [AttendanceOTResult_User_LastModifiedBy] ON [AttendanceOTResult].[LastModifiedBy]=[AttendanceOTResult_User_LastModifiedBy].[UserId]  ");
                    sbSql.AppendFormat(@" WHERE [AttendanceOTResult].[Date]>='{0}' AND [AttendanceOTResult].[Date]<='{1}'", dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" AND [AttendanceOTResult].[Hours]>{0}",numericUpDown1.Value.ToString());
                    sbSql.AppendFormat(@" AND [AttendanceOTResult_CodeInfo_OvertimeKindId].[ScName]='{0}'", comboBox1.Text.ToString());
                    sbSql.AppendFormat(@" ORDER BY [AttendanceOTResult].[Date],[AttendanceOTResult].[AttendanceOTResultId]");
                    sbSql.AppendFormat(@" ");
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
        public void SearchSALOTTIME()
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
                    sbSql.AppendFormat(@" SELECT ");                    
                    sbSql.AppendFormat(@" [Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[CnName] AS '姓名'");
                    sbSql.AppendFormat(@" ,[OtDate] AS '加班日'");
                    sbSql.AppendFormat(@" ,[OtHours] AS '加班時間'");
                    sbSql.AppendFormat(@" ,[OtADJHours] AS '調整加班時間'");
                    sbSql.AppendFormat(@" ,[ID] AS 'ID'");
                    sbSql.AppendFormat(@" FROM [TKHR].[dbo].[SALOTTIME]");
                    sbSql.AppendFormat(@" WHERE [OtDate]>='{0}' AND [OtDate]<='{1}'",dateTimePicker3.Value.ToString("yyyyMMdd"), dateTimePicker4.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" ");
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

        public void ADJUSTOTTIME()
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                ID= row.Cells["ID"].Value.ToString();
                Code = row.Cells["工號"].Value.ToString();
                CnName = row.Cells["姓名"].Value.ToString();
                OtDate = Convert.ToDateTime(row.Cells["加班日"].Value.ToString());
                OtHours = Convert.ToDecimal(row.Cells["加班時數"].Value.ToString());
                OtADJHours = Convert.ToDecimal(row.Cells["加班時數"].Value.ToString())-numericUpDown1.Value;

                // MessageBox.Show(ID + " " + Code + " " + CnName + " " + OtDate + " " + OtHours + " " + OtADJHours);

                INSERTOTTIME(ID, Code, CnName, OtDate, OtHours, OtADJHours);
            }

            SearchSALOTTIME();
        }

        public void INSERTOTTIME(string ID, string Code, string CnName,DateTime OtDate,decimal OtHours, decimal OtADJHours)
        {
            int result;
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" INSERT INTO [TKHR].[dbo].[SALOTTIME] ( [ID],[Code],[CnName],[OtDate],[OtHours],[OtADJHours])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}' ,{4} ,{5})",ID,Code,CnName,OtDate.ToString("yyyy/MM/dd"), OtHours,OtADJHours);
                sbSql.AppendFormat(" ");



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


        public void SearchSALOTTIMEV2()
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
                    sbSql.AppendFormat(@" SELECT ");
                    sbSql.AppendFormat(@" [Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[CnName] AS '姓名'");
                    sbSql.AppendFormat(@" ,CONVERT(NVARCHAR,[OtDate],111) AS '加班日'");
                    sbSql.AppendFormat(@" ,[OtHours] AS '加班時間'");
                    sbSql.AppendFormat(@" ,[OtADJHours] AS '調整加班時間'");
                    sbSql.AppendFormat(@" ,[ID] AS 'ID'");
                    sbSql.AppendFormat(@" FROM [TKHR].[dbo].[SALOTTIME]");
                    sbSql.AppendFormat(@" WHERE [OtDate]>='{0}' AND [OtDate]<='{1}'", dateTimePicker5.Value.ToString("yyyyMMdd"), dateTimePicker6.Value.ToString("yyyyMMdd"));
                    sbSql.AppendFormat(@" ");
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
        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count >= 1)
            {
                textBox1.Text = dataGridView3.CurrentRow.Cells["工號"].Value.ToString();
                textBox2.Text = dataGridView3.CurrentRow.Cells["姓名"].Value.ToString();
                textBox3.Text = dataGridView3.CurrentRow.Cells["加班日"].Value.ToString();
                textBox4.Text = dataGridView3.CurrentRow.Cells["調整加班時間"].Value.ToString();
                textBox5.Text = dataGridView3.CurrentRow.Cells["ID"].Value.ToString();
                numericUpDown2.Value= Convert.ToDecimal(dataGridView3.CurrentRow.Cells["調整加班時間"].Value.ToString());
            }
        }

        public void UPDATESALOTTIME()
        {
            int result;
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [TKHR].[dbo].[SALOTTIME] SET [OtADJHours]={0} WHERE [ID]='{1}'",numericUpDown2.Value,textBox5.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");



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
                    MessageBox.Show("完成");
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
        public void DELSALOTTIME()
        {
            int result;
            try
            {

                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" DELETE [TKHR].[dbo].[SALOTTIME]  WHERE [ID]='{0}'",  textBox5.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" ");



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
                    MessageBox.Show("完成");
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
            SearchV2();
        }
 
        private void button3_Click(object sender, EventArgs e)
        {
            SearchSALOTTIME();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADJUSTOTTIME();
            MessageBox.Show("完成");
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SearchSALOTTIMEV2();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            UPDATESALOTTIME();
            SearchSALOTTIMEV2();
        }
        private void button7_Click(object sender, EventArgs e)
        {
            DELSALOTTIME();
            SearchSALOTTIMEV2();
        }

        #endregion


    }
}
