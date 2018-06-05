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
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;

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
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSql.AppendFormat(@" SELECT ");
                    sbSql.AppendFormat(@" [AttendanceOTResult_Employee_EmployeeId].[Code] AS '工號'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_Employee_EmployeeId].[CnName] AS '姓名'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult_CodeInfo_OvertimeKindId].[ScName] AS '班次'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[Hours] AS '加班時數'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[Date],112) AS '加班日'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[BeginDate],112) AS '加班開始日'");
                    sbSql.AppendFormat(@" ,[AttendanceOTResult].[BeginTime] AS '加班開始時間'");
                    sbSql.AppendFormat(@" ,CONVERT(nvarchar,[AttendanceOTResult].[EndDate],112) AS '加班結束日'");
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
                    sbSql.AppendFormat(@" WHERE [AttendanceOTResult].[Date]>='{0}' AND [AttendanceOTResult].[Date]<='{0}'",dateTimePicker1.Value.ToString("yyyyMMdd"), dateTimePicker2.Value.ToString("yyyyMMdd"));
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

        #endregion

        #region BUTTON


        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }

        #endregion

       
    }
}
