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

namespace TKHR
{
    public partial class frmChangTime : Form
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
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();

        DataTable dt = new DataTable();
        string tablename = null;

        DateTime dt1 = new DateTime();
        DateTime dt2 = new DateTime();

        StringBuilder sbSqlEXE = new StringBuilder();
        DataSet ds = new DataSet();
        int result;

        public frmChangTime()
        {
            InitializeComponent();
            comboBox1load();
        }

        #region FUNCTION
        public void comboBox1load()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [NAME],[EmployeeId] FROM [TKHR].[dbo].[CARDEMP] ORDER BY [NAME]  ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("NAME", typeof(string));
            dt.Columns.Add("EmployeeId", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            dt1 = dateTimePicker1.Value;
            dt2 = dt1.AddDays(1);
        }

        public void SEARCHTIME()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                StringBuilder sbSq2 = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();
                

                sbSql.AppendFormat(@"  SELECT TOP 1 [Time],[Date] FROM [HRMDB].dbo.AttendanceCollect");
                sbSql.AppendFormat(@"  WHERE [Date]>='{0}' AND [Date]<'{1}'", dt1.ToString("yyyy/MM/dd"), dt2.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  AND [EmployeeId] IN (SELECT [EmployeeId] FROM [TKHR].[dbo].[CARDEMP] WHERE [NAME]='{0}')",comboBox1.Text.ToString());
                sbSql.AppendFormat(@"  ORDER BY [Time]");
                sbSql.AppendFormat(@"  ");

                sbSq2.AppendFormat(@"  SELECT TOP 1 [Time],[Date] FROM [HRMDB].dbo.AttendanceCollect");
                sbSq2.AppendFormat(@"  WHERE [Date]>='{0}' AND [Date]<'{1}'", dt1.ToString("yyyy/MM/dd"), dt2.ToString("yyyy/MM/dd"));
                sbSq2.AppendFormat(@"  AND [EmployeeId] IN (SELECT [EmployeeId] FROM [TKHR].[dbo].[CARDEMP] WHERE [NAME]='{0}')", comboBox1.Text.ToString());
                sbSq2.AppendFormat(@"  ORDER BY [Time] DESC");
                sbSq2.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();

                adapter = new SqlDataAdapter(@"" + sbSq2, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                       label7.Text= (ds1.Tables["TEMPds1"].Rows[0]["Time"].ToString());
                       dateTimePicker2.Value = Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["Date"].ToString());
                    }
                }

                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {

                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        label8.Text = (ds2.Tables["TEMPds2"].Rows[0]["Time"].ToString());
                        dateTimePicker3.Value = Convert.ToDateTime(ds2.Tables["TEMPds2"].Rows[0]["Date"].ToString());
                    }
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


        public void SETAttendanceCollect()
        {
            int result;
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE [HRMDB].dbo.AttendanceRollcall");
                sbSql.AppendFormat(" SET DailyCards='{0}',EmpRankCards='{1}',CollectBegin='{2}',CollectEnd='{3}'",dateTimePicker2.Value.ToString("HH:mm")+"| "+ dateTimePicker3.Value.ToString("HH:mm"), dateTimePicker2.Value.ToString("HH:mm") + "| " + dateTimePicker3.Value.ToString("HH:mm"), dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm"));
                sbSql.AppendFormat(" WHERE  [Date] = '{0}'",dateTimePicker1.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(" AND [EmployeeId] IN (SELECT [EmployeeId] FROM [TKHR].[dbo].[CARDEMP] WHERE [NAME]='{0}')",comboBox1.Text.ToString());
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

        public void SETAttendanceRollcall()
        {
            int result;
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                sbSql.AppendFormat(" UPDATE  [HRMDB].dbo.AttendanceCollect");
                sbSql.AppendFormat(" SET  [Time]='{0}',[Date]='{1}'", dateTimePicker2.Value.ToString("HH:mm"), dateTimePicker2.Value.ToString("yyyy-MM-dd HH:mm"));
                sbSql.AppendFormat(" WHERE [Date]>='{0}' AND [Date]<'{1}'",dt1.ToString("yyy/MM/dd"), dt2.ToString("yyy/MM/dd"));
                sbSql.AppendFormat(" AND [EmployeeId] IN (SELECT [EmployeeId] FROM [TKHR].[dbo].[CARDEMP] WHERE [NAME]='{0}')", comboBox1.Text.ToString());
                sbSql.AppendFormat(" AND [Time]='{0}'", label7.Text);
                sbSql.AppendFormat(" ");
                sbSql.AppendFormat(" UPDATE  [HRMDB].dbo.AttendanceCollect");
                sbSql.AppendFormat(" SET  [Time]='{0}',[Date]='{1}'", dateTimePicker3.Value.ToString("HH:mm"), dateTimePicker3.Value.ToString("yyyy-MM-dd HH:mm"));
                sbSql.AppendFormat(" WHERE [Date]>='{0}' AND [Date]<'{1}'", dt1.ToString("yyy/MM/dd"), dt2.ToString("yyy/MM/dd"));
                sbSql.AppendFormat(" AND [EmployeeId] IN (SELECT [EmployeeId] FROM [TKHR].[dbo].[CARDEMP] WHERE [NAME]='{0}')", comboBox1.Text.ToString());
                sbSql.AppendFormat(" AND [Time]='{0}'", label8.Text);
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

        public void ADDHRCARD()
        {
            DateTime workdt1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 30, 0);
            DateTime workdt2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 17, 30, 0);
            DateTime carddt1 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 8, 20, 0);
            DateTime carddt2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 18, 30, 0);
            DateTime operdat = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 09, 10, 0);
            DateTime operdat2 = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 09, 10, 0);



            try
            {

                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sbSqlEXE.Clear();
                sbSql.Clear();

                sbSql.AppendFormat(" SELECT [Employee].[EmployeeId],[Employee].[CnName],[AttendanceRank].[Name]");
                sbSql.AppendFormat(" FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[AttendanceEmpRank],[HRMDB].[dbo].[AttendanceRank]");
                sbSql.AppendFormat(" WHERE [Employee].[EmployeeId]=[AttendanceEmpRank].[EmployeeId]");
                sbSql.AppendFormat(" AND [AttendanceEmpRank].[AttendanceRankId]=[AttendanceRank].[AttendanceRankId] ");
                sbSql.AppendFormat(" AND CONVERT(varchar(100),[AttendanceEmpRank].[Date],112)=CONVERT(varchar(100),GETDATE(),112)");
                sbSql.AppendFormat(" AND [Employee].[EmployeeId] IN (SELECT [EmployeeId] FROM [TKHR].[dbo].[CARDEMP])");
                sbSql.AppendFormat(" AND [Employee].[EmployeeId] IN (SELECT [EmployeeId] FROM [HRMDB].dbo.[Employee] WHERE CODE='{0}')",textBox1.Text);
                sbSql.AppendFormat(" ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();

                //sqlConn.Close();
                //sqlConn.Open();
                //tran = sqlConn.BeginTransaction();



                if (ds.Tables["TEMPds"].Rows.Count >= 1)
                {
                    //foreach (DataRow od in ds.Tables["TEMPds"].Rows)
                    for (int i = 0; i <= ds.Tables["TEMPds"].Rows.Count; i++)
                    {
                        sbSqlEXE.Clear();

                        sqlConn.Close();
                        sqlConn.Open();
                        tran = sqlConn.BeginTransaction();


                        //set BeginTime,EndTime
                        Random Begin = new Random();//亂數種子
                        int BeginTime = Begin.Next(15, 29);
                        Random End = new Random();//亂數種子
                        int EndTime = End.Next(25, 59);

                        string SBeginTime = "08:" + BeginTime.ToString();
                        string SEndTime = "18:" + EndTime.ToString();

                        string emp = ds.Tables["TEMPds"].Rows[i]["EmployeeId"].ToString();
                        string NAME = ds.Tables["TEMPds"].Rows[i]["CnName"].ToString();
                        string office = ds.Tables["TEMPds"].Rows[i]["Name"].ToString();
                        Guid guid1 = Guid.NewGuid();
                        Guid guid2 = Guid.NewGuid();
                        Guid guid3 = Guid.NewGuid();
                        Guid guid4 = Guid.NewGuid();


                        if (!office.ToString().Contains("休息"))
                        {
                            sbSqlEXE.AppendFormat(" ");
                            sbSqlEXE.AppendFormat(" INSERT INTO [HRMDB].[dbo].[AttendanceRollcall] ([AttendanceRollcallId],[EmployeeId],[Date],[BeginTime],[EndTime],[AttendanceRankId],[AttendanceTypeId],[Hours],[QuartersHours],[QuartersHoursUnit],[IsConfirm],[OperationDate],[UserId],[Recover],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[AssignReason],[OwnerId],[VirObjectId],[ActualBeginTime],[ActualEndTime],[Count],[DailyCards],[EmpRankCards],[CollectBegin],[CollectEnd],[IsAbnormal]) ");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1 NEWID() AS [AttendanceRollcallId]", guid1.ToString());
                            sbSqlEXE.AppendFormat(" ,[EmployeeId]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 00:00:00.000' AS [Date]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{0}'+' '+CONVERT(varchar(100),[AttendanceRollcall].[BeginTime],114) AS [BeginTime] ", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{0}'+' '+CONVERT(varchar(100),[AttendanceRollcall].[EndTime],114) AS [EndTime]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[AttendanceRankId]");
                            sbSqlEXE.AppendFormat(" ,[AttendanceTypeId]");
                            sbSqlEXE.AppendFormat(" ,[Hours]");
                            sbSqlEXE.AppendFormat(" ,[QuartersHours]");
                            sbSqlEXE.AppendFormat(" ,[QuartersHoursUnit]");
                            sbSqlEXE.AppendFormat(" ,[IsConfirm]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 09:30:00.000' AS [OperationDate]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[UserId]");
                            sbSqlEXE.AppendFormat(" ,[Recover]");
                            sbSqlEXE.AppendFormat(" ,[Remark]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 09:30:00.000' AS [CreateDate]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 09:30:00.000' AS [LastModifiedDate]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[CreateBy]");
                            sbSqlEXE.AppendFormat(" ,[LastModifiedBy]");
                            sbSqlEXE.AppendFormat(" ,[CorporationId]");
                            sbSqlEXE.AppendFormat(" ,[Flag]");
                            sbSqlEXE.AppendFormat(" ,[AssignReason]");
                            sbSqlEXE.AppendFormat(" ,[OwnerId]");
                            sbSqlEXE.AppendFormat(" ,[VirObjectId]");
                            sbSqlEXE.AppendFormat(" ,[ActualBeginTime]");
                            sbSqlEXE.AppendFormat(" ,[ActualEndTime]");
                            sbSqlEXE.AppendFormat(" ,[Count]");
                            sbSqlEXE.AppendFormat(" ,' '+'{0}'+'| '+'{1}'  AS [DailyCards] ", SBeginTime.ToString(), SEndTime.ToString());
                            sbSqlEXE.AppendFormat(" ,' '+'{0}'+'| '+'{1}'  AS [EmpRankCards]", SBeginTime.ToString(), SEndTime.ToString());
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [CollectBegin]", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [CollectEnd]", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[IsAbnormal]");
                            sbSqlEXE.AppendFormat(" FROM [HRMDB].[dbo].[AttendanceRollcall] WITH (NOLOCK)");
                            sbSqlEXE.AppendFormat(" WHERE [Hours]>0 AND [EmployeeId]='{0}'", emp);
                            sbSqlEXE.AppendFormat(" ORDER BY [AttendanceRollcall].[Date] DESC ");
                            sbSqlEXE.AppendFormat(" ");
                            sbSqlEXE.AppendFormat(" INSERT INTO [HRMDB].[dbo].[AttendanceCollect] ([AttendanceCollectId],[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode],[Date],[Time],[IsManual],[Source],[IsUnusual],[UnusualTypeId],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId],[ApproveDate],[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId],[RepealOperationDate],[RepealUserId],[StateId],[IsFromEss],[IsForAttendance] )");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1  NEWID() AS [AttendanceCollectId]", guid2.ToString());
                            sbSqlEXE.AppendFormat(" ,[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [Date] ", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd")); 
                            sbSqlEXE.AppendFormat(" ,'{0}' AS [Time]", SBeginTime.ToString());
                            sbSqlEXE.AppendFormat(" ,[IsManual]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}'+' 000459 03'  AS [Source]", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[IsUnusual],[UnusualTypeId],[Remark]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [CreateDate]", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd")); ;
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [LastModifiedDate]", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [ApproveDate]", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}'+' '+CONVERT(varchar(100),[AttendanceCollect].[RepealOperationDate],114) AS [RepealOperationDate]", SBeginTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[RepealUserId],[StateId],[IsFromEss],[IsForAttendance]");
                            sbSqlEXE.AppendFormat(" FROM  [HRMDB].[dbo].[AttendanceCollect] WITH (NOLOCK)");
                            sbSqlEXE.AppendFormat(" WHERE CONVERT(varchar(100),[AttendanceCollect].[Date],114) >='08:00:00' AND CONVERT(varchar(100),[AttendanceCollect].[Date],114) <='09:00:00'");
                            sbSqlEXE.AppendFormat(" AND  [EmployeeId]='{0}'", emp);
                            sbSqlEXE.AppendFormat(" ORDER BY [AttendanceCollect].[Date] DESC ");
                            sbSqlEXE.AppendFormat(" ");
                            sbSqlEXE.AppendFormat(" INSERT INTO [HRMDB].[dbo].[AttendanceCollect] ([AttendanceCollectId],[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode],[Date],[Time],[IsManual],[Source],[IsUnusual],[UnusualTypeId],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId],[ApproveDate],[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId],[RepealOperationDate],[RepealUserId],[StateId],[IsFromEss],[IsForAttendance] )");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1 NEWID() AS  [AttendanceCollectId]", guid3.ToString());
                            sbSqlEXE.AppendFormat(" ,[MachineId],[MachineCode],[CardId],[CardCode],[EmployeeName],[EmployeeCode],[EmployeeId],[DepartmentName],[DepartmentId],[CostCenterId],[CostCenterCode]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [Date] ", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{0}' AS [Time]", SEndTime.ToString());
                            sbSqlEXE.AppendFormat(" ,[IsManual]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}'+' 000459 03'  AS [Source]", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[IsUnusual],[UnusualTypeId],[Remark]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [CreateDate]", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd")); ;
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [LastModifiedDate]", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[RepairId],[AttendanceTypeId],[OldLogIds],[AttendanceCollectLogId],[AssignReason],[OwnerId],[IsEss],[IsEF],[EssNo],[EssType],[ClassCode],[SubmitOperationDate],[SubmitUserId],[ConfirmOperationDate],[ConfirmUserId],[ApproveResultId],[FoundOperationDate],[FoundUserId]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}' AS [ApproveDate]", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[ApproveEmployeeId],[ApproveEmployeeName],[ApproveRemark],[ApproveOperationDate],[ApproveUserId]");
                            sbSqlEXE.AppendFormat(" ,'{1}'+' '+'{0}'+' '+CONVERT(varchar(100),[AttendanceCollect].[RepealOperationDate],114) AS [RepealOperationDate]", SEndTime.ToString(), dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[RepealUserId],[StateId],[IsFromEss],[IsForAttendance]");
                            sbSqlEXE.AppendFormat(" FROM  [HRMDB].[dbo].[AttendanceCollect] WITH (NOLOCK)");
                            sbSqlEXE.AppendFormat(" WHERE CONVERT(varchar(100),[AttendanceCollect].[Date],114) >='17:00:00'");
                            sbSqlEXE.AppendFormat(" AND  [EmployeeId]='{0}'", emp);
                            sbSqlEXE.AppendFormat(" ORDER BY [AttendanceCollect].[Date] DESC ");
                            sbSqlEXE.AppendFormat(" ");
                        }
                        else
                        {
                            sbSqlEXE.AppendFormat(" ");
                            sbSqlEXE.AppendFormat(" INSERT INTO [HRMDB].[dbo].[AttendanceRollcall] ([AttendanceRollcallId],[EmployeeId],[Date],[BeginTime],[EndTime],[AttendanceRankId],[AttendanceTypeId],[Hours],[QuartersHours],[QuartersHoursUnit],[IsConfirm],[OperationDate],[UserId],[Recover],[Remark],[CreateDate],[LastModifiedDate],[CreateBy],[LastModifiedBy],[CorporationId],[Flag],[AssignReason],[OwnerId],[VirObjectId],[ActualBeginTime],[ActualEndTime],[Count],[DailyCards],[EmpRankCards],[CollectBegin],[CollectEnd],[IsAbnormal]) ");
                            sbSqlEXE.AppendFormat(" SELECT TOP 1 NEWID() AS [AttendanceRollcallId]", guid4.ToString());
                            sbSqlEXE.AppendFormat(" ,[EmployeeId]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 00:00:00.000' AS [Date]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{0}'+' '+CONVERT(varchar(100),[AttendanceRollcall].[BeginTime],114) AS [BeginTime] ", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,'{0}'+' '+CONVERT(varchar(100),[AttendanceRollcall].[EndTime],114) AS [EndTime]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[AttendanceRankId]");
                            sbSqlEXE.AppendFormat(" ,[AttendanceTypeId]");
                            sbSqlEXE.AppendFormat(" ,[Hours]");
                            sbSqlEXE.AppendFormat(" ,[QuartersHours]");
                            sbSqlEXE.AppendFormat(" ,[QuartersHoursUnit]");
                            sbSqlEXE.AppendFormat(" ,[IsConfirm]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 09:30:00.000' AS [OperationDate]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[UserId]");
                            sbSqlEXE.AppendFormat(" ,[Recover]");
                            sbSqlEXE.AppendFormat(" ,[Remark]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 09:30:00.000' AS [CreateDate]", dateTimePicker4.Value.ToString("yyyy-MM-dd")); ;
                            sbSqlEXE.AppendFormat(" ,'{0}'+' 09:30:00.000' AS [LastModifiedDate]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[CreateBy]");
                            sbSqlEXE.AppendFormat(" ,[LastModifiedBy]");
                            sbSqlEXE.AppendFormat(" ,[CorporationId]");
                            sbSqlEXE.AppendFormat(" ,[Flag]");
                            sbSqlEXE.AppendFormat(" ,[AssignReason]");
                            sbSqlEXE.AppendFormat(" ,[OwnerId]");
                            sbSqlEXE.AppendFormat(" ,[VirObjectId]");
                            sbSqlEXE.AppendFormat(" ,[ActualBeginTime]");
                            sbSqlEXE.AppendFormat(" ,[ActualEndTime]");
                            sbSqlEXE.AppendFormat(" ,[Count]");
                            sbSqlEXE.AppendFormat(" ,[DailyCards]");
                            sbSqlEXE.AppendFormat(" ,[EmpRankCards]");
                            sbSqlEXE.AppendFormat(" ,'{0}'+' '+CONVERT(varchar(100),[AttendanceRollcall].[CollectBegin],114) AS [CollectBegin]", dateTimePicker4.Value.ToString("yyyy-MM-dd")); 
                            sbSqlEXE.AppendFormat(" ,'{0}'+' '+CONVERT(varchar(100),[AttendanceRollcall].[CollectEnd],114) AS [CollectEnd]", dateTimePicker4.Value.ToString("yyyy-MM-dd"));
                            sbSqlEXE.AppendFormat(" ,[IsAbnormal]");
                            sbSqlEXE.AppendFormat(" FROM [HRMDB].[dbo].[AttendanceRollcall] WITH (NOLOCK)");
                            sbSqlEXE.AppendFormat(" WHERE [Hours]=0 AND [EmployeeId]='{0}'", emp);
                            sbSqlEXE.AppendFormat(" ORDER BY [AttendanceRollcall].[Date] DESC ");
                            sbSqlEXE.AppendFormat(" ");

                        }

                        cmd.Connection = sqlConn;
                        cmd.CommandTimeout = 60;
                        cmd.CommandText = sbSqlEXE.ToString();
                        cmd.Transaction = tran;
                        result = cmd.ExecuteNonQuery();

                        textBox1.Text = sbSqlEXE.ToString();
                        if (result == 0)
                        {
                            tran.Rollback();    //交易取消
                            INSERTLOG(NAME, "N");
                            MessageBox.Show("NG");
                            //label3.Text = DateTime.Now.ToString("yyyy/MM/dd") + "ADD FAIL";
                        }
                        else
                        {
                            tran.Commit();      //執行交易
                            INSERTLOG(NAME, "Y");
                            MessageBox.Show("OK");
                            //label3.Text = DateTime.Now.ToString("yyyy/MM/dd") + "ADD DONE";

                        }
                    }

                }




                sqlConn.Close();

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        public void INSERTLOG(string NAME, string STATUS)
        {
            sqlConn.Close();
            sqlConn.Open();
            tran = sqlConn.BeginTransaction();
            sbSqlEXE.AppendFormat(" INSERT INTO [TKSCHEDULE].[dbo].[LOG] ([NAME],[LOGTIME],[STATES]) VALUES ('{0}',GETDATE(),'{1}')", NAME, STATUS);

            cmd.Connection = sqlConn;
            cmd.CommandTimeout = 60;
            cmd.CommandText = sbSqlEXE.ToString();
            cmd.Transaction = tran;
            result = cmd.ExecuteNonQuery();

            textBox1.Text = sbSqlEXE.ToString();
            if (result == 0)
            {
                tran.Rollback();    //交易取消
            }
            else
            {
                tran.Commit();      //執行交易   
            }

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHTIME();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            SETAttendanceCollect();
            SETAttendanceRollcall();
            MessageBox.Show("已修改完成.");
             SEARCHTIME();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDHRCARD();
        }


        #endregion


    }
}
