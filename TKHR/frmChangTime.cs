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


        #endregion

       
    }
}
