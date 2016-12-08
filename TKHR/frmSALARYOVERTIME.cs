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
                sbSql.AppendFormat(@"  ,[ETIME] AS '打卡迄',[SHOURS] AS '打卡時數',[AHOURS] AS '核可時數");
                sbSql.AppendFormat(@"  ',[SUNITMONEY] AS '時薪',[AUNITMONEY] AS '核可金額' ");
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        #endregion


    }
}
