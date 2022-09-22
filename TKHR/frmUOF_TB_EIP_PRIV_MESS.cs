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
using FastReport;
using FastReport.Data;
using TKITDLL;
using FastReport.Export.Pdf;
using System.Net.Mail;
using System.Net.Mime;

namespace TKHR
{
    public partial class frmUOF_TB_EIP_PRIV_MESS : Form
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
        string tablename = null;

        int result;

        public frmUOF_TB_EIP_PRIV_MESS()
        {
            InitializeComponent();

            textBox2.Text = "b6f50a95-17ec-47f2-b842-4ad12512b431";
        }

        #region FUNCTION
        public void SearchUOFTB_EIP_PRIV_MESS(string ACCOUNT)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

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



                sbSql.Clear();
                sbSqlQuery.Clear();

                if(!string.IsNullOrEmpty(ACCOUNT))
                {
                    sbSqlQuery.AppendFormat(@" 
                                            AND ACCOUNT LIKE '{0}%'
                                            ", ACCOUNT);
                }

                sbSql.AppendFormat(@"  
                                   SELECT ACCOUNT AS '工號',NAME AS '姓名',USER_GUID
                                    FROM [192.168.1.223].[UOF].dbo.TB_EB_USER
                                    WHERE 1=1
                                    AND IS_SUSPENDED='0'
                                    AND(ACCOUNT LIKE '0%' OR ACCOUNT LIKE '1%' OR ACCOUNT LIKE '2%' OR ACCOUNT LIKE '3%' OR ACCOUNT LIKE '4%' OR ACCOUNT LIKE '5%' OR ACCOUNT LIKE '6%' OR ACCOUNT LIKE '7%' OR ACCOUNT LIKE '8%' OR ACCOUNT LIKE '9%' )
                                    {0}
                                    ORDER BY ACCOUNT

                                    ", sbSqlQuery.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();

                        dataGridView1.Columns["工號"].Width = 100;
                        dataGridView1.Columns["姓名"].Width = 100;
                      
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;
            textBox5.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox4.Text = row.Cells["USER_GUID"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["姓名"].Value.ToString().Trim();



                }
                else
                {
                    textBox4.Text = null;
                    textBox5.Text = null;
                }
            }
        }

        public void ADDToUOF_TB_EIP_PRIV_MESS(string MESSAGE_TO, string MESSAGE_FROM, string CONTENT)
        { 


        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SearchUOFTB_EIP_PRIV_MESS(textBox1.Text.ToString());
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }
        #endregion


    }
}
