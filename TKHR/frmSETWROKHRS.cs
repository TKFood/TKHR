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
    public partial class frmSETWROKHRS : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlDataAdapter adapter1 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
        SqlDataAdapter adapter2 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder2 = new SqlCommandBuilder();
        SqlDataAdapter adapter3 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder3 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataTable dt = new DataTable();
        string SAVE;
        int result;
        string ID;

        public frmSETWROKHRS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void SEARCHAspNetRoles()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT [Name] AS '代號',[NormalizedName] AS '名稱',[Id],[ConcurrencyStamp]");
                sbSql.AppendFormat(@"  FROM [TKWEB].[dbo].[AspNetRoles]");
                sbSql.AppendFormat(@"  ORDER BY [Name]");
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                    SETNULL();
                }
                else
                {
                    dataGridView1.DataSource = ds1.Tables["TEMPds1"];
                    dataGridView1.AutoResizeColumns();
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
            if (dataGridView1.Rows.Count >= 1)
            {
                textBox1.Text = dataGridView1.CurrentRow.Cells["代號"].Value.ToString();
                textBox2.Text = dataGridView1.CurrentRow.Cells["名稱"].Value.ToString();
                textBox3.Text = dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
            }
            else
            {
                SETNULL();
            }
        }

        public void SETNULL()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCHAspNetRoles();
        }

        #endregion


    }
}
