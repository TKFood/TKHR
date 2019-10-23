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
        SqlDataAdapter adapter4 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder4 = new SqlCommandBuilder();
        SqlDataAdapter adapter5 = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder5 = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5= new DataSet();
        DataTable dt = new DataTable();
        string SAVE;
        int result;
        string ID;

        string STATUSAspNetRoles;
        string STATUSWORKID;
        string RoleId;

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

        public void SEARCHWORKID()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                
                sbSql.AppendFormat(@"  SELECT [WORKID] AS '代號',[WORKNAME] AS '名稱'");
                sbSql.AppendFormat(@"  FROM [TKWEB].[dbo].[HRWORK]");
                sbSql.AppendFormat(@"  ORDER BY [WORKID]");
                sbSql.AppendFormat(@"  ");

                adapter2 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder2 = new SqlCommandBuilder(adapter2);
                sqlConn.Open();
                ds2.Clear();
                adapter2.Fill(ds2, "ds2");
                sqlConn.Close();


                if (ds2.Tables["ds2"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                    SETNULL2();
                }
                else
                {
                    dataGridView2.DataSource = ds2.Tables["ds2"];
                    dataGridView2.AutoResizeColumns();
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

        public void SEARCHAspNetRoles2()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@"  SELECT [Name] AS '代號',[NormalizedName] AS '名稱',[Id],[ConcurrencyStamp]");
                sbSql.AppendFormat(@"  FROM [TKWEB].[dbo].[AspNetRoles]");
                sbSql.AppendFormat(@"  ORDER BY [Name]");
                sbSql.AppendFormat(@"  ");


                adapter3 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder3 = new SqlCommandBuilder(adapter3);
                sqlConn.Open();
                ds3.Clear();
                adapter3.Fill(ds3, "ds3");
                sqlConn.Close();


                if (ds3.Tables["ds3"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                   
                }
                else
                {
                    dataGridView3.DataSource = ds3.Tables["ds3"];
                    dataGridView3.AutoResizeColumns();
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
        public void SEARCHAspNetUserRoles(string RoleId)
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();

               
                sbSql.AppendFormat(@"  SELECT [Name] AS '角色',[UserName]  AS '帳號',[UserId],[RoleId],[AspNetUsers].[Id],[AspNetRoles].[Id],[NormalizedName]");
                sbSql.AppendFormat(@"  FROM [TKWEB].[dbo].[AspNetUserRoles],[TKWEB].[dbo].[AspNetRoles],[TKWEB].[dbo].[AspNetUsers]");
                sbSql.AppendFormat(@"  WHERE [AspNetUserRoles].[UserId]=[AspNetUsers].[Id] ");
                sbSql.AppendFormat(@"  AND [AspNetUserRoles].[RoleId]=[AspNetRoles].[Id]");
                sbSql.AppendFormat(@"  AND [AspNetUserRoles].[RoleId]='{0}'",RoleId);
                sbSql.AppendFormat(@"  ");


                adapter4 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder4 = new SqlCommandBuilder(adapter4);
                sqlConn.Open();
                ds4.Clear();
                adapter4.Fill(ds4, "ds4");
                sqlConn.Close();


                if (ds4.Tables["ds4"].Rows.Count == 0)
                {
                    dataGridView4.DataSource = null;

                }
                else
                {
                    dataGridView4.DataSource = ds4.Tables["ds4"];
                    dataGridView4.AutoResizeColumns();
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

        public void SEARCHAspNetUsers()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                                
                sbSql.AppendFormat(@"  SELECT [UserName] AS '帳號',[Id]");
                sbSql.AppendFormat(@"  FROM [TKWEB].[dbo].[AspNetUsers]");
                sbSql.AppendFormat(@"  WHERE [Id] NOT IN (");
                sbSql.AppendFormat(@"  SELECT [UserId]");
                sbSql.AppendFormat(@"  FROM [TKWEB].[dbo].[AspNetUserRoles],[TKWEB].[dbo].[AspNetRoles],[TKWEB].[dbo].[AspNetUsers]");
                sbSql.AppendFormat(@"  WHERE [AspNetUserRoles].[UserId]=[AspNetUsers].[Id]");
                sbSql.AppendFormat(@"  AND [AspNetUserRoles].[RoleId]=[AspNetRoles].[Id]");
                sbSql.AppendFormat(@"  AND [AspNetUserRoles].[RoleId]='{0}'",RoleId);
                sbSql.AppendFormat(@"  )");
                sbSql.AppendFormat(@"  ORDER BY [UserName]");
                sbSql.AppendFormat(@"  ");


                adapter5 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder5 = new SqlCommandBuilder(adapter5);
                sqlConn.Open();
                ds5.Clear();
                adapter5.Fill(ds5, "ds5");
                sqlConn.Close();


                if (ds5.Tables["ds5"].Rows.Count == 0)
                {
                    dataGridView5.DataSource = null;

                }
                else
                {
                    dataGridView5.DataSource = ds5.Tables["ds5"];
                    dataGridView5.AutoResizeColumns();
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


        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count >= 1)
            {
                textBox4.Text = dataGridView2.CurrentRow.Cells["代號"].Value.ToString();
                textBox5.Text = dataGridView2.CurrentRow.Cells["名稱"].Value.ToString();
            }
            else
            {
                SETNULL2();
            }

        }

        private void dataGridView3_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView3.Rows.Count >= 1)
            {
                RoleId= dataGridView3.CurrentRow.Cells["Id"].Value.ToString();
                SEARCHAspNetUserRoles(RoleId);

                textBox6.Text = dataGridView3.CurrentRow.Cells["代號"].Value.ToString();
                textBox8.Text = dataGridView3.CurrentRow.Cells["Id"].Value.ToString();
            }
            else
            {
                SETNULL2();
            }

        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView4.Rows.Count >= 1)
            {
                textBox10.Text= dataGridView4.CurrentRow.Cells["角色"].Value.ToString();
                textBox11.Text = dataGridView4.CurrentRow.Cells["帳號"].Value.ToString();
                textBox12.Text = dataGridView4.CurrentRow.Cells["RoleId"].Value.ToString();
                textBox13.Text = dataGridView4.CurrentRow.Cells["UserId"].Value.ToString();
               
            }
            else
            {
                SETNULL3();
            }
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView5.Rows.Count >= 1)
            {
                textBox7.Text = dataGridView5.CurrentRow.Cells["帳號"].Value.ToString();
                textBox9.Text = dataGridView5.CurrentRow.Cells["Id"].Value.ToString();

            }
            else
            {
                
            }
        }
        public void SETNULL()
        {
            textBox1.Text = null;
            textBox2.Text = null;
            textBox3.Text = null;

            textBox1.ReadOnly = true;
            textBox2.ReadOnly = true;
            //textBox3.ReadOnly = true;
        }
        public void SETREADONLY()
        {
            textBox1.ReadOnly = false;
            textBox2.ReadOnly = false;
            //textBox3.ReadOnly = false;
        }

        public void SETNULL2()
        {
            textBox6.Text = null;
            textBox8.Text = null;            

           
        }
        public void SETREADONLY2()
        {
            textBox4.ReadOnly = false;
            textBox5.ReadOnly = false;
            
        }

        public void SETNULL3()
        {
            textBox10.Text = null;
            textBox11.Text = null;
            textBox12.Text = null;
            textBox13.Text = null;

        }

        public void ADDAspNetRoles()
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
               
                sbSql.AppendFormat(" INSERT INTO [TKWEB].[dbo].[AspNetRoles]");
                sbSql.AppendFormat(" ([Id],[Name],[NormalizedName],[ConcurrencyStamp])");
                sbSql.AppendFormat(" VALUES (NEWID(),'{0}','{1}',NEWID())",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("FAIL");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("OK");

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

        public void UPDATEAspNetRoles()
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

                sbSql.AppendFormat(" UPDATE [TKWEB].[dbo].[AspNetRoles]");
                sbSql.AppendFormat(" SET [Name]='{0}',[NormalizedName]='{1}'",textBox1.Text,textBox2.Text);
                sbSql.AppendFormat(" WHERE [Id]='{0}'",textBox3.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("FAIL");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("OK");

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

        public void ADDWORKID()
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

                sbSql.AppendFormat(" INSERT INTO [TKWEB].[dbo].[HRWORK]");
                sbSql.AppendFormat(" ([WORKID],[WORKNAME])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}')",textBox4.Text,textBox5.Text);
                sbSql.AppendFormat(" ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("FAIL");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("OK");

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

        public void UPDATEWORKID()
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

                sbSql.AppendFormat(" UPDATE [TKWEB].[dbo].[HRWORK]");
                sbSql.AppendFormat(" SET [WORKNAME]='{0}'", textBox5.Text);
                sbSql.AppendFormat(" WHERE [WORKID]='{0}'", textBox4.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("FAIL");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("OK");

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
        public void ADDAspNetUserRoles()
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

                sbSql.AppendFormat(" INSERT INTO [TKWEB].[dbo].[AspNetUserRoles]");
                sbSql.AppendFormat(" ([UserId],[RoleId])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}')",textBox9.Text,textBox8.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("FAIL");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("OK");

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
        public void DELETEAspNetUserRoles()
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

                sbSql.AppendFormat(" DELETE [TKWEB].[dbo].[AspNetUserRoles]");
                sbSql.AppendFormat(" WHERE [UserId]='{0}' AND [RoleId]='{1}'",textBox13.Text,textBox12.Text);
                sbSql.AppendFormat(" ");

                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                    MessageBox.Show("FAIL");
                }
                else
                {
                    tran.Commit();      //執行交易  
                    MessageBox.Show("OK");

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
            SEARCHAspNetRoles();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SETNULL();
            SETREADONLY();
            STATUSAspNetRoles = "ADD";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETREADONLY();
            STATUSAspNetRoles = "EDIT";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if(STATUSAspNetRoles.Equals("ADD"))
            {
                ADDAspNetRoles();
            }
            else if(STATUSAspNetRoles.Equals("EDIT"))
            {
                UPDATEAspNetRoles();
            }

            STATUSAspNetRoles = null;

            SETNULL();
            SEARCHAspNetRoles();
            
        }
        private void button8_Click(object sender, EventArgs e)
        {
            SEARCHWORKID();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SETNULL2();
            SETREADONLY2();
            STATUSWORKID = "ADD";
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SETREADONLY2();
            STATUSWORKID = "EDIT";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (STATUSWORKID.Equals("ADD"))
            {
                ADDWORKID();
            }
            else if (STATUSWORKID.Equals("EDIT"))
            {
                UPDATEWORKID();
            }

            STATUSWORKID = null;

            SETNULL2();
            SEARCHWORKID();
        }
        private void button9_Click(object sender, EventArgs e)
        {
            SEARCHAspNetRoles2();
        }
        private void button10_Click(object sender, EventArgs e)
        {
            SEARCHAspNetUsers();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            ADDAspNetUserRoles();

            SEARCHAspNetUserRoles(RoleId);
            SEARCHAspNetUsers();
        }
        
        private void button14_Click(object sender, EventArgs e)
        {
            DELETEAspNetUserRoles();

            SEARCHAspNetUserRoles(RoleId);
            SEARCHAspNetUsers();
        }



        #endregion


    }
}
