﻿using System;
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
using TKITDLL;

namespace TKHR
{
    public partial class frmEMPLOYEELEAVE : Form
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
        DataSet ds3 = new DataSet();
        DataTable dt = new DataTable();
        string SAVE;
        int result;
        string ID;
        public frmEMPLOYEELEAVE()
        {
            InitializeComponent();
            comboBox1load();
            comboBox2load();
            comboBox3load();
            comboBox4load();
            comboBox5load();
            comboBox6load();
            comboBox7load();
            comboBox8load();
            comboBox9load();
            comboBox10load();
            comboBox11load();
            comboBox12load();

        }

        #region FUNCTION
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }


        public void comboBox1load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '11%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "NAME";
            comboBox1.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox2load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '12%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "NAME";
            comboBox2.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox3load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '13%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "NAME";
            comboBox3.DisplayMember = "NAME";
            sqlConn.Close();


        }

        public void comboBox4load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '14%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox4.DataSource = dt.DefaultView;
            comboBox4.ValueMember = "NAME";
            comboBox4.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox5load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '15%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox5.DataSource = dt.DefaultView;
            comboBox5.ValueMember = "NAME";
            comboBox5.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox6load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '3%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox6.DataSource = dt.DefaultView;
            comboBox6.ValueMember = "NAME";
            comboBox6.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox7load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '21%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox7.DataSource = dt.DefaultView;
            comboBox7.ValueMember = "NAME";
            comboBox7.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox8load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '22%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox8.DataSource = dt.DefaultView;
            comboBox8.ValueMember = "NAME";
            comboBox8.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox9load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '23%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox9.DataSource = dt.DefaultView;
            comboBox9.ValueMember = "NAME";
            comboBox9.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox10load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '24%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox10.DataSource = dt.DefaultView;
            comboBox10.ValueMember = "NAME";
            comboBox10.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox11load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '25%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox11.DataSource = dt.DefaultView;
            comboBox11.ValueMember = "NAME";
            comboBox11.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void comboBox12load()
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            StringBuilder Sequel = new StringBuilder();
            Sequel.AppendFormat(@"SELECT  [ID],[NAME] FROM [TKHR].[dbo].[EMPLOYEELEAVESELECT] WHERE [ID] LIKE '3%' ORDER BY [ID]");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("NAME", typeof(string));
            da.Fill(dt);
            comboBox12.DataSource = dt.DefaultView;
            comboBox12.ValueMember = "NAME";
            comboBox12.DisplayMember = "NAME";
            sqlConn.Close();


        }
        public void SEARCH()
        {
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT TOP 1 ");
                sbSql.AppendFormat(@"  [Employee].[CODE]");
                sbSql.AppendFormat(@"  ,[Employee].[Date]");
                sbSql.AppendFormat(@"  ,[Employee].[CnName] ");
                sbSql.AppendFormat(@"  ,[Employee].[Telephone]");
                sbSql.AppendFormat(@"  ,[Employee].[Location]");
                sbSql.AppendFormat(@"  ,CASE [Employee].[GenderId] WHEN 'Gender_001' THEN '男' ELSE '女' END  AS [GenderId]");
                sbSql.AppendFormat(@"  ,[Job].[NAME]  AS Job");
                sbSql.AppendFormat(@"  ,[Department].[NAME] AS Department ");
                sbSql.AppendFormat(@"  FROM [HRMDB].[dbo].[Employee],[HRMDB].[dbo].[Job],[HRMDB].[dbo].[Department]");
                sbSql.AppendFormat(@"  WHERE [Employee].[JobId]=[Job].[JobId]");
                sbSql.AppendFormat(@"  AND [Department].[DepartmentId]=[Employee].[DepartmentId]");
                sbSql.AppendFormat(@"  AND  [Employee].[CODE]='{0}'",textBox1.Text.ToString());
                sbSql.AppendFormat(@"  ");


                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds1.Clear();
                adapter.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    SETNULL();
                }
                else
                {
                    if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        textBox2.Text = ds1.Tables["TEMPds1"].Rows[0]["CnName"].ToString();
                        textBox3.Text = ds1.Tables["TEMPds1"].Rows[0]["Department"].ToString();
                        textBox4.Text = ds1.Tables["TEMPds1"].Rows[0]["Job"].ToString();                        
                        textBox6.Text = ds1.Tables["TEMPds1"].Rows[0]["Telephone"].ToString();
                        textBox7.Text = ds1.Tables["TEMPds1"].Rows[0]["Location"].ToString();
                        dateTimePicker1.Value=Convert.ToDateTime(ds1.Tables["TEMPds1"].Rows[0]["Date"].ToString());
                        comboBox13.Text = ds1.Tables["TEMPds1"].Rows[0]["GenderId"].ToString();


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
        public void SEARCHEMPLOYEELEAVE()
        {
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT [NO] AS '編號'");
                sbSql.AppendFormat(@"  ,[CODE] AS '工號'");
                sbSql.AppendFormat(@"  ,[Date] AS '填表日'");
                sbSql.AppendFormat(@"  ,[CnName] AS '姓名'");
                sbSql.AppendFormat(@"  ,[Telephone] AS '電話'");
                sbSql.AppendFormat(@"  ,[Location] AS '地址'");
                sbSql.AppendFormat(@"  ,[GenderId] AS '性別'");
                sbSql.AppendFormat(@"  ,[Job] AS '職稱'");
                sbSql.AppendFormat(@"  ,[Department] AS '部門'");
                sbSql.AppendFormat(@"  ,[EVAWORK1] AS '工作量'");
                sbSql.AppendFormat(@"  ,[EVAWORK2] AS '困難度'");
                sbSql.AppendFormat(@"  ,[EVAWORK3] AS '適應度'");
                sbSql.AppendFormat(@"  ,[EVAWORK4] AS '順暢度'");
                sbSql.AppendFormat(@"  ,[EVAWORK5] AS '工作程序'");
                sbSql.AppendFormat(@"  ,[EVAWORKSUG] AS '工作建議'");
                sbSql.AppendFormat(@"  ,[EVAWORK1REVIWER] AS '工作量-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK2REVIWER] AS '困難度-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK3REVIWER] AS '適應度-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK4REVIWER] AS '順暢度-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK5REVIWER] AS '工作程序-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORKSUGREVIWER] AS '面談結論'");
                sbSql.AppendFormat(@"  ,[REASON] AS '離職原因'");
                sbSql.AppendFormat(@"  ,[REASONSUG] AS '對公司建議'");
                sbSql.AppendFormat(@"  ,[REASONREVIWER] AS '離職原因-面談'");
                sbSql.AppendFormat(@"  ,[REASONSUGREVIWER] AS '面談總結論'");
                sbSql.AppendFormat(@"  ,[COMMENT] AS '簽核意見'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKHR].[dbo].[EMPLOYEELEAVE]");
                sbSql.AppendFormat(@"  WHERE [CODE]='{0}'", textBox1.Text);
                sbSql.AppendFormat(@"  ORDER BY [NO] ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                    SETNULLDETAIL();
                }
                else
                {
                    if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds2.Tables["TEMPds2"];
                        dataGridView1.AutoResizeColumns();
                       
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
        public void SETNULL()
        {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            
            textBox6.Text = "";
            textBox7.Text = "";
            comboBox13.Text = "男";
        }
        public void SETNULLDETAIL()
        {
            comboBox1.Text = "";
            comboBox2.Text = "";
            comboBox3.Text = "";
            comboBox4.Text = "";
            comboBox5.Text = "";
            comboBox6.Text = "";
            comboBox7.Text = "";
            comboBox8.Text = "";
            comboBox9.Text = "";
            comboBox10.Text = "";
            comboBox11.Text = "";
            comboBox12.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";

        }
        public void UPDATE()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
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
               
                sbSql.AppendFormat(" UPDATE [TKHR].[dbo].[EMPLOYEELEAVE]");
                sbSql.AppendFormat(" SET [NO]='{0}'",textBox12.Text);
                sbSql.AppendFormat(" ,[CODE]='{0}'", textBox1.Text);
                sbSql.AppendFormat(" ,[Date]='{0}'", dateTimePicker2.Value.ToString("yyy/MM/dd"));
                sbSql.AppendFormat(" ,[CnName]='{0}'", textBox2.Text);
                sbSql.AppendFormat(" ,[Telephone]='{0}'", textBox6.Text);
                sbSql.AppendFormat(" ,[Location]='{0}'", textBox7.Text);
                sbSql.AppendFormat(" ,[GenderId]='{0}'", comboBox13.Text.ToString());
                sbSql.AppendFormat(" ,[Job]='{0}'", textBox4.Text);
                sbSql.AppendFormat(" ,[Department]='{0}'", textBox3.Text);
                sbSql.AppendFormat(" ,[EVAWORK1]='{0}'", comboBox1.Text);
                sbSql.AppendFormat(" ,[EVAWORK2]='{0}'", comboBox2.Text);
                sbSql.AppendFormat(" ,[EVAWORK3]='{0}'", comboBox3.Text);
                sbSql.AppendFormat(" ,[EVAWORK4]='{0}'", comboBox4.Text);
                sbSql.AppendFormat(" ,[EVAWORK5]='{0}'", comboBox5.Text);
                sbSql.AppendFormat(" ,[EVAWORKSUG]='{0}'", textBox8.Text);
                sbSql.AppendFormat(" ,[EVAWORK1REVIWER]='{0}'", comboBox7.Text);
                sbSql.AppendFormat(" ,[EVAWORK2REVIWER]='{0}'", comboBox8.Text);
                sbSql.AppendFormat(" ,[EVAWORK3REVIWER]='{0}'", comboBox9.Text);
                sbSql.AppendFormat(" ,[EVAWORK4REVIWER]='{0}'", comboBox10.Text);
                sbSql.AppendFormat(" ,[EVAWORK5REVIWER]='{0}'", comboBox11.Text);
                sbSql.AppendFormat(" ,[EVAWORKSUGREVIWER]='{0}'", textBox9.Text);
                sbSql.AppendFormat(" ,[REASON]='{0}'", comboBox6.Text);
                sbSql.AppendFormat(" ,[REASONSUG]='{0}'", textBox10.Text);
                sbSql.AppendFormat(" ,[REASONREVIWER]='{0}'", comboBox12.Text);
                sbSql.AppendFormat(" ,[REASONSUGREVIWER]='{0}'", textBox11.Text);
                sbSql.AppendFormat(" ,[COMMENT]=''");
                sbSql.AppendFormat(" WHERE [ID]='{0}'",ID);
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

        public void ADD()
        {
            try
            {
                //add ZWAREWHOUSEPURTH
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
                sbSql.AppendFormat(" INSERT INTO [TKHR].[dbo].[EMPLOYEELEAVE]");
                sbSql.AppendFormat(" ([ID],[NO],[CODE],[Date],[CnName]");
                sbSql.AppendFormat(" ,[Telephone],[Location],[GenderId],[Job],[Department]");
                sbSql.AppendFormat(" ,[EVAWORK1],[EVAWORK2],[EVAWORK3],[EVAWORK4],[EVAWORK5]");
                sbSql.AppendFormat(" ,[EVAWORKSUG],[EVAWORK1REVIWER],[EVAWORK2REVIWER],[EVAWORK3REVIWER],[EVAWORK4REVIWER]");
                sbSql.AppendFormat(" ,[EVAWORK5REVIWER],[EVAWORKSUGREVIWER],[REASON],[REASONSUG],[REASONREVIWER]");
                sbSql.AppendFormat(" ,[REASONSUGREVIWER],[COMMENT])");
                sbSql.AppendFormat(" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}','{17}','{18}','{19}','{20}','{21}','{22}','{23}','{24}','{25}','{26}')", Guid.NewGuid(), textBox12.Text, textBox1.Text, dateTimePicker2.Value.ToString("yyy/MM/dd"), textBox2.Text, textBox6.Text, textBox7.Text, comboBox13.Text.ToString(), textBox4.Text, textBox3.Text, comboBox1.Text, comboBox2.Text, comboBox3.Text, comboBox4.Text, comboBox5.Text, textBox8.Text, comboBox7.Text, comboBox8.Text, comboBox9.Text, comboBox10.Text, comboBox11.Text, textBox9.Text, comboBox6.Text, textBox10.Text, comboBox12.Text, textBox11.Text,"");
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
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count >= 1)
            {
                comboBox1.Text = dataGridView1.CurrentRow.Cells["工作量"].Value.ToString();
                comboBox2.Text = dataGridView1.CurrentRow.Cells["困難度"].Value.ToString();
                comboBox3.Text = dataGridView1.CurrentRow.Cells["適應度"].Value.ToString();
                comboBox4.Text = dataGridView1.CurrentRow.Cells["順暢度"].Value.ToString();
                comboBox5.Text = dataGridView1.CurrentRow.Cells["工作程序"].Value.ToString();
                comboBox6.Text = dataGridView1.CurrentRow.Cells["離職原因"].Value.ToString();
                comboBox7.Text = dataGridView1.CurrentRow.Cells["工作量-面談"].Value.ToString();
                comboBox8.Text = dataGridView1.CurrentRow.Cells["困難度-面談"].Value.ToString();
                comboBox9.Text = dataGridView1.CurrentRow.Cells["適應度-面談"].Value.ToString();
                comboBox10.Text = dataGridView1.CurrentRow.Cells["順暢度-面談"].Value.ToString();
                comboBox11.Text = dataGridView1.CurrentRow.Cells["工作程序-面談"].Value.ToString();
                comboBox12.Text = dataGridView1.CurrentRow.Cells["離職原因-面談"].Value.ToString();
                textBox8.Text = dataGridView1.CurrentRow.Cells["工作建議"].Value.ToString();
                textBox9.Text = dataGridView1.CurrentRow.Cells["面談結論"].Value.ToString();
                textBox10.Text = dataGridView1.CurrentRow.Cells["對公司建議"].Value.ToString();
                textBox11.Text = dataGridView1.CurrentRow.Cells["面談總結論"].Value.ToString();

                ID= dataGridView1.CurrentRow.Cells["ID"].Value.ToString();
            }
        }
        public void SEARCHEMPLOYEELEAVEREPORT()
        {
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


                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds1.Clear();

                sbSql.AppendFormat(@"  SELECT [NO] AS '編號'");
                sbSql.AppendFormat(@"  ,[CODE] AS '工號'");
                sbSql.AppendFormat(@"  ,[Date] AS '填表日'");
                sbSql.AppendFormat(@"  ,[CnName] AS '姓名'");
                sbSql.AppendFormat(@"  ,[Telephone] AS '電話'");
                sbSql.AppendFormat(@"  ,[Location] AS '地址'");
                sbSql.AppendFormat(@"  ,[GenderId] AS '性別'");
                sbSql.AppendFormat(@"  ,[Job] AS '職稱'");
                sbSql.AppendFormat(@"  ,[Department] AS '部門'");
                sbSql.AppendFormat(@"  ,[EVAWORK1] AS '工作量'");
                sbSql.AppendFormat(@"  ,[EVAWORK2] AS '困難度'");
                sbSql.AppendFormat(@"  ,[EVAWORK3] AS '適應度'");
                sbSql.AppendFormat(@"  ,[EVAWORK4] AS '順暢度'");
                sbSql.AppendFormat(@"  ,[EVAWORK5] AS '工作程序'");
                sbSql.AppendFormat(@"  ,[EVAWORKSUG] AS '工作建議'");
                sbSql.AppendFormat(@"  ,[EVAWORK1REVIWER] AS '工作量-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK2REVIWER] AS '困難度-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK3REVIWER] AS '適應度-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK4REVIWER] AS '順暢度-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORK5REVIWER] AS '工作程序-面談'");
                sbSql.AppendFormat(@"  ,[EVAWORKSUGREVIWER] AS '面談結論'");
                sbSql.AppendFormat(@"  ,[REASON] AS '離職原因'");
                sbSql.AppendFormat(@"  ,[REASONSUG] AS '對公司建議'");
                sbSql.AppendFormat(@"  ,[REASONREVIWER] AS '離職原因-面談'");
                sbSql.AppendFormat(@"  ,[REASONSUGREVIWER] AS '面談總結論'");
                sbSql.AppendFormat(@"  ,[COMMENT] AS '簽核意見'");
                sbSql.AppendFormat(@"  ,[ID] ");
                sbSql.AppendFormat(@"  FROM [TKHR].[dbo].[EMPLOYEELEAVE]");
                sbSql.AppendFormat(@"  WHERE [Date]>='{0}' AND [Date]<='{1}'", dateTimePicker3.Value.ToString("yyyy/MM/dd"), dateTimePicker4.Value.ToString("yyyy/MM/dd"));
                sbSql.AppendFormat(@"  ORDER BY [NO] ");
                sbSql.AppendFormat(@"  ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds3.Clear();
                adapter.Fill(ds3, "TEMPds3");
                sqlConn.Close();


                if (ds3.Tables["TEMPds3"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                    SETNULLDETAIL();
                }
                else
                {
                    if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds3.Tables["TEMPds3"];
                        dataGridView2.AutoResizeColumns();

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
        public void ExcelExport()
        {
            SEARCHEMPLOYEELEAVEREPORT();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds3.Tables["TEMPds3"];
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
            foreach (DataGridViewRow dr in this.dataGridView2.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString());
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString());
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString());
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString());
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString());
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString());
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString());
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString());
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString());
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString());
                ws.GetRow(j + 1).CreateCell(15).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString());
                ws.GetRow(j + 1).CreateCell(16).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[16].ToString());
                ws.GetRow(j + 1).CreateCell(17).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[17].ToString());
                ws.GetRow(j + 1).CreateCell(18).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[18].ToString());
                ws.GetRow(j + 1).CreateCell(19).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[19].ToString());
                ws.GetRow(j + 1).CreateCell(20).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[20].ToString());
                ws.GetRow(j + 1).CreateCell(21).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[21].ToString());
                ws.GetRow(j + 1).CreateCell(22).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[22].ToString());
                ws.GetRow(j + 1).CreateCell(23).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[23].ToString());
                ws.GetRow(j + 1).CreateCell(24).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[24].ToString());
                ws.GetRow(j + 1).CreateCell(25).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[25].ToString());
                ws.GetRow(j + 1).CreateCell(26).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[26].ToString());




                j++;
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
            filename.AppendFormat(@"c:\temp\離職記錄{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SEARCH();
            SEARCHEMPLOYEELEAVE();
        }
        private void button2_Click(object sender, EventArgs e)
        {

            if (SAVE.ToString().Equals("UPDATE"))
            {
                UPDATE();
            }
            else if (SAVE.ToString().Equals("ADD"))
            {
                ADD();
            }

            button2.Visible = false;
            button3.Visible = true;
            button4.Visible = true;
            SEARCHEMPLOYEELEAVE();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SAVE = "ADD";
            button2.Visible = true;
            button3.Visible = false;
            button4.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SAVE = "UPDATE";
            button2.Visible = true;
            button3.Visible = false;
            button4.Visible = false;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHEMPLOYEELEAVEREPORT();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        #endregion


    }
}
