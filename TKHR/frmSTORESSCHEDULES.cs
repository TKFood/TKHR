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
    public partial class frmSTORESSCHEDULES : Form
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
    

        Report report1 = new Report();
     
        public frmSTORESSCHEDULES()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
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


                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [ID]
                                    ,[NAMES] AS '姓名'
                                    ,[ORIBREAKDATES] AS '到期日'
                                    ,[NOWDATES] AS '起算日'
                                    ,[NEWBREAKDATES] AS '預計到期日'
                                    FROM [TKHR].[dbo].[STORESSCHEDULES]

                                    ORDER BY  [ID]

                                    ");

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

                        dataGridView1.Columns["姓名"].Width = 100;
                        dataGridView1.Columns["到期日"].Width = 100;
                        dataGridView1.Columns["起算日"].Width = 100;
                        dataGridView1.Columns["預計到期日"].Width = 160;
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

        public void Search2()
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


                sbSql.AppendFormat(@"  
                                  SELECT 
                                    [ID]
                                    ,[NAMES] AS '姓名'
                                    ,CONVERT(NVARCHAR,[ORIBREAKDATES],111) AS '到期日'
                   
                                    FROM [TKHR].[dbo].[STORESSCHEDULES]

                                    ORDER BY  [ID]

                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView2.DataSource = ds.Tables["TEMPds1"];
                        dataGridView2.AutoResizeColumns();

                        dataGridView2.Columns["姓名"].Width = 100;
                        dataGridView2.Columns["到期日"].Width = 100;
                        dataGridView2.Columns["起算日"].Width = 100;
                        dataGridView2.Columns["預計到期日"].Width = 160;
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
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    textBox2.Text = row.Cells["姓名"].Value.ToString().Trim();
                    textBox4.Text = row.Cells["姓名"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["ID"].Value.ToString().Trim();
                    textBox6.Text = row.Cells["ID"].Value.ToString().Trim();

                    dateTimePicker2.Value = Convert.ToDateTime(row.Cells["到期日"].Value.ToString());





                }
                else
                {
                    textBox2.Text = "";

                    textBox4.Text = "";
                    textBox5.Text = "";
                    textBox6.Text = "";

                }


            }
        }

        /// <summary>
        /// 將預計 開始日 存到STORESSCHEDULES 的 NOWDATES，用NOWDATES跟特休到期日排序
        /// </summary>
        /// <param name="NOWDATES"></param>
        public void UPDATESTORESSCHEDULESNOWDATES(string NOWDATES)
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

               
                sbSql.AppendFormat(@"  
                                    UPDATE [TKHR].[dbo].[STORESSCHEDULES]
                                    SET [NOWDATES]='{0}'
                                    ", NOWDATES);

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
        /// <summary>
        /// 如果 開始日 跟 到期日是負天數=過期，就把 到期日 延後1年
        /// </summary>
        public void UPDATESTORESSCHEDULESNEWBREAKDATES()
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                   UPDATE [TKHR].[dbo].[STORESSCHEDULES]
                                    SET [NEWBREAKDATES]=DATEADD(YEAR,1,[NEWBREAKDATES])
                                    WHERE DATEDIFF(DAY,[NOWDATES],[NEWBREAKDATES])<0

                                    " );

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
        /// <summary>
        /// 新增排休到 STORESSCHEDULESRESULTS
        /// </summary>
        /// <param name="SEQ"></param>
        public void ADDSTORESSCHEDULESRESULTS(string SEQ)
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  

                                    INSERT INTO [TKHR].[dbo].[STORESSCHEDULESRESULTS]
                                    (
                                    [SEQ]
                                    ,[NAMES]
                                    ,[NOWDATES]
                                    ,[NEWBREAKDATES]
                                    ,[DAYS]
                                    )

                                    SELECT 
                                    {0}
                                    ,[NAMES]
                                    ,[NOWDATES]
                                    ,[NEWBREAKDATES]
                                    ,DATEDIFF(DAY,[NOWDATES],[NEWBREAKDATES]) AS 'DAYS'
                                    FROM [TKHR].[dbo].[STORESSCHEDULES]
                                    ORDER BY DATEDIFF(DAY,[NOWDATES],[NEWBREAKDATES])

                                    ", SEQ);

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
        /// <summary>
        /// 清空STORESSCHEDULESRESULTS
        /// </summary>
        public void DELETESTORESSCHEDULESRESULTS()
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    DELETE [TKHR].[dbo].[STORESSCHEDULESRESULTS]

                                    ");

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
        /// <summary>
        /// 排休，依 開始日、排休次數
        /// </summary>
        public void SETSTORESSCHEDULESRESULTS(int COUNTS)
        {
            DateTime SDT = dateTimePicker1.Value;
            int NUMS = SEARCHSTORESSCHEDULES();

            DELETESTORESSCHEDULESRESULTS();

            for (int i = 1; i <= COUNTS; i++)
            {
                UPDATESTORESSCHEDULESNOWDATES(SDT.ToString("yyyyMMdd"));
                UPDATESTORESSCHEDULESNEWBREAKDATES();

                ADDSTORESSCHEDULESRESULTS(i.ToString());

                SDT = SDT.AddDays(7* NUMS);
            }

            //MessageBox.Show(COUNTS+" "+ SDT.ToString("yyyyMMdd"));
        }

        public void RESETSTORESSCHEDULESNEWBREAKDATES()
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    UPDATE [TKHR].[dbo].[STORESSCHEDULES]
                                    SET [NEWBREAKDATES]=[ORIBREAKDATES]

                                    ");

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

        public int SEARCHSTORESSCHEDULES()
        {
            int NUMS = 0;

            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();

            SqlTransaction tran;
            SqlCommand cmd = new SqlCommand();
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



                StringBuilder sbSql = new StringBuilder();
                sbSql.Clear();
                sbSqlQuery.Clear();
                ds.Clear();

              
                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(COUNT(*),0)  AS COUNTS FROM [TKHR].[dbo].[STORESSCHEDULES]

                                    ");

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds");
                sqlConn.Close();


                if (ds.Tables["TEMPds"].Rows.Count > 0)
                {
                    return Convert.ToInt32(ds.Tables["TEMPds"].Rows[0]["COUNTS"].ToString());
                }
                else
                {
                    return 0;
                }

            }
            catch
            {
                return 0;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void SETFASTREPORT()
        {
            StringBuilder SQL = new StringBuilder();
            report1 = new Report();

            report1.Load(@"REPORT\觀光排休表.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;


            //report1.Dictionary.Connections[0].ConnectionString = "server=192.168.1.105;database=TKPUR;uid=sa;pwd=dsc";

            TableDataSource Table = report1.GetDataSource("Table") as TableDataSource;

            SQL = SETFASETSQL();

            Table.SelectCommand = SQL.ToString(); ;

            report1.Preview = previewControl1;
            report1.Show();

        }

        public StringBuilder SETFASETSQL()
        {
            StringBuilder FASTSQL = new StringBuilder();
            StringBuilder STRQUERY = new StringBuilder();

          

            FASTSQL.AppendFormat(@"      
                                SELECT 
                                [ID]
                                ,[SEQ] AS '排休次數'
                                ,[NAMES] AS '姓名'
                                ,[NEWBREAKDATES] AS '預計到期日'
                                ,[NOWDATES] AS '排休起算日'
                                ,[NEWBREAKDATES] AS '預計到期日'
                                ,[DAYS] AS '差異天數'
                                FROM [TKHR].[dbo].[STORESSCHEDULESRESULTS]
                                ORDER BY [ID]
                                ");

            return FASTSQL;
        }

        public void UPDATESTORESSCHEDULES(string ID, string NAMES,string ORIBREAKDATES)
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                   
                                    UPDATE [TKHR].[dbo].[STORESSCHEDULES]
                                    SET [NAMES]='{1}',[ORIBREAKDATES]='{2}'
                                    WHERE [ID]='{0}'

                                    ",ID, NAMES, ORIBREAKDATES);

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

        public void ADDSTORESSCHEDULES(string NAMES, string ORIBREAKDATES)
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                   
                                   INSERT INTO [TKHR].[dbo].[STORESSCHEDULES]
                                    ([NAMES],[ORIBREAKDATES])
                                    VALUES
                                    ('{0}','{1}') 
				

                                    ", NAMES, ORIBREAKDATES);

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

        public void DELETESTORESSCHEDULES(string ID)
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



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                   
                                    DELETE [TKHR].[dbo].[STORESSCHEDULES]                                 
                                    WHERE [ID]='{0}'

                                    ", ID);

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
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            int n;
            string number = textBox1.Text.ToString().Trim();

            bool result = Int32.TryParse(number, out n);

            if (result)
            {
                RESETSTORESSCHEDULESNEWBREAKDATES();

                SETSTORESSCHEDULESRESULTS(n);

                SETFASTREPORT();
                MessageBox.Show("完成");
            }
            else
            {
                MessageBox.Show("排休次數不是整數");
            }

           

            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            RESETSTORESSCHEDULESNEWBREAKDATES();

            Search();
            MessageBox.Show("完成");
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Search2();

            MessageBox.Show("完成");
        }

        private void button4_Click(object sender, EventArgs e)
        {
            UPDATESTORESSCHEDULES(textBox5.Text,textBox2.Text, dateTimePicker2.Value.ToString("yyyy/MM/dd"));

            Search2();

            MessageBox.Show("完成");
        }

        private void button6_Click(object sender, EventArgs e)
        {
            ADDSTORESSCHEDULES(textBox3.Text,dateTimePicker3.Value.ToString("yyyy/MM/dd"));

            Search2();

            MessageBox.Show("完成");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("要刪除了?", "要刪除了?", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                if (!string.IsNullOrEmpty(textBox6.Text))
                {
                    DELETESTORESSCHEDULES(textBox6.Text);

                    Search2();

                    MessageBox.Show("完成");
                }

            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
        }
    }

    #endregion


}
