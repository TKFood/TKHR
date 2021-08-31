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
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Globalization;
using FastReport;
using FastReport.Data;


namespace TKHR
{
    public partial class frmTBEIPPUNCH : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        int result;

        string START = "N";



        public frmTBEIPPUNCH()
        {
            InitializeComponent();

            textBox2.Text = "08:10";


            textBox1.Text = @"D:\SCSHR\Card";

            timer1.Enabled = true;
            timer1.Interval = 1000 * 60;
            //timer1.Interval = 1000 ;
            timer1.Start();
        }

        #region FUNCTION

 
        public DataTable SERACHTB_EIP_PUNCH(string STARTTIME,string ENDTIME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
               

                sbSql.AppendFormat(@"  
                                   SELECT [ACCOUNT]+' '+CONVERT(NVARCHAR,[MODIFY_DATE],112)+' '+SUBSTRING((REPLACE((CONVERT(varchar(100), [MODIFY_DATE], 108)),':','')),1,4) AS DATAS,[MODIFY_DATE]
                                    FROM [UOF].[dbo].[TB_EIP_PUNCH_LOG],[UOF].[dbo].[TB_EB_USER]
                                    WHERE [TB_EIP_PUNCH_LOG].USER_GUID=[TB_EB_USER].USER_GUID
                                    AND CONVERT(NVARCHAR,[MODIFY_DATE],112)+SUBSTRING((REPLACE((CONVERT(varchar(100), [MODIFY_DATE], 108)),':','')),1,4)>='{0}' 
                                    AND CONVERT(NVARCHAR,[MODIFY_DATE],112)+SUBSTRING((REPLACE((CONVERT(varchar(100), [MODIFY_DATE], 108)),':','')),1,4)<='{1}'


                              
                                    ", STARTTIME, ENDTIME);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }
        private void timer1_Tick(object sender, EventArgs e)
        {
            dateTimePicker2.Value = DateTime.Now;

            //DateTime DT1 = dateTimePicker1.Value;
            DateTime DT2 = dateTimePicker2.Value;
            DateTime DT3 = dateTimePicker3.Value;

            //TimeSpan TS1 = new TimeSpan( DT2.Ticks- DT1.Ticks );

            if (START.Equals("Y"))
            {
                //if (TS1.TotalHours >0.9)
                if (DT2.ToString("yyyyMMddHHMM").Equals(DT3.ToString("yyyyMMddHHMM")))
                {
                    ADDFILE();

                    ADDFILERCLOCK();

                    dateTimePicker1.Value = dateTimePicker2.Value;
                    //MessageBox.Show("GO");
                    dateTimePicker3.Value = DateTime.Now.AddHours(1);
                        

                }

                string CHECKTIMES = DateTime.Now.ToString("HH:mm");
                if (CHECKTIMES.Equals(textBox2.Text))
                {
                    ADDFILE3();
                }
            }



        }

        public void ADDFILE()
        {
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "刷卡紀錄.txt";

            DateTime SDT = dateTimePicker3.Value;
            SDT = SDT.AddHours(-1);
            DateTime EDT = dateTimePicker2.Value;

            DataTable DT = SERACHTB_EIP_PUNCH(SDT.ToString("yyyyMMddHH") + "00", EDT.ToString("yyyyMMddHH") + "00");

           
            if (DT!=null && DT.Rows.Count > 0)
            {
                try
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path + @"\" + Filename, false))
                    {
                        foreach (DataRow dr in DT.Rows)
                        {
                            file.WriteLine(dr["DATAS"].ToString());
                        }
                    }
                    ADDTB_EIP_PUNCH_RECORD(EDT.ToString("yyyy/MM/dd HH:mm:dd"), Filename);
                    //MessageBox.Show("OK");
                }
                catch
                {

                }

                finally
                {

                }
            }
        }

        public void ADDFILERCLOCK()
        {
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "刷卡紀錄RCLOCK.txt";

            DateTime SDT = dateTimePicker3.Value;
            SDT = SDT.AddHours(-1);
            DateTime EDT = dateTimePicker2.Value;

            DataTable DT = SERACHRCLOCK(SDT.ToString("yyyyMMddHH") + "00", EDT.ToString("yyyyMMddHH") + "00");


            if (DT != null && DT.Rows.Count > 0)
            {
                try
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path + @"\" + Filename, false))
                    {
                        foreach (DataRow dr in DT.Rows)
                        {
                            file.WriteLine(dr["DATAS"].ToString());
                        }
                    }

                    ADDTB_EIP_PUNCH_RECORD(EDT.ToString("yyyy/MM/dd HH:mm:dd"), Filename);
                    //MessageBox.Show("OK");
                }
                catch
                {

                }

                finally
                {

                }
            }
        }

        public DataTable SERACHRCLOCK(string STARTTIME, string ENDTIME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT [PF_USER]+' '+CONVERT(NVARCHAR,[PF_CLOCK],112)+' '+SUBSTRING((REPLACE((CONVERT(varchar(100), [PF_CLOCK], 108)),':','')),1,4) AS DATAS,[PF_CLOCK]
                                    FROM [PF_clock].[dbo].[RCLOCK]
                                    WHERE 
                                    CONVERT(NVARCHAR,[PF_CLOCK],112)+SUBSTRING((REPLACE((CONVERT(varchar(100), [PF_CLOCK], 108)),':','')),1,4)>='{0}' 
                                    AND CONVERT(NVARCHAR,[PF_CLOCK],112)+SUBSTRING((REPLACE((CONVERT(varchar(100), [PF_CLOCK], 108)),':','')),1,4)<='{1}'

                              
                                    ", STARTTIME, ENDTIME);


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void ADDTB_EIP_PUNCH_RECORD(string EXETIME,string TXTNAME)
        {

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();


      
                sbSql.AppendFormat(@" 
                                    INSERT [TKHR].[dbo].[TB_EIP_PUNCH_RECORD]
                                    ([EXETIME],[TXTNAME])
                                    VALUES
                                    ('{0}','{1}')
                                    ", EXETIME, TXTNAME);


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

        public void ADDFILE2()
        {
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "補卡紀錄.txt";

            DateTime SDT = dateTimePicker4.Value;
            //SDT = SDT.AddHours(-1);
            DateTime EDT = dateTimePicker5.Value;

            DataTable DT = SERACHTB_EIP_PUNCH(SDT.ToString("yyyyMMddHH") + "00", EDT.ToString("yyyyMMddHH") + "00");


            if (DT != null && DT.Rows.Count > 0)
            {
                try
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path + @"\" + Filename, false))
                    {
                        foreach (DataRow dr in DT.Rows)
                        {
                            file.WriteLine(dr["DATAS"].ToString());
                        }
                    }

                    ADDTB_EIP_PUNCH_RECORD(SDT.ToString("yyyy/MM/dd HH:mm:dd"), Filename);
                    MessageBox.Show("OK");
                }
                catch
                {

                }

                finally
                {

                }
            }
        }

        public void ADDFILERCLOCK2()
        {
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "補卡紀錄RCLOCK.txt";

            DateTime SDT = dateTimePicker4.Value;
            //SDT = SDT.AddHours(-1);
            DateTime EDT = dateTimePicker5.Value;

            DataTable DT = SERACHRCLOCK(SDT.ToString("yyyyMMddHH") + "00", EDT.ToString("yyyyMMddHH") + "00");


            if (DT != null && DT.Rows.Count > 0)
            {
                try
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path + @"\" + Filename, false))
                    {
                        foreach (DataRow dr in DT.Rows)
                        {
                            file.WriteLine(dr["DATAS"].ToString());
                        }
                    }

                    ADDTB_EIP_PUNCH_RECORD(SDT.ToString("yyyy/MM/dd HH:mm:dd"), Filename);
                    MessageBox.Show("OK");
                }
                catch
                {

                }

                finally
                {

                }
            }
        }

        public void ADDFILE3()
        {
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "理級補卡紀錄.txt";


            DateTime SDT = dateTimePicker4.Value;
            //SDT = SDT.AddHours(-1);
            DateTime EDT = dateTimePicker5.Value;

            DataTable DT = SERACHTB_CARDEMP();

           

            if (DT != null && DT.Rows.Count > 0)
            {
                try
                {
                    using (System.IO.StreamWriter file = new System.IO.StreamWriter(Path + @"\" + Filename, false))
                    {
                        foreach (DataRow dr in DT.Rows)
                        {
                            //set BeginTime,EndTime
                            Random Begin = new Random(Guid.NewGuid().GetHashCode());//亂數種子
                            int BeginTime = Begin.Next(15, 29);
                            Random End = new Random(Guid.NewGuid().GetHashCode());//亂數種子
                            int EndTime = End.Next(25, 59);

                            string SBeginTime = "08:" + BeginTime.ToString();
                            string SEndTime = "18:" + EndTime.ToString();

                            file.WriteLine(dr["ID"].ToString() + " " + DateTime.Now.ToString("yyyyMMdd") + " " + SBeginTime);
                            file.WriteLine(dr["ID"].ToString() + " " + DateTime.Now.ToString("yyyyMMdd") + " " + SEndTime);
                        }
                    }

                    //ADDTB_EIP_PUNCH_RECORD(SDT.ToString("yyyy/MM/dd HH:mm:dd"), Filename);

                    MessageBox.Show("OK");
                }
                catch
                {

                }

                finally
                {

                }
            }
        }

        public DataTable SERACHTB_CARDEMP()
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT
                                    [EmployeeId]
                                    ,[NAME]
                                    ,[ID]
                                    FROM [TKHR].[dbo].[CARDEMP]
                              
                                    ");


                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "ds1");
                sqlConn.Close();

                if (ds1.Tables["ds1"].Rows.Count >= 1)
                {
                    return ds1.Tables["ds1"];

                }
                else
                {
                    return null;
                }

            }
            catch
            {
                return null;
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
            if (START.Equals("N"))
            {
                START = "Y";

                button1.Text = "啟動";
                button1.BackColor = Color.Blue;

                dateTimePicker1.Value = dateTimePicker2.Value;

                DateTime dt = dateTimePicker2.Value;
                int MINS = dt.Minute*-1;
                dt =dt.AddHours(1);
                dt = dt.AddMinutes(MINS);

                //dt = dt.AddHours(1);

                dateTimePicker3.Value = dt;

            }
            else
            {
                START = "N";

                button1.Text = "未啟動";
                button1.BackColor = Color.Red;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.SelectedPath = this.textBox1.Text;
            path.ShowDialog();

            this.textBox1.Text = path.SelectedPath;
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ADDFILE();

            ADDFILERCLOCK();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            ADDFILE2();

            ADDFILERCLOCK2();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //string CHECKTIMES = DateTime.Now.ToString("HH:mm");
            //if (CHECKTIMES.Equals(textBox2.Text))
            //{
            //    ADDFILE3();
            //}

            ADDFILE3();
        }

        #endregion


    }
}
