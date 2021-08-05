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

        string START = "N";



        public frmTBEIPPUNCH()
        {
            InitializeComponent();

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
                                    FROM [UOF].[dbo].[TB_EIP_PUNCH],[UOF].[dbo].[TB_EB_USER]
                                    WHERE [TB_EIP_PUNCH].USER_GUID=[TB_EB_USER].USER_GUID
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

            DateTime DT1 = dateTimePicker1.Value;
            DateTime DT2 = dateTimePicker2.Value;

            TimeSpan TS1 = new TimeSpan( DT2.Ticks- DT1.Ticks );

            if (START.Equals("Y"))
            {
                if (TS1.TotalHours >0.9)
                {
                    dateTimePicker1.Value = dateTimePicker2.Value;

                    ADDFILE();
                    //MessageBox.Show("GO");


                }
            }
            

            
        }

        public void ADDFILE()
        {
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "刷卡紀錄.txt";

            DateTime SDT = dateTimePicker1.Value;
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
            string Path = textBox1.Text;
            string Filename = DateTime.Now.ToString("yyyyMMddHH") + "刷卡紀錄.txt";

            DateTime SDT = dateTimePicker1.Value;
            SDT = SDT.AddHours(-1);
            DateTime EDT = dateTimePicker2.Value;

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



        #endregion

        
    }
}
