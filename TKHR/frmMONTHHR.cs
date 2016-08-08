using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI;
using NPOI.HPSF;
using NPOI.HSSF;
using NPOI.HSSF.UserModel;
using NPOI.POIFS;
using NPOI.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.Extractor;
using System.IO;
using System.Data.SqlClient;
using NPOI.SS.UserModel;
using System.Configuration;
using NPOI.XSSF.UserModel;

namespace TKHR
{
    public partial class frmMONTHHR : Form
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
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;
        string NowDB = "test";


        public frmMONTHHR()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void Search()
        {
            try
            {

                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    if (checkBox1.Checked == true)
                    {
                        sbSql.AppendFormat(@"SELECT [HRYEARS] AS '年',[HRMONTHS] AS '月',[HRIN] AS '到職人數',[HRRESIGN] AS '離職人數',[HROFRESIGN] AS '正職離職人數',[HRMANUIN] AS '生產部到職人數',[HRMANURESIGN] AS '生產部離職人數',[HRPTRESIGN] AS 'PT離職人數',[HRNOW] AS '現有人數',[HRLOST] AS '缺編人數',[HRINRATE] AS '到職率',[HRRESIGNRATE] AS '離職率',[HROUTRATE] AS '耗損率',[HRSTAYRATE] AS '留職率',[HRMOVERATE] AS '流動率' FROM [TKHR].[dbo].[MONTHHR] WHERE  [HRYEARS]='{0}' AND [HRMONTHS]='{1}'  ", dateTimePicker1.Value.Year.ToString(), dateTimePicker1.Value.Month.ToString());
                    }
                    else
                    {
                        
                        sbSql.AppendFormat(@"SELECT [HRYEARS] AS '年',[HRMONTHS] AS '月',[HRIN] AS '到職人數',[HRRESIGN] AS '離職人數',[HROFRESIGN] AS '正職離職人數',[HRMANUIN] AS '生產部到職人數',[HRMANURESIGN] AS '生產部離職人數',[HRPTRESIGN] AS 'PT離職人數',[HRNOW] AS '現有人數',[HRLOST] AS '缺編人數',[HRINRATE] AS '到職率',[HRRESIGNRATE] AS '離職率',[HROUTRATE] AS '耗損率',[HRSTAYRATE] AS '留職率',[HRMOVERATE] AS '流動率' FROM [TKHR].[dbo].[MONTHHR] WHERE  [HRYEARS]='{0}'   ", dateTimePicker1.Value.Year.ToString());
                    }
                    

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "TEMPds");
                    sqlConn.Close();


                    if (ds.Tables["TEMPds"].Rows.Count == 0)
                    {

                    }
                    else
                    {

                        dataGridView1.DataSource = ds.Tables["TEMPds"];
                        dataGridView1.AutoResizeColumns();
                    }
                }
                else
                {

                }

            }
            catch
            {

            }
            finally
            {

            }
        }

        public void HRADD()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD COPTC
                sbSql.Append(" ");
                sbSql.AppendFormat(" INSERT INTO [TKHR].[dbo].[MONTHHR] ([HRYEARS],[HRMONTHS],[HRIN],[HRRESIGN],[HROFRESIGN],[HRMANUIN],[HRMANURESIGN],[HRPTRESIGN],[HRNOW],[HRLOST]) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}') ", dateTimePicker2.Value.Year.ToString(), dateTimePicker2.Value.Month.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString());

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

                sqlConn.Close();
                Search();
            }
            catch
            {

            }
            finally
            {

            }
        }

        public void HRUPDATE()
        {
            try
            {
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                //ADD COPTC
                sbSql.Append(" ");
                sbSql.AppendFormat(" UPDATE  [TKHR].[dbo].[MONTHHR] SET [HRIN]='{2}',[HRRESIGN]='{3}',[HROFRESIGN]='{4}',[HRMANUIN]='{5}',[HRMANURESIGN]='{6}',[HRPTRESIGN]='{7}',[HRNOW]='{8}',[HRLOST]='{9}' WHERE  [HRYEARS]='{0}' AND [HRMONTHS]='{1}'", dateTimePicker2.Value.Year.ToString(), dateTimePicker2.Value.Month.ToString(), textBox1.Text.ToString(), textBox2.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString());

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

                sqlConn.Close();
                Search();
            }
            catch
            {

            }
            finally
            {

            }
        }

        public void HRDEL()
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("是否真的要刪除", "del?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    //ADD COPTC
                    sbSql.Append(" ");
                    sbSql.AppendFormat(" DELETE  [TKHR].[dbo].[MONTHHR]   WHERE  [HRYEARS]='{0}' AND [HRMONTHS]='{1}'  ", dateTimePicker2.Value.Year.ToString(), dateTimePicker2.Value.Month.ToString());

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
                    sqlConn.Close();
                    Search();
                }

                else if (dialogResult == DialogResult.No)
                {
                    //do something else
                }
            }
            catch
            {

            }
            finally
            {

            }

        }

        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            int columnIndex = 0;
            int rowIndex = 0;
            try
            {
                columnIndex = dataGridView1.CurrentCell.ColumnIndex;
                rowIndex = dataGridView1.CurrentCell.RowIndex;

                DateTime dt = Convert.ToDateTime(dataGridView1.Rows[rowIndex].Cells[0].Value.ToString() + "/" + dataGridView1.Rows[rowIndex].Cells[1].Value.ToString() + "/1");

                dateTimePicker2.Value = dt;
                textBox1.Text = dataGridView1.Rows[rowIndex].Cells[2].Value.ToString();
                textBox2.Text = dataGridView1.Rows[rowIndex].Cells[3].Value.ToString();
                textBox3.Text = dataGridView1.Rows[rowIndex].Cells[4].Value.ToString();
                textBox4.Text = dataGridView1.Rows[rowIndex].Cells[5].Value.ToString();
                textBox5.Text = dataGridView1.Rows[rowIndex].Cells[6].Value.ToString();
                textBox6.Text = dataGridView1.Rows[rowIndex].Cells[7].Value.ToString();
                textBox7.Text = dataGridView1.Rows[rowIndex].Cells[8].Value.ToString();
                textBox8.Text = dataGridView1.Rows[rowIndex].Cells[9].Value.ToString();
            }
            catch 
            {
            }

            
        }

        public void ExcelExport()
        {

            string NowDB = "TKHR";
            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;


            dt = ds.Tables["TEMPds"];
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
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString()));
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString()));
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));


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
            filename.AppendFormat(@"c:\temp\人力資源分析{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
            Search();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            HRADD();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            HRUPDATE();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            HRDEL();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        #endregion


    }
}
