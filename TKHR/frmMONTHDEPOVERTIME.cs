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
    public partial class frmMONTHDEPOVERTIME : Form
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
        DataSet dsYear = new DataSet();
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;
        string NowDB = "test";
        int rownum = 0;
        string NowTable = null;

        public frmMONTHDEPOVERTIME()
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

                    sbSql.AppendFormat(@"SELECT [HRYEARS] AS '年',[HRMONTHS] AS '月',[DEPNO] AS '部門代號',[DEPNAME] AS '部門',[HROTHR] AS '總加班時數',[HROTFEE] AS '總加班金額' FROM [TKHR].[dbo].[MONTHDEPOVERTIME] WHERE [HRYEARS]='{0}' AND [HRMONTHS]='{1}'  ORDER BY [HRYEARS],[HRMONTHS],[DEPNO] ", dateTimePicker1.Value.Year.ToString(), dateTimePicker1.Value.Month.ToString());
                    
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
                        dataGridView1.CurrentCell = dataGridView1[3, rownum];
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

        public void Loaddep()
        {
            try
            {

                if (!string.IsNullOrEmpty(dateTimePicker1.Value.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [HRYEARS],[HRMONTHS],[DEPNO],[DEPNAME],[HROTHR],[HROTFEE] FROM [TKHR].[dbo].[MONTHDEPOVERTIME] WHERE [HRYEARS]='{0}' AND [HRMONTHS]='{1}'  ORDER BY [HRYEARS],[HRMONTHS],[DEPNO] ", dateTimePicker1.Value.Year.ToString(), dateTimePicker1.Value.Month.ToString());
                    
                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "TEMPds");
                    sqlConn.Close();


                    if (ds.Tables["TEMPds"].Rows.Count == 0)
                    {
                        defaultdep();
                    }
                    else
                    {
                        DialogResult dialogResult = MessageBox.Show("是否真的要清空", "del?", MessageBoxButtons.YesNo);
                        if (dialogResult == DialogResult.Yes)
                        {
                            defaultdep();
                        }
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

        public void defaultdep()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            sqlConn.Close();
            sqlConn.Open();
            tran = sqlConn.BeginTransaction();

            sbSql.Clear();
            //ADD COPTC
            sbSql.Append(" ");
            sbSql.AppendFormat(" INSERT INTO [TKHR].[dbo].[MONTHDEPOVERTIME] ([HRYEARS],[HRMONTHS],[DEPNO],[DEPNAME],[HROTHR],[HROTFEE]) SELECT '{0}','{1}',[Code],[Name],0,0 FROM [TKHR].[dbo].[Department] ORDER BY Code ", dateTimePicker1.Value.Year.ToString(), dateTimePicker1.Value.Month.ToString());

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
                sbSql.AppendFormat(" UPDATE  [TKHR].[dbo].[MONTHDEPOVERTIME]  SET [HROTHR] ='{3}',[HROTFEE]='{4}' WHERE  [HRYEARS]='{0}' AND [HRMONTHS]='{1}' AND [DEPNO]='{2}'", dateTimePicker2.Value.Year.ToString(), dateTimePicker2.Value.Month.ToString(), textBox1.Text.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString());

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

                rownum = dataGridView1.CurrentCell.RowIndex; ;
                Search();


            }
            catch
            {

            }
            finally
            {

            }
        }


        public void HRUPDATENULL()
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
                sbSql.AppendFormat(" UPDATE  [TKHR].[dbo].[MONTHDEPOVERTIME]  SET [HROTHR] ='{3}',[HROTFEE]='{4}' WHERE  [HRYEARS]='{0}' AND [HRMONTHS]='{1}' AND [DEPNO]='{2}'", dateTimePicker2.Value.Year.ToString(), dateTimePicker2.Value.Month.ToString(), textBox1.Text.ToString(),"0", "0");

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
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString()));
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToDouble(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                

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
            filename.AppendFormat(@"c:\temp\每月各部門加班明細表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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

        public void ExportExcel()
        {
            String NowDB = "TKHR";
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@"SELECT [HRYEARS] AS '年',[HRMONTHS] AS '月',[DEPNO] AS '部門代號',[DEPNAME] AS '部門',[HROTHR] AS '總加班時數',[HROTFEE] AS '總加班金額' FROM [TKHR].[dbo].[MONTHDEPOVERTIME] WHERE [HRYEARS]='{0}'   ORDER BY [HRYEARS],[HRMONTHS],[DEPNO] ", dateTimePicker1.Value.Year.ToString());

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            dsYear.Clear();
            adapter.Fill(dsYear, "TEMPdsYear");
            sqlConn.Close();

            //建立Excel 2003檔案
            IWorkbook wb = new XSSFWorkbook();
            ISheet ws;

            dt = dsYear.Tables["TEMPdsYear"];

            ////建立Excel 2007檔案
            //IWorkbook wb = new XSSFWorkbook();
            //ISheet ws;

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

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                ws.CreateRow(i + 1);
                for (int j = 0; j < dt.Columns.Count; j++)
                {                   
                    if (j==4)
                    {
                        ws.GetRow(i + 1).CreateCell(j).SetCellValue(Convert.ToInt32(dt.Rows[i][j].ToString()));
                    }
                    else  if (j == 5)
                    {
                        ws.GetRow(i + 1).CreateCell(j).SetCellValue(Convert.ToInt32(dt.Rows[i][j].ToString()));
                    }
                    else
                    {
                        ws.GetRow(i + 1).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                    }
                }
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
            filename.AppendFormat(@"c:\temp\各部門每月加班費總表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        private void button6_Click(object sender, EventArgs e)
        {
            Loaddep();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            HRUPDATE();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            HRUPDATENULL();
        }


        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            ExportExcel();
        }
        #endregion


    }
}
