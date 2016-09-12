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
    public partial class frmSAL : Form
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
        DataSet ds2 = new DataSet();
        DataSet ds3 = new DataSet();
        DataSet ds4 = new DataSet();
        DataSet ds5 = new DataSet();
        DataSet ds6 = new DataSet();
        DataSet dsYear = new DataSet();
        DataTable dt = new DataTable();
        string strFilePath;
        OpenFileDialog file = new OpenFileDialog();
        int result;
        string NowDay;
        string NowDB = "test";
        int rownum = 0;
        string NowTable = null;

        public frmSAL()
        {
            InitializeComponent();
            combobox1load();
            combobox2load();
            combobox3load();
        }

        #region FUNCTION
        public void combobox1load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT [ID],[JOBNAME]  FROM [TKHR].[dbo].[SALJOB]  ORDER BY [ID]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("JOBNAME", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "ID";
            comboBox1.DisplayMember = "JOBNAME";
            sqlConn.Close();

            comboBox1.SelectedValue = "01";

        }

        public void combobox2load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[EMPNAME] FROM [TKHR].[dbo].[SALEMPALOWANCE] ORDER BY [ID]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("EMPNAME", typeof(string));
            da.Fill(dt);
            comboBox2.DataSource = dt.DefaultView;
            comboBox2.ValueMember = "ID";
            comboBox2.DisplayMember = "EMPNAME";
            sqlConn.Close();

            comboBox2.SelectedValue = "00";

        }
        public void combobox3load()
        {

            connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            String Sequel = "SELECT  [ID],[JOBNAME],[JOBALLOWANCE]FROM [TKHR].[dbo].[SALJOBALOWANCE] ORDER BY [ID]";
            SqlDataAdapter da = new SqlDataAdapter(Sequel, sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("ID", typeof(string));
            dt.Columns.Add("JOBNAME", typeof(string));
            da.Fill(dt);
            comboBox3.DataSource = dt.DefaultView;
            comboBox3.ValueMember = "ID";
            comboBox3.DisplayMember = "JOBNAME";
            sqlConn.Close();

            comboBox3.SelectedValue = "00";

        }
        public void Search()
        {
            try
            {                
                connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);

                sbSql.Clear();
                sbSqlQuery.Clear();
 
                sbSql.Append(@" SELECT [SALEMPINFO].[ID] AS '工號',[NAME] AS '姓名',[SALJOB].[JOBNAME] AS '職務',[SALEMPALOWANCE].[EMPNAME] AS '職能別',[SALJOBALOWANCE].[JOBNAME] AS '幕僚別' ,[YEARS] AS '年資',[SALEMPINFO].[JOBLEVEL] AS '職等',[SALEMPINFO].[JOBYEAR] AS '職級',[TOTALMONEY] AS '總薪資',[SALJOB] AS '主管',[SALJOBLEVEL] AS '薪資點',[SALJOBALOWANCE] AS '幕僚',[SALEMPALOWANCE] AS '職能',[SALOTHER] AS '久任',[JOBADD] AS '主管津貼',[JOBALOWANCEADD] AS '幕僚加給',[EMPALOWANCEADD] AS '職能加給' ,[SALEMPINFO].[JOBID],[SALEMPINFO].[EMPID],[SALEMPINFO].[ALOWANCEID] ");
                sbSql.Append(@" FROM [TKHR].[dbo].[SALEMPINFO],[TKHR].[dbo].[SALJOB],[TKHR].[dbo].[SALEMPALOWANCE],[TKHR].[dbo].[SALJOBALOWANCE]");
                sbSql.Append(@" WHERE [SALEMPINFO].[JOBID]=[SALJOB].[ID] AND [SALEMPINFO].[EMPID]=[SALEMPALOWANCE].[ID] AND [SALEMPINFO].[ALOWANCEID]=[SALJOBALOWANCE].[ID]");

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
                    dt= ds.Tables["TEMPds"];
                    dataGridView1.DataSource = dt;
                    dataGridView1.AutoResizeColumns();
                    dataGridView1.CurrentCell = dataGridView1[1, rownum];
                    var curRow = dataGridView1.CurrentRow;
                    if (curRow != null)
                    {
                        textBox1.Text = dataGridView1.CurrentRow.Cells["工號"].Value.ToString();
                        textBox2.Text = dataGridView1.CurrentRow.Cells["姓名"].Value.ToString();
                        textBox3.Text = dataGridView1.CurrentRow.Cells["職等"].Value.ToString();
                        textBox4.Text = dataGridView1.CurrentRow.Cells["職級"].Value.ToString();
                        textBox5.Text = dataGridView1.CurrentRow.Cells["總薪資"].Value.ToString();
                        textBox6.Text = dataGridView1.CurrentRow.Cells["主管"].Value.ToString();
                        textBox7.Text = dataGridView1.CurrentRow.Cells["薪資點"].Value.ToString();
                        textBox8.Text = dataGridView1.CurrentRow.Cells["幕僚"].Value.ToString();
                        textBox9.Text = dataGridView1.CurrentRow.Cells["職能"].Value.ToString();
                        textBox10.Text = dataGridView1.CurrentRow.Cells["久任"].Value.ToString();
                        textBox11.Text = dataGridView1.CurrentRow.Cells["主管津貼"].Value.ToString();
                        textBox12.Text = dataGridView1.CurrentRow.Cells["幕僚加給"].Value.ToString();
                        textBox13.Text = dataGridView1.CurrentRow.Cells["職能加給"].Value.ToString();
                        comboBox1.SelectedValue = dataGridView1.CurrentRow.Cells["JOBID"].Value.ToString();
                        comboBox2.SelectedValue = dataGridView1.CurrentRow.Cells["EMPID"].Value.ToString();
                        comboBox3.SelectedValue = dataGridView1.CurrentRow.Cells["ALOWANCEID"].Value.ToString();
                        numericUpDown1.Value = Convert.ToInt32(dataGridView1.CurrentRow.Cells["年資"].Value.ToString());

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


        public void ExcelExport()
        {
            Search();

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
                if(i <=16)
                {
                    ws.GetRow(0).CreateCell(i).SetCellValue(dt.Columns[i].ColumnName);
                }
                
            }


            int j = 0;
            foreach (DataGridViewRow dr in this.dataGridView1.Rows)
            {
                ws.CreateRow(j + 1);
                ws.GetRow(j + 1).CreateCell(0).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[0].ToString());
                ws.GetRow(j + 1).CreateCell(1).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[1].ToString());
                ws.GetRow(j + 1).CreateCell(2).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[2].ToString());
                ws.GetRow(j + 1).CreateCell(3).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[3].ToString());
                ws.GetRow(j + 1).CreateCell(4).SetCellValue(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[4].ToString());
                ws.GetRow(j + 1).CreateCell(5).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[5].ToString()));
                ws.GetRow(j + 1).CreateCell(6).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[6].ToString()));
                ws.GetRow(j + 1).CreateCell(7).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[7].ToString()));
                ws.GetRow(j + 1).CreateCell(8).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[8].ToString()));
                ws.GetRow(j + 1).CreateCell(9).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[9].ToString()));
                ws.GetRow(j + 1).CreateCell(10).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[10].ToString()));
                ws.GetRow(j + 1).CreateCell(11).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[11].ToString()));
                ws.GetRow(j + 1).CreateCell(12).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[12].ToString()));
                ws.GetRow(j + 1).CreateCell(13).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[13].ToString()));
                ws.GetRow(j + 1).CreateCell(14).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[14].ToString()));
                ws.GetRow(j + 1).CreateCell(15).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[15].ToString()));
                ws.GetRow(j + 1).CreateCell(16).SetCellValue(Convert.ToInt32(((System.Data.DataRowView)(dr.DataBoundItem)).Row.ItemArray[16].ToString()));
                
                

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
            filename.AppendFormat(@"c:\temp\薪資明細表{0}.xlsx", DateTime.Now.ToString("yyyyMMdd"));

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
        private void comboBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            if (!string.IsNullOrEmpty(comboBox1.SelectedValue.ToString()))
            {               

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT [ID],[JOBNAME],[JOBMONEY],[JOBLEVEL],[JOBYEAR] FROM [TKHR].[dbo].[SALJOB] WHERE [ID]='{0}'",comboBox1.SelectedValue.ToString());
              
                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds2.Clear();
                adapter.Fill(ds2, "TEMPds2");
                sqlConn.Close();


                if (ds2.Tables["TEMPds2"].Rows.Count >= 1)
                {
                    textBox3.Text = ds2.Tables["TEMPds2"].Rows[0]["JOBLEVEL"].ToString();
                    textBox4.Text = ds2.Tables["TEMPds2"].Rows[0]["JOBYEAR"].ToString();
                    textBox6.Text = ds2.Tables["TEMPds2"].Rows[0]["JOBMONEY"].ToString();
                  
                }
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            CALLEVEMONEY();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            CALLEVEMONEY();
        }

        public void CALLEVEMONEY()
        {
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);
            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@" SELECT [JOBLEVEL],[JOBYEAR],[JOBLVMONEY] FROM [TKHR].[dbo].[SALJOBLEVEL] WHERE [JOBLEVEL]='{0}' AND [JOBYEAR]='{1}'", textBox3.Text.ToString(),textBox4.Text.ToString());

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            ds3.Clear();
            adapter.Fill(ds3, "TEMPds3");
            sqlConn.Close();


            if (ds3.Tables["TEMPds3"].Rows.Count >= 1)
            {
                textBox7.Text = ds3.Tables["TEMPds3"].Rows[0]["JOBLVMONEY"].ToString();    
            }

            if(textBox3.Text.ToString().Equals("0")&& textBox3.Text.ToString().Equals("0"))
            {
                textBox7.Text = "0";
            }
        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            if (!string.IsNullOrEmpty(comboBox2.SelectedValue.ToString()))
            {

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT  [ID],[EMPNAME],[EMPALOWANCE] FROM [TKHR].[dbo].[SALEMPALOWANCE] WHERE  [ID]='{0}'", comboBox2.SelectedValue.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds4.Clear();
                adapter.Fill(ds4, "TEMPds4");
                sqlConn.Close();


                if (ds4.Tables["TEMPds4"].Rows.Count >= 1)
                {
                    textBox9.Text = ds4.Tables["TEMPds4"].Rows[0]["EMPALOWANCE"].ToString();
                   

                }
            }
        }

        private void comboBox3_SelectedValueChanged(object sender, EventArgs e)
        {
            connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            sqlConn = new SqlConnection(connectionString);

            if (!string.IsNullOrEmpty(comboBox3.SelectedValue.ToString()))
            {

                sbSql.Clear();
                sbSqlQuery.Clear();

                sbSql.AppendFormat(@" SELECT  [ID],[JOBNAME],[JOBALLOWANCE] FROM [TKHR].[dbo].[SALJOBALOWANCE] WHERE [ID]='{0}'", comboBox3.SelectedValue.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                sqlCmdBuilder = new SqlCommandBuilder(adapter);

                sqlConn.Open();
                ds5.Clear();
                adapter.Fill(ds5, "TEMPds5");
                sqlConn.Close();


                if (ds5.Tables["TEMPds5"].Rows.Count >= 1)
                {
                    textBox8.Text = ds5.Tables["TEMPds5"].Rows[0]["JOBALLOWANCE"].ToString();


                }
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            if(numericUpDown1.Value>=1)
            {
                textBox11.Text = (300 * numericUpDown1.Value).ToString();
                textBox12.Text = (300 * numericUpDown1.Value).ToString();
                textBox13.Text = (200 * numericUpDown1.Value).ToString();
            }
            if((numericUpDown1.Value%2)==0)
            {
                textBox4.Text = (Convert.ToInt32(ds2.Tables["TEMPds2"].Rows[0]["JOBYEAR"].ToString()) + Convert.ToInt32(numericUpDown1.Value / 2)).ToString();
            }
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        public void CALOTHER()
        {
            int int5, int6, int7, int8, int9, int10, int11, int12, int13;
            int5 = Convert.ToInt32(textBox5.Text.ToString());
            int6 = Convert.ToInt32(textBox6.Text.ToString());
            int7 = Convert.ToInt32(textBox7.Text.ToString());
            int8 = Convert.ToInt32(textBox8.Text.ToString());
            int9 = Convert.ToInt32(textBox9.Text.ToString());
            int10 = Convert.ToInt32(textBox10.Text.ToString());
            int11 = Convert.ToInt32(textBox11.Text.ToString());
            int12 = Convert.ToInt32(textBox12.Text.ToString());
            int13 = Convert.ToInt32(textBox13.Text.ToString());

            int10 = int5 - int6 - int7 - int8 - int9 - int11 - int12 - int13;
            textBox10.Text = int10.ToString();


        }
        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        private void textBox12_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            CALOTHER();
        }


        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            var curRow = dataGridView1.CurrentRow;
            if (curRow != null)
            {
                textBox1.Text = dataGridView1.CurrentRow.Cells["工號"].Value.ToString();
                textBox2.Text = dataGridView1.CurrentRow.Cells["姓名"].Value.ToString();
                textBox3.Text = dataGridView1.CurrentRow.Cells["職等"].Value.ToString();
                textBox4.Text = dataGridView1.CurrentRow.Cells["職級"].Value.ToString();
                textBox5.Text = dataGridView1.CurrentRow.Cells["總薪資"].Value.ToString();
                textBox6.Text = dataGridView1.CurrentRow.Cells["主管"].Value.ToString();
                textBox7.Text = dataGridView1.CurrentRow.Cells["薪資點"].Value.ToString();
                textBox8.Text = dataGridView1.CurrentRow.Cells["幕僚"].Value.ToString();
                textBox9.Text = dataGridView1.CurrentRow.Cells["職能"].Value.ToString();
                textBox10.Text = dataGridView1.CurrentRow.Cells["久任"].Value.ToString();
                textBox11.Text = dataGridView1.CurrentRow.Cells["主管津貼"].Value.ToString();
                textBox12.Text = dataGridView1.CurrentRow.Cells["幕僚加給"].Value.ToString();
                textBox13.Text = dataGridView1.CurrentRow.Cells["職能加給"].Value.ToString();
                comboBox1.SelectedValue= dataGridView1.CurrentRow.Cells["JOBID"].Value.ToString();
                comboBox2.SelectedValue = dataGridView1.CurrentRow.Cells["EMPID"].Value.ToString();
                comboBox3.SelectedValue = dataGridView1.CurrentRow.Cells["ALOWANCEID"].Value.ToString();
                numericUpDown1.Value= Convert.ToInt32(dataGridView1.CurrentRow.Cells["年資"].Value.ToString());

            }

        }

        public void SAVESAL()
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
                sbSql.Append(" INSERT INTO  [TKHR].[dbo].[SALEMPINFO] ([ID],[NAME],[JOBID],[EMPID],[ALOWANCEID],[YEARS],[JOBLEVEL],[JOBYEAR],[TOTALMONEY],[SALJOB],[SALJOBLEVEL],[SALJOBALOWANCE],[SALEMPALOWANCE],[SALOTHER],[JOBADD],[JOBALOWANCEADD],[EMPALOWANCEADD])");
                sbSql.AppendFormat(" VALUES('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}','{15}','{16}')", textBox1.Text.ToString(), textBox2.Text.ToString(), comboBox1.SelectedValue.ToString(), comboBox2.SelectedValue.ToString(), comboBox3.SelectedValue.ToString(), numericUpDown1.Value.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString());
                sbSql.Append(" ");

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
                    Search();
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

        public void UPDATESAL()
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
                sbSql.AppendFormat(" UPDATE  [TKHR].[dbo].[SALEMPINFO] SET [NAME]='{1}',[JOBID]='{2}',[EMPID]='{3}',[ALOWANCEID]='{4}',[YEARS]='{5}',[JOBLEVEL]='{6}',[JOBYEAR]='{7}',[TOTALMONEY]='{8}',[SALJOB]='{9}',[SALJOBLEVEL]='{10}',[SALJOBALOWANCE]='{11}',[SALEMPALOWANCE]='{12}',[SALOTHER]='{13}',[JOBADD]='{14}',[JOBALOWANCEADD]='{15}',[EMPALOWANCEADD]='{16}' WHERE [ID]='{0}'", textBox1.Text.ToString(),textBox2.Text.ToString(),comboBox1.SelectedValue.ToString(),comboBox2.SelectedValue.ToString(),comboBox3.SelectedValue.ToString(),numericUpDown1.Value.ToString(), textBox3.Text.ToString(), textBox4.Text.ToString(), textBox5.Text.ToString(), textBox6.Text.ToString(), textBox7.Text.ToString(), textBox8.Text.ToString(), textBox9.Text.ToString(), textBox10.Text.ToString(), textBox11.Text.ToString(), textBox12.Text.ToString(), textBox13.Text.ToString());

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
                    rownum = dataGridView1.CurrentCell.RowIndex; 
                    Search();
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

        public void DELSQL()
        {
            try
            {
                DialogResult dialogResult = MessageBox.Show("是否真的要刪除", "del?", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    //add ZWAREWHOUSEPURTH
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sqlConn.Close();
                    sqlConn.Open();
                    tran = sqlConn.BeginTransaction();

                    sbSql.Clear();
                    sbSql.AppendFormat(" DELETE   [TKHR].[dbo].[SALEMPINFO] WHERE ID='{0}'", textBox1.Text.ToString());

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
                        Search();
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

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            FINDCNAME();
        }
        public void FINDCNAME()
        {
            sbSql.Clear();
            sbSqlQuery.Clear();

            sbSql.AppendFormat(@" SELECT [CnName],[Code] FROM [HRMDB].[dbo].[Employee] WHERE [Code]='{0}'  ", textBox1.Text.ToString());

            adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
            sqlCmdBuilder = new SqlCommandBuilder(adapter);

            sqlConn.Open();
            ds6.Clear();
            adapter.Fill(ds6, "TEMPds6");
            sqlConn.Close();


            if (ds6.Tables["TEMPds6"].Rows.Count >= 1)
            {
                textBox2.Text = ds6.Tables["TEMPds6"].Rows[0]["CnName"].ToString();


            }
        }
        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            Search();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            ExcelExport();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "0";
            textBox4.Text = "0";
            textBox5.Text = "0";
            textBox6.Text = "0";
            textBox7.Text = "0";
            textBox8.Text = "0";
            textBox9.Text = "0";
            textBox10.Text = "0";
            textBox11.Text = "0";
            textBox12.Text = "0";
            textBox13.Text = "0";
            comboBox1.SelectedValue = "01";
            comboBox2.SelectedValue = "00";
            comboBox3.SelectedValue = "00";
            numericUpDown1.Value = 0;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            UPDATESAL();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DELSQL();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SAVESAL();
        }

        
        #endregion


    }
}
