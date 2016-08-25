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
    public partial class frmSysRESETPS : Form
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

        public frmSysRESETPS()
        {
            InitializeComponent();
        }

        #region FUNCTION
        public void RESETPS()
        {
            try
            {

                if (!string.IsNullOrEmpty(txt_UserName.Text.ToString())&& !string.IsNullOrEmpty(txt_Password.Text.ToString())&& !string.IsNullOrEmpty(textBox1.Text.ToString()))
                {
                    connectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
                    sqlConn = new SqlConnection(connectionString);

                    sbSql.Clear();
                    sbSqlQuery.Clear();

                    sbSql.AppendFormat(@"SELECT [UserName],[Password]FROM [TKHR].[dbo].[MNU_Login] WHERE [UserName]='{0}' AND [Password]='{1}'", txt_UserName.Text.ToString(),txt_Password.Text.ToString());

                    adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);
                    sqlCmdBuilder = new SqlCommandBuilder(adapter);

                    sqlConn.Open();
                    ds.Clear();
                    adapter.Fill(ds, "TEMPds");
                    sqlConn.Close();


                    if (ds.Tables["TEMPds"].Rows.Count == 0)
                    {
                        DialogResult dialogResult = MessageBox.Show("找不到此人員", "CHECK?");

                        txt_UserName.Text = null;
                        txt_Password.Text = null;
                        textBox1.Text = null;
                        txt_UserName.Select();
                    }
                    else
                    {

                        DialogResult dialogResult = MessageBox.Show("是否真的要設定密碼", "CHECK?", MessageBoxButtons.YesNo);
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
                            sbSql.AppendFormat(" UPDATE [TKHR].[dbo].[MNU_Login] SET [Password]='{0}' WHERE [UserName]='{1}' ", textBox1.Text.ToString(),txt_UserName.Text.ToString());

                            cmd.Connection = sqlConn;
                            cmd.CommandTimeout = 60;
                            cmd.CommandText = sbSql.ToString();
                            cmd.Transaction = tran;
                            result = cmd.ExecuteNonQuery();

                            if (result == 0)
                            {
                                tran.Rollback();    //交易取消   
                                dialogResult = MessageBox.Show("更改失敗", "CHECK?");
                                txt_UserName.Select();
                            }
                            else
                            {
                                tran.Commit();      //執行交易          
                                dialogResult = MessageBox.Show("更改成功", "OK");
                                txt_UserName.Select();
                            }

                            txt_UserName.Text = null;
                            txt_Password.Text = null;
                            textBox1.Text = null;
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

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            RESETPS();
        }
        #endregion

    }
}
