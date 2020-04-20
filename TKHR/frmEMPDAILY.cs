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

namespace TKHR
{
    public partial class frmEMPDAILY : Form
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

        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataSet ds1 = new DataSet();
        DataTable tableSql = new DataTable();

        public static int rows = 1;   
        public static int colums = 1;


        public frmEMPDAILY()
        {
            InitializeComponent();
        }

        #region FUNCTION

        #endregion
        public void OPENFILE()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "Select file";
            dialog.InitialDirectory = ".\\";
            dialog.Filter = "Excel Files(.xlsx)|*.xlsx|xls files (*.*)|*.xls ";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                //MessageBox.Show(dialog.FileName);

                string strSQL;  //SQL字串
                String cnnS = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source="+ dialog.FileName + ";" + "Extended Properties=\"EXCEL 12.0;HDR=YES\"";  //資料庫連接字串
                OleDbConnection cnn = new OleDbConnection(cnnS);
                strSQL = " Select * From [表單回應 1$] "; //選擇所有資料列從工作表TABLE1 
                using (OleDbDataAdapter dr = new OleDbDataAdapter(strSQL, cnn))
                {

                    dr.Fill(tableSql);  //將所有資料填充至tableSql
                    this.dataGridView1.DataSource = tableSql;
                }

                cnn.Close();

                

            }
        }

        public void ADDEMPDAILY()
        {
            StringBuilder SQL = new StringBuilder();
            if(tableSql.Rows.Count>0)
            {
                connectionString = ConfigurationManager.ConnectionStrings["dberp"].ConnectionString;
                sqlConn = new SqlConnection(connectionString);
                sqlConn.Open();

                if (dataGridView1.Rows.Count > 0)
                {
                    DataRow dr = null;
                    for (int i = 0; i < tableSql.Rows.Count; i++)
                    {
                        dr = tableSql.Rows[i];
                    
                        DateTime dt = Convert.ToDateTime(dr["時間戳記"].ToString().Replace("下午", "PM").Replace("上午", "AM"));

                        SQL.AppendFormat(@" INSERT INTO [TKHR].[dbo].[EMPDAILY]");
                        SQL.AppendFormat(@" ( [ID],[NAME],[DATES])");
                        SQL.AppendFormat(@" VALUES ('{0}','{1}','{2}')", dr["ID"].ToString(), dr["NAME"].ToString(), dt.ToString("yyyy-MM-dd HH:mm:ss"));
                        SQL.AppendFormat(@" ");
                    }

                    cmd = new SqlCommand(SQL.ToString(), sqlConn);
                    cmd.ExecuteNonQuery();
                    sqlConn.Close();
                    MessageBox.Show("匯入成功");
                }
                else
                {
                    MessageBox.Show("匯入失敗");
                }
            }
        }

      
        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            OPENFILE();
        }
        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            ADDEMPDAILY();
        }
    }
}
