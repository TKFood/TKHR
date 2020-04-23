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

                        SQL.AppendFormat(@" INSERT INTO [TKWEB].[dbo].[QUESTIONNAIRES]");
                        SQL.AppendFormat(@" ([DATES],[NO],[NAME],[DEP],[QUESTION1],[QUESTION2],[QUESTION3],[QUESTION4],[QUESTION5],[QUESTION6],[QUESTION7],[QUESTION8],[QUESTION9],[QUESTION10],[QUESTION11])");
                        SQL.AppendFormat(@" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}','{14}')", dt.ToString("yyyy-MM-dd HH:mm:ss"), dr["工號"].ToString(), dr["姓名"].ToString(), dr["部門"].ToString(), dr["請問24小時內，您與您同住的家屬/室友否出現以下微狀(複選)"].ToString(), dr["承上題，如有症狀請簡短說明何時、何地、何人"].ToString(), dr["請問24小時內您與您的同住的家屬/室友是否從其他國家入境台灣？"].ToString(), dr["承上題，簡短說明何時、何地、何人、班次?"].ToString(), dr["請問24小時內您與您的同住的家屬/室友是否曾與已確診/疑似/正在接受檢驗之新型冠狀病毒肺炎病患有接觸？"].ToString(), dr["承上題，簡短說明何時接觸、何地接觸、何人接觸?"].ToString(), dr["請問24小時內您與您的同住的家屬/室友是否曾前往非閉密空間但人潮擁擠的公共場所(無適當社交距離1M)，如旅遊地區、夜市、風景地區"].ToString(), dr["承上題，簡短說明何時、何地、何人、共約幾人"].ToString(), dr["請問24小時內您與您的同住的家屬/室友是否曾搭乘大眾交通運輸工具（公車、台鐵、高鐵、捷運、渡輪、客運、遊覽車…）"].ToString(), dr["承上題，簡短說明何時、何地、何人、何種交通工具、班次?"].ToString(), dr["其他想告知的事項"].ToString());
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

        public void SETFASTREPORT()
        {
            StringBuilder SQL1 = new StringBuilder();

            SQL1 = SETSQL();
            Report report1 = new Report();
            report1.Load(@"REPORT\特別-未回覆問卷.frx");

            report1.Dictionary.Connections[0].ConnectionString = ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL1.ToString();

            report1.SetParameterValue("P1", dateTimePicker1.Value.ToString("yyyyMMdd"));
            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL()
        {
            StringBuilder SB = new StringBuilder();

            SB.AppendFormat(" SELECT ID AS '工號',NAME AS '姓名',ME001 AS '代號',ME002 AS '部門'");
            SB.AppendFormat(" FROM [TKHR].[dbo].[EMP],[TK].dbo.CMSMV,[TK].dbo.CMSME");
            SB.AppendFormat(" WHERE  MV004=ME001");
            SB.AppendFormat(" AND ID=MV001");
            SB.AppendFormat(" AND ID NOT IN (SELECT [ID] FROM [TKHR].[dbo].[EMPDAILY] WHERE CONVERT(nvarchar,[DATES],112)='{0}')",dateTimePicker1.Value.ToString("yyyyMMdd"));
            SB.AppendFormat(" ORDER BY ME001,ID,NAME ");
            SB.AppendFormat(" ");
            SB.AppendFormat(" ");

            return SB;

        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            OPENFILE();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ADDEMPDAILY();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
