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
using FastReport;
using FastReport.Data;
using TKITDLL;

namespace TKHR
{
    public partial class frmQUERYCARD : Form
    {
        SqlConnection sqlConn = new SqlConnection();
        SqlCommand sqlComm = new SqlCommand();
        string connectionString;
        StringBuilder sbSql = new StringBuilder();
        StringBuilder sbSqlQuery = new StringBuilder();
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        private Report report1;

        public frmQUERYCARD()
        {
            InitializeComponent();

            comboBox1load();
        }

        #region FUNCTION
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
            Sequel.AppendFormat(@"SELECT AttendanceRankId,Name FROM [HRMDB].[dbo].[AttendanceRank] WHERE [Name] LIKE '%休假%' OR [Name] LIKE '%生產%'  ORDER BY Name ");
            SqlDataAdapter da = new SqlDataAdapter(Sequel.ToString(), sqlConn);
            DataTable dt = new DataTable();
            sqlConn.Open();

            dt.Columns.Add("AttendanceRankId", typeof(string));
            dt.Columns.Add("Name", typeof(string));
            da.Fill(dt);
            comboBox1.DataSource = dt.DefaultView;
            comboBox1.ValueMember = "AttendanceRankId";
            comboBox1.DisplayMember = "Name";
            sqlConn.Close();


        }

        public void SETFASTREPORT()
        {
            report1 = new Report();

            report1.Load(@"REPORT\查班別打卡.frx");

            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dberp"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;

            //report1.SetParameterValue("P1", comboBox1.Text.ToString());
            report1.SetParameterValue("P2", dateTimePicker1.Value.ToString("yyyyMMdd"));
            report1.SetParameterValue("P3", dateTimePicker2.Value.ToString("yyyyMMdd"));

            report1.Preview = previewControl1;
            report1.Show();
        }

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SETFASTREPORT();
        }

        #endregion
    }
}
