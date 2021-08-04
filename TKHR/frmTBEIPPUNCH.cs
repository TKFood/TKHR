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
        string START = "N";



        public frmTBEIPPUNCH()
        {
            InitializeComponent();

            textBox1.Text = @"C:\SCSHR\Card";
        }

        #region FUNCTION

        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            if (START.Equals("N"))
            {
                START = "Y";

                button1.Text = "啟動";
                button1.BackColor = Color.Blue;
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
            path.ShowDialog();
            this.textBox1.Text = path.SelectedPath;
        }

        #endregion


    }
}
