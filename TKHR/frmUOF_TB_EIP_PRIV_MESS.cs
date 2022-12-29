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
using System.Configuration;
using FastReport;
using FastReport.Data;
using System.Net.Mail;//<-基本上發mail就用這個class
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Diagnostics;
using System.Threading;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HSSF.Util;
using NPOI.XSSF.UserModel;
using TKITDLL;
using System.Net.Http;
using System.Net;

namespace TKHR
{
    public partial class frmUOF_TB_EIP_PRIV_MESS : Form
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
        string tablename = null;

        int result;


        FileInfo info;
        string[] tempFile;
        string tFileName = "";

        string PHOTO_TOPIC = "";
        string PHOTO_DESC = "";
        string RESIZE_FILE_ID = "";

        public frmUOF_TB_EIP_PRIV_MESS()
        {
            InitializeComponent();

            SETTEXT();
        }
        private void frmUOF_TB_EIP_PRIV_MESS_Load(object sender, EventArgs e)
        {
            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;
            dataGridView3.AlternatingRowsDefaultCellStyle.BackColor = Color.PaleTurquoise;


            //先建立個 CheckBox 欄
            DataGridViewCheckBoxColumn cbCol = new DataGridViewCheckBoxColumn();
            cbCol.Width = 20;   //設定寬度
            cbCol.HeaderText = "　全選";
            cbCol.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;   //置中
            cbCol.TrueValue = true;
            cbCol.FalseValue = false;
            dataGridView3.Columns.Insert(0, cbCol);


            //建立个矩形，等下计算 CheckBox 嵌入 GridView 的位置
            Rectangle rect = dataGridView3.GetCellDisplayRectangle(0, -1, true);
            rect.X = rect.Location.X + rect.Width / 4 - 8;
            rect.Y = rect.Location.Y + (rect.Height / 2 - 1);

            CheckBox cbHeader = new CheckBox();
            cbHeader.Name = "checkboxHeader";
            cbHeader.Size = new Size(18, 18);
            cbHeader.Location = rect.Location;

            //全选要设定的事件
            cbHeader.CheckedChanged += new EventHandler(cbHeader_CheckedChanged);

            //将 CheckBox 加入到 dataGridView
            dataGridView3.Controls.Add(cbHeader);
        }
        #region FUNCTION
        private void cbHeader_CheckedChanged(object sender, EventArgs e)
        {
            dataGridView3.EndEdit();

            foreach (DataGridViewRow dr in dataGridView3.Rows)
            {
                dr.Cells[0].Value = ((CheckBox)dataGridView3.Controls.Find("checkboxHeader", true)[0]).Checked;

            }

        }
        public void SETTEXT()
        {
            textBox7.Text = DateTime.Now.Month.ToString();
        }

        public void SearchUOFTB_EIP_PRIV_MESS(string ACCOUNT)
        {
            SqlDataAdapter adapter = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
            DataSet ds = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sbSql.Clear();
                sbSqlQuery.Clear();

                if(!string.IsNullOrEmpty(ACCOUNT))
                {
                    sbSqlQuery.AppendFormat(@" 
                                            AND ACCOUNT LIKE '{0}%'
                                            ", ACCOUNT);
                }

                sbSql.AppendFormat(@"  
                                   SELECT ACCOUNT AS '工號',NAME AS '姓名',USER_GUID
                                    FROM [192.168.1.223].[UOF].dbo.TB_EB_USER
                                    WHERE 1=1
                                    AND IS_SUSPENDED='0'
                                    AND(ACCOUNT LIKE '0%' OR ACCOUNT LIKE '1%' OR ACCOUNT LIKE '2%' OR ACCOUNT LIKE '3%' OR ACCOUNT LIKE '4%' OR ACCOUNT LIKE '5%' OR ACCOUNT LIKE '6%' OR ACCOUNT LIKE '7%' OR ACCOUNT LIKE '8%' OR ACCOUNT LIKE '9%' )
                                    {0}
                                    ORDER BY ACCOUNT

                                    ", sbSqlQuery.ToString());

                adapter = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder = new SqlCommandBuilder(adapter);
                sqlConn.Open();
                ds.Clear();
                adapter.Fill(ds, "TEMPds1");
                sqlConn.Close();


                if (ds.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView1.DataSource = null;
                }
                else
                {
                    if (ds.Tables["TEMPds1"].Rows.Count >= 1)
                    {
                        dataGridView1.DataSource = ds.Tables["TEMPds1"];
                        dataGridView1.AutoResizeColumns();

                        dataGridView1.Columns["工號"].Width = 100;
                        dataGridView1.Columns["姓名"].Width = 100;
                      
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

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            textBox4.Text = null;
            textBox5.Text = null;

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    textBox4.Text = row.Cells["USER_GUID"].Value.ToString().Trim();
                    textBox5.Text = row.Cells["姓名"].Value.ToString().Trim();



                }
                else
                {
                    textBox4.Text = null;
                    textBox5.Text = null;
                }
            }
        }

        public void ADDToUOF_TB_EIP_PRIV_MESS(string MESSAGE_TO, string MESSAGE_FROM,string TOPIC, string CONTENT)
        {
            Guid MESSAGE_GUID = Guid.NewGuid();
            Guid MASTER_GUID = Guid.NewGuid();
            Guid NOTIFY_ID = Guid.NewGuid();

            StringBuilder SBTEXT = new StringBuilder();

            MESSAGE_TO = "b6f50a95-17ec-47f2-b842-4ad12512b431";
            MESSAGE_FROM = "b6f50a95-17ec-47f2-b842-4ad12512b431";
            TOPIC = "TEST";
            //SBTEXT.AppendFormat(@"
            //                    <p style=""font-size:160%;color:red;"">This is a paragraph.</p>
            //                    <br></br>
            //                    <p style=""font-size:160%;color:blue;"">This is a paragraph.</p>
            //                      ");
            //CONTENT = SBTEXT.ToString();

            string CREATOR = MESSAGE_FROM;
            string MODIFIER = MESSAGE_FROM;
            string MESSAGE_TOUSER = @"<UserSet><Element type=""user""><userId>b6f50a95-17ec-47f2-b842-4ad12512b431</userId></Element></UserSet>";
            string MESSAGE_CONTENT = CONTENT;
            string TB_EIP_PRIV_MESS_CREATE_TIME= DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss ") + "+08:00";
            string CREATE_TIME = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string SENDER_TIME = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string CREATE_FROM = "192.168.1.57";
            string MODIFY_FROM = "192.168.1.57";
            string CREATE_DATE = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string MODIFY_DATE = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            string USER_GUID = MESSAGE_FROM;
            string TITLE = TOPIC;

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);



                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();
                
                sbSql.AppendFormat(@" 

                                    INSERT [UOF].dbo.TB_EIP_PRIV_MESS
                                    ( 
                                    MESSAGE_GUID
                                    , TOPIC
                                    , MESSAGE_CONTENT
                                    , MESSAGE_TO
                                    , MESSAGE_FROM
                                    , CREATE_TIME
                                    , FROM_DELETED
                                    , TO_DELETED
                                    , FILE_GROUP_ID
                                    , MASTER_GUID 
                                    ) 
                                    VALUES 
                                    ( 
                                    @MESSAGE_GUID
                                    , @TOPIC
                                    , @MESSAGE_CONTENT
                                    , @MESSAGE_TO
                                    , @MESSAGE_FROM
                                    , @TB_EIP_PRIV_MESS_CREATE_TIME
                                    , 0
                                    , 0
                                    , N''
                                    , @MASTER_GUID
                                    )



                                    INSERT [UOF].dbo.TB_EIP_PRIV_MESS_MASTER
                                    ( 
                                    MASTER_GUID
                                    , TOPIC
                                    , MESSAGE_FROM
                                    , MESSAGE_TO
                                    , SENDER_TIME
                                    , CREATOR
                                    , CREATE_FROM
                                    , CREATE_DATE
                                    , MODIFIER
                                    , MODIFY_FROM
                                    , MODIFY_DATE
                                    ) 
                                    VALUES 
                                    (  
                                    @MESSAGE_GUID
                                    ,@TOPIC
                                    ,@MESSAGE_FROM
                                    ,@MESSAGE_TOUSER
                                    ,@SENDER_TIME
                                    ,@CREATOR 
                                    ,@CREATE_FROM
                                    ,@CREATE_DATE
                                    ,@MODIFIER
                                    ,@MODIFY_FROM
                                    ,@MODIFY_DATE
                                    )



                                    INSERT TB_EIP_PUSH_QUEUE
                                    ( 
                                    [NOTIFY_ID]
                                    , [USER_GUID]
                                    , [DESCRIPTION]
                                    , [TITLE]
                                    , [DISPLAY_NUMBER]
                                    , [MODULE_NAME]
                                    , [MODULE_KEY]
                                    , [CREATOR]
                                    , [CREATE_FROM]
                                    , [CREATE_DATE]
                                    , [MODIFIER]
                                    , [MODIFY_FROM]
                                    , [MODIFY_DATE]
                                    )
                                    VALUES
                                    (
                                    @NOTIFY_ID
                                    , @USER_GUID
                                    , N''
                                    ,@TITLE
                                    , 1,
                                    N'PrivateMessage'
                                    , N'PrivateMessage?id=@MESSAGE_GUID'
                                    ,@CREATOR
                                    , @CREATE_FROM
                                    , @CREATE_DATE
                                    , @MODIFIER
                                    , @MODIFY_FROM
                                    , @MODIFY_DATE
                                    )

                                    ");


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();

                cmd.Parameters.AddWithValue("@MESSAGE_GUID", MESSAGE_GUID);
                cmd.Parameters.AddWithValue("@MASTER_GUID", MASTER_GUID);
                cmd.Parameters.AddWithValue("@NOTIFY_ID", NOTIFY_ID);
                cmd.Parameters.AddWithValue("@MESSAGE_FROM", MESSAGE_FROM);
                cmd.Parameters.AddWithValue("@MESSAGE_TO", MESSAGE_TO);
                cmd.Parameters.AddWithValue("@MESSAGE_TOUSER", MESSAGE_TOUSER);
                cmd.Parameters.AddWithValue("@TOPIC", TOPIC);
                cmd.Parameters.AddWithValue("@MESSAGE_CONTENT", MESSAGE_CONTENT);
                cmd.Parameters.AddWithValue("@CREATE_TIME", CREATE_TIME);
                cmd.Parameters.AddWithValue("@SENDER_TIME", SENDER_TIME);
                cmd.Parameters.AddWithValue("@CREATOR", CREATOR);
                cmd.Parameters.AddWithValue("@MODIFIER", MODIFIER);
                cmd.Parameters.AddWithValue("@CREATE_FROM", CREATE_FROM);
                cmd.Parameters.AddWithValue("@MODIFY_FROM", MODIFY_FROM);
                cmd.Parameters.AddWithValue("@CREATE_DATE", CREATE_DATE);
                cmd.Parameters.AddWithValue("@MODIFY_DATE", MODIFY_DATE);
                cmd.Parameters.AddWithValue("@TITLE", TITLE);
                cmd.Parameters.AddWithValue("@USER_GUID", USER_GUID);
                cmd.Parameters.AddWithValue("@TB_EIP_PRIV_MESS_CREATE_TIME", TB_EIP_PRIV_MESS_CREATE_TIME);


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

        public void SETTEXTBOX3()
        {
            StringBuilder SBTEXT = new StringBuilder();

            SBTEXT.AppendFormat(@"
                                <p style=""font-size:160%;color:red;"">This is a paragraph.</p>
                                <br></br>
                                <p style=""font-size:200%;color:blue;"">This is a paragraph.</p>
                                  ");

            textBox3.Text = SBTEXT.ToString();

        }


        public void SEARCHPIC(string NAME)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    
                                    SELECT 
                                    [TB_EIP_ALBUM_CLASS].[CLASS_NAME] AS '分類'
                                    ,[TB_EIP_ALBUM].[ALBUM_TOPIC] AS '主題'
                                    ,[PHOTO_TOPIC] AS '照片名稱'
                                    ,[TB_EIP_ALBUM_CLASS].[CLASS_GUID]
                                    ,[TB_EIP_ALBUM].[ALBUM_GUID]
                                    ,[PHOTO_GUID]
                                    ,[FILE_ID]
                                    ,[THUMBNAIL_FILE_ID]
                                    ,[PHOTO_DESC]
                                    ,[FRONT_COVER]
                                    ,[COMMEND_COUNT]
                                    ,[RESIZE_FILE_ID]
                                    FROM [UOF].[dbo].[TB_EIP_ALBUM_CLASS], [UOF].[dbo].[TB_EIP_ALBUM],[UOF].[dbo].[TB_EIP_ALBUM_PHOTO]
                                    WHERE 1=1
                                    AND [TB_EIP_ALBUM_CLASS].CLASS_GUID=[TB_EIP_ALBUM].CLASS_GUID
                                    AND [TB_EIP_ALBUM].ALBUM_GUID=[TB_EIP_ALBUM_PHOTO].ALBUM_GUID
                                    AND [CLASS_NAME] LIKE '%賀圖區%'
                                    AND [PHOTO_TOPIC] LIKE '%{0}%'
                                    ORDER BY [PHOTO_TOPIC]

                                    ", NAME);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView2.DataSource = null;
                }
                else if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView2.DataSource = ds1.Tables["TEMPds1"];

                    dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView2.Columns["分類"].Width = 100;
                    dataGridView2.Columns["主題"].Width = 100;
                    dataGridView2.Columns["照片名稱"].Width = 100;

                }


            }
            catch
            {

            }
            finally
            {

            }
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView2.CurrentRow != null)
                {
                    int rowindex = dataGridView2.CurrentRow.Index;

                    if (rowindex >= 0)
                    {
                        DataGridViewRow row = dataGridView2.Rows[rowindex];

                        PHOTO_TOPIC = row.Cells["照片名稱"].Value.ToString();
                        PHOTO_DESC = row.Cells["PHOTO_DESC"].Value.ToString();
                        RESIZE_FILE_ID = row.Cells["RESIZE_FILE_ID"].Value.ToString();

                        Image O_Image = Image.FromStream(WebRequest.Create("https://eip.tkfood.com.tw/UOF/Common/FileCenter/V3/Handler/FileControlHandler.ashx?id=" + RESIZE_FILE_ID + "").GetResponse().GetResponseStream());
                        //将获取的图片赋给图片框
                        pictureBox1.Image = O_Image;
                        pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;

                    }
                    else
                    {


                    }
                }
            }
            catch
            {

            }

        }

        public void SEARCHUSER(string MONTHS)
        {
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);


                sbSql.Clear();
                sbSqlQuery.Clear();


                sbSql.AppendFormat(@"  
                                    
                                    SELECT TB_EB_USER.ACCOUNT AS '工號',TB_EB_USER.NAME AS '姓名',TB_EB_EMPL.BIRTHDAY AS '生日',TB_EB_USER.USER_GUID,MONTH(TB_EB_EMPL.BIRTHDAY) AS BIRTHMONTHS
                                    FROM [UOF].[dbo].TB_EB_USER,[UOF].[dbo].TB_EB_EMPL
                                    WHERE 1=1
                                    AND TB_EB_USER.USER_GUID=TB_EB_EMPL.USER_GUID
                                    AND TB_EB_USER.IS_SUSPENDED<>'1'
                                    AND ISNULL(TB_EB_EMPL.BIRTHDAY,'')<>''
                                    AND MONTH(TB_EB_EMPL.BIRTHDAY)={0}
                                    ORDER BY NAME

                                    ", MONTHS);

                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count == 0)
                {
                    dataGridView3.DataSource = null;
                }
                else if (ds1.Tables["TEMPds1"].Rows.Count >= 1)
                {
                    dataGridView3.DataSource = ds1.Tables["TEMPds1"];

                    dataGridView3.ColumnHeadersDefaultCellStyle.Font = new Font("Tahoma", 9);
                    dataGridView3.DefaultCellStyle.Font = new Font("Tahoma", 10);
                    dataGridView3.Columns["工號"].Width = 100;
                    dataGridView3.Columns["姓名"].Width = 100;
                    dataGridView3.Columns["生日"].Width = 100;

                }


            }
            catch
            {

            }
            finally
            {

            }
        }

        public void NEW_MESSAGE()
        {
            string MESSAGE_TO = "";
            string MESSAGE_FROM = "916e213c-7b2e-46e3-8821-b7066378042b";

            StringBuilder TEXTBOX = new StringBuilder();
            for (int i = 0; i < textBox8.Lines.Length; i++)
            {
                TEXTBOX.AppendFormat("<p>" + textBox8.Lines[i] + "</p>");
            }

            foreach (DataGridViewRow DR in dataGridView3.Rows)
            {
                if (Convert.ToBoolean(DR.Cells[0].Value) == true)
                {
                    MESSAGE_TO = DR.Cells["USER_GUID"].Value.ToString();

                    StringBuilder MESSAGE_CONTENT = new StringBuilder();

                    MESSAGE_CONTENT.AppendFormat(TEXTBOX.ToString());
                    MESSAGE_CONTENT.AppendFormat(@"                                               
                                                <p>[img alt="""" src=""/common/FileCenter/V3/Handler/FileControlHandler.ashx?id={0}""class=""UOF"" style=""width: 331px; "" /]</p>
                                              
                                                ", RESIZE_FILE_ID);


                    ADD_UOF_TB_EIP_PRIV_MESS(
                    Guid.NewGuid().ToString("")
                    , "HR TEST"
                    , MESSAGE_CONTENT.ToString()
                    , MESSAGE_TO
                    , MESSAGE_FROM
                    , ""
                    , DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK")
                    , ""
                    , ""
                    , "0"
                    , "0"
                    , ""
                    , Guid.NewGuid().ToString("")
                    , ""
                    );
                }
            }
            //StringBuilder MESSAGE_CONTENT = new StringBuilder();
            //MESSAGE_CONTENT.AppendFormat(@"
            //                            <p>HR TEST</p>
            //                            <p>&nbsp;</p>
            //                            <p>[img alt="""" src=""/common/filecenter/v3/handler/downloadhandler.ashx?id=150fff01-47d5-4b97-a6a2-76c7207fa395&path=ALBUM%5C2022%5C11&contentType=image%2Fjpeg&name=40100331068090.jpg&e=HU1s3YUxz%2f%2f%2f59HuM52yYHkLtMC3WfMTCVazCg9KbOfjc2pxNV2dVM1j%2btqCuPZK&l=Nxv%2b0JZKKdGc8%2fv6xuCvtDw0QbJcGvHE9nd1Vbm8zaQ%3d&enc=0&uid=b6f50a95-17ec-47f2-b842-4ad12512b431"" class=""UOF"" style=""width: 331px; "" /]</p>
            //                            HI~
            //                            ");


            //ADD_UOF_TB_EIP_PRIV_MESS(
            //Guid.NewGuid().ToString("")
            //, "test"
            //, MESSAGE_CONTENT.ToString()
            //, "b6f50a95-17ec-47f2-b842-4ad12512b431"
            //, "b6f50a95-17ec-47f2-b842-4ad12512b431"
            //, ""
            //, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fffffffK")
            //, ""
            //, ""
            //, "0"
            //, "0"
            //, ""
            //, Guid.NewGuid().ToString("")
            //, ""
            //);
        }

        public void ADD_UOF_TB_EIP_PRIV_MESS(
            string MESSAGE_GUID
            , string TOPIC
            , string MESSAGE_CONTENT
            , string MESSAGE_TO
            , string MESSAGE_FROM
            , string REPLY_MESSAGE_GUID
            , string CREATE_TIME
            , string READED_TIME
            , string REPLY_TIME
            , string FROM_DELETED
            , string TO_DELETED
            , string FILE_GROUP_ID
            , string MASTER_GUID
            , string EVENT_ID

            )
        {
            try
            {
                // 20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbUOF"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);
                using (SqlConnection conn = sqlConn)
                {
                    if (!string.IsNullOrEmpty(MESSAGE_TO))
                    {
                        StringBuilder SBSQL = new StringBuilder();
                        SBSQL.AppendFormat(@" 
                                            INSERT INTO [UOF].[dbo].[TB_EIP_PRIV_MESS]
                                            (
                                            [MESSAGE_GUID]
                                            ,[TOPIC]
                                            ,[MESSAGE_CONTENT]
                                            ,[MESSAGE_TO]
                                            ,[MESSAGE_FROM]
                                            ,[REPLY_MESSAGE_GUID]
                                            ,[CREATE_TIME]
                                            ,[READED_TIME]
                                            ,[REPLY_TIME]
                                            ,[FROM_DELETED]
                                            ,[TO_DELETED]
                                            ,[FILE_GROUP_ID]
                                            ,[MASTER_GUID]
                                            ,[EVENT_ID]
                                            )
                                            VALUES
                                            (
                                            NEWID()
                                            ,@TOPIC
                                            ,@MESSAGE_CONTENT
                                            ,@MESSAGE_TO
                                            ,@MESSAGE_FROM
                                            ,''
                                            ,@CREATE_TIME
                                            ,NULL
                                            ,NULL
                                            ,'0'
                                            ,'0'
                                            ,''
                                            ,NEWID()
                                            ,''
                                            )

                                            ");

                        string sql = SBSQL.ToString();

                        using (SqlCommand cmd = new SqlCommand(sql, conn))
                        {

                            cmd.Parameters.AddWithValue("@MESSAGE_GUID", MESSAGE_GUID);
                            cmd.Parameters.AddWithValue("@TOPIC", TOPIC);
                            cmd.Parameters.AddWithValue("@MESSAGE_CONTENT", MESSAGE_CONTENT);
                            cmd.Parameters.AddWithValue("@MESSAGE_TO", MESSAGE_TO);
                            cmd.Parameters.AddWithValue("@MESSAGE_FROM", MESSAGE_FROM);
                            cmd.Parameters.AddWithValue("@REPLY_MESSAGE_GUID", REPLY_MESSAGE_GUID);
                            cmd.Parameters.AddWithValue("@CREATE_TIME", CREATE_TIME);
                            cmd.Parameters.AddWithValue("@READED_TIME", READED_TIME);
                            cmd.Parameters.AddWithValue("@REPLY_TIME", REPLY_TIME);
                            cmd.Parameters.AddWithValue("@FROM_DELETED", FROM_DELETED);
                            cmd.Parameters.AddWithValue("@TO_DELETED", TO_DELETED);
                            cmd.Parameters.AddWithValue("@FILE_GROUP_ID", FILE_GROUP_ID);
                            cmd.Parameters.AddWithValue("@MASTER_GUID", MASTER_GUID);




                            conn.Open();
                            cmd.ExecuteNonQuery();
                            conn.Close();
                        }
                    }

                }
            }
            catch
            {
                MessageBox.Show("失敗");
            }
        }


        #endregion

        #region BUTTON
        private void button1_Click(object sender, EventArgs e)
        {
            SearchUOFTB_EIP_PRIV_MESS(textBox1.Text.ToString());
        }
        private void button2_Click(object sender, EventArgs e)
        {
            ADDToUOF_TB_EIP_PRIV_MESS("","","", textBox3.Text);

            MessageBox.Show("完成");
        }
        private void button3_Click(object sender, EventArgs e)
        {
            SETTEXTBOX3();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SEARCHPIC(textBox6.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SEARCHUSER(textBox7.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            NEW_MESSAGE();

            MessageBox.Show("完成");
        }



        #endregion

      
    }
}
