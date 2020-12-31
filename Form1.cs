using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Web;
using CrystalDecisions.Shared;
using System.Configuration;
using System.IO;
using System.Web;
using System.Threading;
using System.Data.SqlClient;
using System.Linq;
using System.Web;
using System.Web.Hosting;
using System;
using System.Collections.Generic;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Drawing;
using System.Net;
using System.IO;
using System.Collections;
using System.Text;
using System.Diagnostics;
using System.Configuration;
using System.Data.SqlClient;
//using Renci.SshNet;
//using Renci.SshNet.Common;
//using Renci.SshNet.Sftp;
using System.Text.RegularExpressions;
using System.Linq;
using System.Data.OleDb;
//using System.Runtime.InteropServices.ExternalException;
using System;
using System.Security.Principal;
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace soa
{      
    public partial class Form1 : Form
    {
        int complete = 0;
        bool v=false;
        bool m = true;
        int cr1 = 0;
        int cr2 = 0;
      
        public Form1()
        {
            InitializeComponent();
        }
        private void button1_Click(object sender, EventArgs e)
        {

            //   string uploadfile = HostingEnvironment.MapPath("~/report_load/AccountStatement_DWH_PF_Portrait_Email_New_Daily.rpt");
                      
          //  var watch = System.Diagnostics.Stopwatch.StartNew();      
            try
            {

                //////////////////////////////////////////////

                clsDBOperation5 clsdb = new clsDBOperation5();

                bool purchase = false;
                bool non_purchase = false;
                bool purchase_mraf = false;
                ///////////////////////////////////////////////


                //String executingFolder = AppDomain.CurrentDomain.BaseDirectory;
                //executingFolder = System.IO.Path.Combine(executingFolder, @"report_load\AccountStatement_DWH_PF_Portrait_Email_New_Daily.rpt");                
                string START_DATE_QUERY = "SELECT daily_statement_startdate FROM StatementDate";


                string connetionString_DATE = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn_SDATE;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn_SDATE = new SqlConnection(connetionString_DATE);
                cnn_SDATE.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd_DATE = new SqlCommand(START_DATE_QUERY, cnn_SDATE);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                SqlDataReader drCN_DATE = cmd_DATE.ExecuteReader();
                DataTable dtCN_DATE = new DataTable();
                dtCN_DATE.Load(drCN_DATE);
                int sDATE = dtCN_DATE.Rows.Count;
                cnn_SDATE.Close();

                string SDATE = dtCN_DATE.Rows[0][0].ToString();

                DateTime startdate = Convert.ToDateTime(SDATE);

                //  DateTime startdate = DateTime.ParseExact(SDATE,"yyyy-MM-dd",System.Globalization.CultureInfo.InvariantCulture);

                DateTime query_date = dateTimePicker1.Value.Date;
                // DateTime startdate =new DateTime (2017,07,01);
                DateTime enddate = dateTimePicker2.Value.Date;
                //  DateTime enddate = new DateTime(2018, 03, 02); 
               DateTime process_date = dateTimePicker3.Value.Date;

                DateTime endate_limit_time = enddate.AddDays(1);


                DateTime letter_date = dateTimePicker1.Value.Date;

                string report_query = "select distinct Portfolio_ID  from Customer_Trades ct";
                //report_query = report_query + " where TradeDateTime >= '2017-10-23'";
                //report_query = report_query + " and TradeDateTime <= '2017-10-24'";
                //report_query = report_query + " and flag = 0 and email like '%@%'";
                //report_query = report_query + " and TransType like '%FR%'";
                //report_query = report_query + " and TransDesc !='PURCHASE'";
                report_query = report_query + " inner join Customer c on ct.PortfolioID = c.Portfolio_ID ";
                report_query = report_query + " where TradeDateTime BETWEEN " + "' " + query_date.ToString("yyyyMMdd") + "'" + " AND " + "'" + enddate.ToString("yyyyMMdd") + "'";
                report_query = report_query + " and c.Email like'%@%' and Flag=0 and ct.customerid not in (select customerid from MemberBlockCustomer)";
                //report_query = report_query + " and Portfolio_ID='132449-1'";   
                //report_query = report_query + " where TradeDateTime BETWEEN  '2017-11-17' AND '2017-11-20' and Flag=0 and c.Email like'%@%'";



                if (checkBox1.Checked == true)
                {
                    // report_query = report_query + " and   TransDesc='DIVIDEND INVEST'";
                    report_query = "";
                    report_query = "select distinct Portfolio_ID  from Customer_Trades ct";
                    report_query = report_query + " inner join Customer c on ct.PortfolioID = c.Portfolio_ID ";
                    report_query = report_query + " where TradeDateTime BETWEEN " + "'" + query_date.ToString("yyyyMMdd") + "'" + " AND " + "'" + enddate.ToString("yyyyMMdd") + "'";
                    report_query = report_query + " and c.Email like'%@%' and Flag=0 and ct.customerid not in (select customerid  from MemberBlockCustomer) ";
                    report_query = report_query + " and ct.time_stamp > =(SELECT LastSOATime FROM  StatementDate)";
                    report_query = report_query + " and ct.time_stamp <'" + endate_limit_time.ToString("yyyy-MM-dd hh:mm:ss.fff") + "'";
                   
                }

                string letter_query = "select distinct Portfolio_ID from Customer_Trades ct";
                letter_query = letter_query + " inner join Customer c on ct.PortfolioID = c.Portfolio_ID";
                letter_query = letter_query + " where TradeDateTime BETWEEN " + "' " + query_date.ToString("yyyyMMdd") + "'" + " AND " + "'" + enddate.ToString("yyyyMMdd") + "'";
                letter_query = letter_query + "and Flag=0 and c.Email like'%@%' and transdesc in('PURCHASE')  and ct.customerid not in (select customerid  from MemberBlockCustomer) ";

                if (checkBox1.Checked == true)
                {
                    letter_query = "";
                    letter_query = "select distinct Portfolio_ID from Customer_Trades ct";
                    letter_query = letter_query + " inner join Customer c on ct.PortfolioID = c.Portfolio_ID";
                    letter_query = letter_query + " where TradeDateTime BETWEEN " + "' " + query_date.ToString("yyyyMMdd") + "'" + " AND " + "'" + enddate.ToString("yyyyMMdd") + "'";
                    letter_query = letter_query + "and Flag=0 and c.Email like'%@%' and transdesc in('PURCHASE') and ct.customerid not in (select customerid  from MemberBlockCustomer)  ";
                    letter_query = letter_query + " and ct.time_stamp > =(SELECT LastSOATime FROM  StatementDate)";
                    letter_query = letter_query + " and ct.time_stamp <'" + endate_limit_time.ToString("yyyy-MM-dd hh:mm:ss.fff") + "'";

                }





                //enddate = new DateTime(2018, 03, 02); 

                //  letter_query = letter_query + " where TradeDateTime BETWEEN  '2017-11-17' AND '2017-11-20' and Flag=0 and c.Email like'%@%' and transdesc in('PURCHASE')";


                //string report_query="SELECT distinct Portfolio_ID from Customer_Trades ct inner join Customer c on ct.PortfolioID=c.Portfolio_ID" ;
                //report_query=report_query+" where TradeDateTime BETWEEN '2017-11-01'AND '2017-11-02' ";
                //report_query=report_query+" and Flag=0 and c.Email like'%@%' and transdesc in('PURCHASE') " ;


                string bal_query = "exec sp_TrdBalSPBalDiff '" + process_date.ToString("yyyyMMdd") + "' ";

                string connetionString1 = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn1;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn1 = new SqlConnection(connetionString1);
                cnn1.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd12 = new SqlCommand(bal_query, cnn1);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                SqlDataReader drCN1 = cmd12.ExecuteReader();
                DataTable dtCN1 = new DataTable();
                dtCN1.Load(drCN1);
                int bal_count =  dtCN1.Rows.Count;
                cnn1.Close();


                if (bal_count > 0) 
                {
                    DialogResult result = MessageBox.Show("Do you really want to proceed?", "Dialog Title", MessageBoxButtons.YesNo);
                    if (result == DialogResult.No)
                    {
                        Application.ExitThread();
                        Environment.Exit(0);
                    }
                    else
                    {
                      //  return;
                    }
                    
                }
 
                string connetionString = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn = new SqlConnection(connetionString);
                cnn.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd = new SqlCommand(report_query, cnn);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                SqlDataReader drCN = cmd.ExecuteReader();
                DataTable dtCN = new DataTable();
                dtCN.Load(drCN);
                int s = dtCN.Rows.Count;
                cnn.Close();


                //  MessageBox.Show(s.ToString());


                string connetionString_letter = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn_letter;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn_letter = new SqlConnection(connetionString_letter);
                cnn_letter.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd_letter = new SqlCommand(letter_query, cnn_letter);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                SqlDataReader drCN_letter = cmd_letter.ExecuteReader();
                DataTable dtCN_letter = new DataTable();
                dtCN_letter.Load(drCN_letter);
                int s_letter = dtCN_letter.Rows.Count;
                cnn_letter.Close();

                int total = s_letter + s;
                //////////////////////////////////////////////////////////////

                clsdb.remove_trade_details("Delete From soa_daily_trades");
                clsdb.remove_trade_details("Truncate table SP_Daily");   
                
                
                /// table will be refill at every new instance of application
                
                clsdb.fetch_trades_details(startdate.ToString("yyyyMMdd"), query_date.ToString("yyyyMMdd"), enddate.ToString("yyyyMMdd"));

                clsdb.insert_customer_sp(report_query);



                DataTable record_check = new DataTable();
                record_check = clsdb.GetDataTable("select * from EmailPool_BulkPrinting where flag='n'");

                if (checkBox1.Checked == false)
                {
                    if (record_check.Rows.Count > 1)
                    {
                        non_purchase = clsdb.non_purchase_record(query_date.ToString("yyyyMMdd"), enddate.ToString("yyyyMMdd"));
                        purchase = true;
                    }
                    else
                    {
                        purchase = clsdb.purchase_record(query_date.ToString("yyyyMMdd"), enddate.ToString("yyyyMMdd"));
                        non_purchase = clsdb.non_purchase_record(query_date.ToString("yyyyMMdd"), enddate.ToString("yyyyMMdd"));
                    }

                }

                else
                {
                    cnn.Open();
                    SqlCommand cmd1 = new SqlCommand("delete from Soa_Daily_SOA_Morning", cnn);    // to be corrected                
                    cmd1.ExecuteReader();
                    cnn.Close();
                    // to avoid connection creation in loop

                    if (s_letter > 0)
                    {
                        for (int i = 0; i < s_letter; i++)
                        {
                            cnn.Open();
                            string query_fill = "insert into Soa_Daily_SOA_Morning values ('" + dtCN_letter.Rows[i][0].ToString() + "')";
                            SqlCommand cmd2 = new SqlCommand(query_fill, cnn);    // to be corrected                     
                            cmd2.ExecuteReader();
                            cnn.Close();
                        }

                        purchase_mraf = clsdb.purchase_record_Mraf(query_date.ToString("yyyyMMdd"), enddate.ToString("yyyyMMdd"));
                        purchase = true;
                    }


                    cnn.Open();
                    cmd1 = new SqlCommand("delete from Soa_Daily_SOA_Morning", cnn);    // to be corrected                
                    cmd1.ExecuteReader();
                    cnn.Close();
                    
                    for (int i = 0; i < s; i++)
                    {
                        cnn.Open();
                        string query_fill = "insert into Soa_Daily_SOA_Morning values ('" + dtCN.Rows[i][0].ToString() + "')";
                        SqlCommand cmd2 = new SqlCommand(query_fill, cnn);    // to be corrected                     
                        cmd2.ExecuteReader();
                        cnn.Close();
                    }

                    non_purchase = clsdb.non_purchase_record_Mraf(query_date.ToString("yyyyMMdd"), enddate.ToString("yyyyMMdd"));
                    purchase = true;
                }

                if (checkBox1.Checked == true && non_purchase == true && purchase_mraf==true)
                {
                    MessageBox.Show("Number of Statements to be generated #  " + s.ToString() + "\n" + "\n" + "Number of Letters to be generated  #  " + s_letter.ToString()  + "\n" + "\n" + "Record Transfer Successfull");
                }

                if (checkBox1.Checked == true && (non_purchase == false||(s_letter>0 && purchase_mraf==false)))
                {
                    MessageBox.Show("Number of Statements to be generated #  " + s.ToString() + "\n" + "\n" + "Number of Letters to be generated  #  " + s_letter.ToString() +"\n" + "\n" + "Record Transfer Un-Successfull");
                }

                if ((checkBox1.Checked == false) && ((non_purchase == true) && (purchase == true)))
                {
                    MessageBox.Show("Number of Statements to be generated #  " + s.ToString() + "\n" + "\n" + "Number of Letters to be generated  #  " + s_letter.ToString() + "\n" + "\n" + "Record Transfer Successfull");
                }
                if ((checkBox1.Checked == false) && ((non_purchase == false) || (purchase == false)))
                {
                    MessageBox.Show("Number of Statements to be generated #  " + s.ToString() + "\n" + "\n" + "Number of Letters to be generated  #  " + s_letter.ToString() + "\n" + "\n" + "Record Transfer Un-Successfull");
                }


                string load_report ="";
                string load_report_1 ="";
                string load_report_2 ="";
                string load_report_3 ="";
                string load_report_4 ="";
                string load_report_5 ="";
                string load_report_6 ="";
                string load_report_letter = "";
                string _servername ="";
                string _databasename ="";
                string _userid ="";
                string _password ="";
                string _databasename1 ="";


                if (checkBox1.Checked == true)
                {
                    load_report = ConfigurationManager.AppSettings["load_report"];      //--    @"E:\SOA\reports_soa\AccountStatement_DWH_PF_Portrait_Email_New_Daily.rpt";
                    load_report_1 = ConfigurationManager.AppSettings["load_report"];
                    load_report_2 = ConfigurationManager.AppSettings["load_report"];
                    load_report_3 = ConfigurationManager.AppSettings["load_report"];
                    load_report_4 = ConfigurationManager.AppSettings["load_report"];
                    load_report_5 = ConfigurationManager.AppSettings["load_report"];
                    load_report_6 = ConfigurationManager.AppSettings["load_report"];
                    load_report_letter = ConfigurationManager.AppSettings["load_report_letter"];
                    _servername = ConfigurationManager.AppSettings["servername"];
                    _databasename = ConfigurationManager.AppSettings["databasename"];
                    _userid = ConfigurationManager.AppSettings["userid"];
                    _password = ConfigurationManager.AppSettings["password"];
                    _databasename1 = ConfigurationManager.AppSettings["databasename1"];
                }
                else
                {
                    load_report = ConfigurationManager.AppSettings["load_report"];      //--    @"E:\SOA\reports_soa\AccountStatement_DWH_PF_Portrait_Email_New_Daily.rpt";
                    load_report_1 = ConfigurationManager.AppSettings["load_report"];
                    load_report_2 = ConfigurationManager.AppSettings["load_report"];
                    load_report_3 = ConfigurationManager.AppSettings["load_report"];
                    load_report_4 = ConfigurationManager.AppSettings["load_report"];
                    load_report_5 = ConfigurationManager.AppSettings["load_report"];
                    load_report_6 = ConfigurationManager.AppSettings["load_report"];
                    load_report_letter = ConfigurationManager.AppSettings["load_report_letter"];
                    _servername = ConfigurationManager.AppSettings["servername"];
                    _databasename = ConfigurationManager.AppSettings["databasename"];
                    _userid = ConfigurationManager.AppSettings["userid"];
                    _password = ConfigurationManager.AppSettings["password"];
                    _databasename1 = ConfigurationManager.AppSettings["databasename1"];
                }


                string[] items = new string[s];
                //items = dtCN.Rows.OfType<DataRow>().Select(k => k[0].ToString()).ToArray();
                for (int i = 0; i < s; i++)
                // will replace the above linq (code) 
                {
                    items[i] = dtCN.Rows[i][0].ToString();
                }
                string[] items_letter = new string[s_letter];
                //items_letter = dtCN_letter.Rows.OfType<DataRow>().Select(k => k[0].ToString()).ToArray();
                for (int i = 0; i < s_letter; i++)
                // will replace the above linq (code) 
                {
                    items_letter[i] = dtCN_letter.Rows[i][0].ToString();
                }

                //  MessageBox.Show("data copied into array" + items[0].ToString() + "-----" + load_report + "-----" + load_report_letter);

                string t = items[0];
                int s1 = (s / 2) + 1;
                int p1 = (s / 9) + 1;
                int p2 = p1 + p1;
                int p3 = p2 + p1;
                int p4 = p3 + p1;
                int p5 = p4 + p1;
                int p6 = p5 + p1;
                int p7 = p6 + p1;
                int p8 = p7 + p1;

                progressBar1.Value = 0;


                if (checkBox1.Checked == true)
                {
                    progressBar1.Maximum = 9;
                }
                else 
                {

                    progressBar1.Maximum = 11;
                
                }


                progressBar1.Minimum = 0;
                //   string mesg = "3";
                //fund_type.DataSource = dtCN;
                //fund_type.DataTextField = "agent_name";
                //fund_type.DataValueField = "agent_name";
                //fund_type.DataBind();

                //DateTime letter_date=new DateTime(2017,11,17);

                MyThread report_thread = new MyThread();
                if (checkBox1.Checked == true)
                {
                    report_thread.div_check = true;
                }

                // report_thread.Thread2(s1, s, items, v);
                Thread thread1 =
                new Thread(() => report_thread.Thread1(0, p1, items, v, cr1, startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, load_report, load_report_letter, items_letter));
                Thread thread2 =
                new Thread(() => report_thread.Thread2(p1, p2, items, v, cr2, "2", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_1, load_report_letter, items_letter));
                Thread thread3 =
                new Thread(() => report_thread.Thread2(p2, p3, items, v, cr2, "3", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_2, load_report_letter, items_letter));
                Thread thread4 =
                new Thread(() => report_thread.Thread2(p3, p4, items, v, cr2, "4", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_3, load_report_letter, items_letter));
                Thread thread5 =
                new Thread(() => report_thread.Thread2(p4, p5, items, v, cr2, "5", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));
                Thread thread6 =
                new Thread(() => report_thread.Thread2(p5, p6, items, v, cr2, "6", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));
                Thread thread7 =
                new Thread(() => report_thread.Thread2(p6, p7, items, v, cr2, "7", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));
                Thread thread8 =
                new Thread(() => report_thread.Thread2(p7, p8, items, v, cr2, "8", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));
                Thread thread9 =
                new Thread(() => report_thread.Thread2(p8, s, items, v, cr2, "9", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));
                Thread thread10 =
                new Thread(() => report_thread.Thread3(0, s, items, v, cr2, "10", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));
                Thread thread11 =
                new Thread(() => report_thread.Thread3(s/2, s, items, v, cr2, "11", startdate, enddate, letter_date, _servername, _databasename, _userid, _password, _databasename, s_letter, load_report_4, load_report_letter, items_letter));





                button2.Enabled = true;
                button1.Enabled = false;
                checkBox1.Enabled = false;

                label3.Text = "Statement Generation started";
                label3.Visible = true;


                if (v == false)
                {
                    //////////////////////////////////////////// thread work is still in progress //////////////
                    //MyThread report_thread = new MyThread();
                    //// report_thread.Thread2(s1, s, items, v);
                    //Thread thread1 = new Thread(() => report_thread.Thread1(p3,s,items,v,cr1));
                    //Thread thread2 = new Thread(() => report_thread.Thread2(p3,p2,items,v,cr2,""));
                    //Thread thread3 = new Thread(() => report_thread.Thread2(p2, s, items, v, cr2, mesg));
                    //  thread1.Priority = ThreadPriority.AboveNormal;           
                    thread6.Start();
                    thread1.Start();
                    thread2.Start();
                    thread3.Start();
                    thread4.Start();
                    thread5.Start();
                    thread7.Start();
                    thread8.Start();
                    thread9.Start();
                    if (checkBox1.Checked != true)
                    {
                        thread10.Start();
                        thread11.Start();
                    }

                    else if (checkBox1.Checked == true && s_letter>0)
                    {
                        if (s_letter == 1) { thread10.Start(); }
                        else
                        {
                            thread10.Start();
                            thread11.Start();
                        }                    
                    }

                }


            }
              //  v = true;
                //watch.Stop();
                //string elapsedMs = watch.ElapsedMilliseconds.ToString();
                //MessageBox.Show("report generation completed !");   
            catch (Exception eec)
            {
                string s = eec.ToString();
                MessageBox.Show(eec.ToString());
            }
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click_1(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            //DialogResult result = MessageBox.Show("Do you really want to exit?", "Dialog Title", MessageBoxButtons.YesNo);
            //if (result == DialogResult.Yes)
            //{
            //    Application.ExitThread();
            //    Environment.Exit(0);
            //}
            //else
            //{
            //   // e.Cancel = true;
            //}
            Application.ExitThread();
            Environment.Exit(0);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                DialogResult result = MessageBox.Show("Do you really want to exit?", "Dialog Title", MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
                {
                    Application.ExitThread();
                    Environment.Exit(0);
                }
                else
                {
                    e.Cancel = true;
                }
            }
            else
            {
                e.Cancel = true;
            }
        
        
        }

        private void button3_Click(object sender, System.EventArgs e)
        {

            button2.Enabled = true;
            button1.Enabled = false;
            checkBox1.Enabled = false;
            button3.Enabled = false;

            try
            {

                string strquery = "select distinct portfolioid from vw_CGT_Refund where email like '%@%' order by PortfolioID";
                string _strConnectionString = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection oCnn = new SqlConnection(_strConnectionString);
                oCnn.Open();

                SqlDataAdapter oDA = new SqlDataAdapter(strquery, oCnn);
                oDA.SelectCommand.CommandType = CommandType.Text;
                oDA.SelectCommand.CommandText = strquery;

                DataSet oDS = new DataSet();
                DataTable dt = new DataTable();

                oDA.Fill(dt);
                oCnn.Close();

                string servername = ConfigurationManager.AppSettings["servername"];
                string databasename = ConfigurationManager.AppSettings["databasename"];
                string userid = ConfigurationManager.AppSettings["userid"];
                string password = ConfigurationManager.AppSettings["password"];
                string outputPath = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\CGT\Refund\Email\";
                string load_report = ConfigurationManager.AppSettings["load_report"].ToString();


                ReportDocument report = new ReportDocument();
                report.Load(ConfigurationManager.AppSettings["load_report_CGT_REFUND"].ToString());
                report.DataSourceConnections[0].SetConnection(servername,databasename,userid,password);
                foreach (DataRow PortfolioID in dt.Rows)
                {
                    report.SetParameterValue("@CustomerID", PortfolioID["PortfolioID"].ToString());
                   string  FilePath = outputPath + PortfolioID["PortfolioID"].ToString() + ".pdf";
                    if (!Directory.Exists(outputPath))
                    {
                        Directory.CreateDirectory(outputPath);
                    }
                    if (File.Exists(FilePath))
                    {
                        File.Delete(FilePath);
                    }
                    report.ExportToDisk(ExportFormatType.PortableDocFormat, FilePath);
                }

                report.Close();
            }
            catch (Exception)
            {
                throw;
            }
            

        }
    
    
    
    
    
    }

    public class MyThread : Form1
    {


        public bool div_check { get; set; }


        public void UpdateLabel(String text)
        {
            if (this.label1.InvokeRequired)
            {
                this.label1.BeginInvoke(new Action<string>(UpdateLabel), text);
                return;
            }
            //if (!this.IsHandleCreated)
            //{
            //    this.CreateHandle();
            //}
            this.label1.Text = text;
           // label1.Text = status;
           //this.label1.Invalidate();
           //this.label1.Update();
           //this.label1.Refresh();
           //this.label1.Refresh();
           // Application.DoEvents();
           //this.label1.Refresh();
        }

        public delegate void UpdateTextCallback();
        public void updateprogress() 
        {
            progressBar1.Value++;
            this.progressBar1.Invalidate();
            this.progressBar1.Update();
            this.progressBar1.Refresh();
            this.progressBar1.Refresh();
            Application.DoEvents();
        }
       // public delegate void MyDelegate();
        public void Thread1(int s1, int s, string[] items, bool v, int cr, DateTime start_date,DateTime end_date,DateTime letter_date,string _servername, string _databasename, string _userid, string _password, string _databasename1, string load_report, string load_report_letter_, string[] items_letter)
        {
            int y = 0;
            
            for (int h = 0; h < 20; h++)     // used to keep the program running in case of exception occurence  
            {            
            try
            {
                //string d1 = dateTimePicker1.Value.ToShortDateString();
                //string d2 = dateTimePicker2.Value.ToShortDateString();
                //DateTime startdate = DateTime.Parse(d1).Date;
                //DateTime enddate = DateTime.Parse(d2);
                //var startdate = startdate1.Date;
                //var enddate = enddate1.Date;
                //DateTime startdate = dateTimePicker1.Value.Date;
                //DateTime enddate = dateTimePicker2.Value.Date;
                //DateTime startdate = new DateTime(start_date.Year, start_date.Month, start_date.Day);
                //DateTime enddate = new DateTime(end_date.Year, end_date.Month,end_date.Day);
                              
                if (v == false)
                {
                    for (int i = 0; i < s; i++)
                    {              
                        string _REPORTCNIC = items[i];
                        _databasename1 = _REPORTCNIC;
                        
                       // string filePathE = @"E:\SOA\daily_evening\";
                       
                     string filePathE = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\";
                        
                       
                        
                        string check_path = filePathE + _REPORTCNIC + ".pdf";
                        bool isExist_ = File.Exists(check_path);

                        if (isExist_ == true && div_check == true)
                        {
                            int a = 0;
                           // File.Delete(check_path);
                           // isExist_ = false;
                        }
                        
                        if (isExist_ == false)
                        {                            
                            string filePath1 = load_report;
                            // filePath = @"C:\Users\misbah.haque\Desktop\almeezan portal\ACCOUNT_OPENING\AccountOpeningForm.rpt"; // location of rpt file 
                            //filePath = Server.MapPath("~/Report/AccountOpeningForm.rpt");
                            // filePath = Server.MapPath("~/MisReports/Reports/AccountStatement_CRM.rpt");
                            //ReportDocument report = new ReportDocument();
                            //report.Load(filePath);
                           
                            ReportDocument report1 = new ReportDocument();
                            report1.Load(filePath1);
                            MemoryStream oStream;
                            
                            //string _servername = ConfigurationManager.AppSettings["servername"];
                            //string _databasename = ConfigurationManager.AppSettings["databasename"];
                            //string _userid = ConfigurationManager.AppSettings["userid"];
                            //string _password = ConfigurationManager.AppSettings["password"];
                            //string _databasename1 = ConfigurationManager.AppSettings["databasename1"];

                            // _REPORTCNIC = "159951753";
                            report1.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            report1.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            //report1.Subreports[1].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            //report1.Subreports[2].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            report1.RecordSelectionFormula = "({vw_Bulk_AS_Header.Portfolio_ID} = {?@PFNO})";  // get it from crystal report expert option //
                            report1.SetParameterValue("@PFNO", _REPORTCNIC); // BADINDEX EXCEPTION
                            //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
                            report1.SetParameterValue("@PFNO", _REPORTCNIC);
                            report1.SetParameterValue("@StartDate", start_date);
                            report1.SetParameterValue("@EndDate", end_date);
                            //comment
                            report1.SetParameterValue("@PFNO_new", _REPORTCNIC);
                            report1.SetParameterValue("@StartDate_new", start_date);
                            report1.SetParameterValue("@EndDate_new", end_date);

                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

                            //    System.IO.MemoryStream mem = (System.IO.MemoryStream)report1.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                            oStream = (MemoryStream)
                            report1.ExportToStream(
                            CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

                            //HttpContext.Current.Response.Clear();
                            //HttpContext.Current.Response.Buffer = true;
                            //HttpContext.Current.Response.ContentType = "Application/pdf";
                            //HttpContext.Current.Response.BinaryWrite(oStream.ToArray());

                            /////////////////////////////////////////////////////////////////////////////////////////////////

                            //Response.Clear();
                            //Response.Buffer = true;
                            //Response.ContentType = "Application/pdf";
                            //Response.BinaryWrite(oStream.ToArray());
                            // path of the folder where adobe file to be created  
                            //string UserFolder = Session["Username"].ToString();
                            //string path = "";
                            //if (!Directory.Exists(path))
                            //{
                            //    Directory.CreateDirectory(path);
                            //}                            
                            filePathE = filePathE + _REPORTCNIC + ".pdf";
                            //if (!Directory.Exists(filePathE))
                            //{
                            // Directory.CreateDirectory(filePathE);
                            //}
                            bool isExist = File.Exists(filePathE);
                            if (isExist)
                            {
                                continue;
                               //File.Delete(filePathE);
                            }
                            report1.ExportToDisk(ExportFormatType.PortableDocFormat, filePathE);
                            // filePathE = Server.MapPath(path+"/"+a+".pdf");
                            //    filePathE = (path + "\\");  
                            //if (!Directory.Exists(filePathE))
                            //{
                            // Directory.CreateDirectory(filePathE);
                            //}
                            //  report.ExportToDisk(ExportFormatType.PortableDocFormat,filePath);
                            //  Response.End();
                            //oStream.Flush();
                            //oStream.Close();
                            //oStream.Dispose();
                            report1.Close();
                            report1.Dispose();

                         //   bool monthly = true;
                        }
                           // if (items_letter.Contains(items[i]))
                            
                            
                            //if (items_letter.Contains(items[i]))
                            //{

                            //    string _REPORTCNIC_letter = items[i];              //= txt_cnicnum.Text;
                              
                            //  //  string filePathE_letter = @"E:\SOA\daily_evening_letter\";
                                
                            //   string filePathE_letter = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\Letter\";
                                
                                
                            //    string check_path_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
                            //    bool isExist_letter = File.Exists(check_path_letter);
                            //    if (isExist_letter == true) { continue; }

                            //    string filePath1_letter = load_report_letter_;

                            //    ReportDocument report1_letter = new ReportDocument();

                            //    isExist_letter = File.Exists(check_path_letter);
                            //    if (isExist_ == true) { continue; }

                            //    report1_letter.Load(filePath1_letter);
                            //    MemoryStream oStream1;
                            //    report1_letter.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            //    // report1_letter.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);

                            //    isExist_letter = File.Exists(check_path_letter);
                            //    if (isExist_letter == true) { continue; }

                            //    report1_letter.RecordSelectionFormula = "{InvestmentAcknowLetter.Portfolio_ID}={?portfolio_id}  and {InvestmentAcknowLetter.TradeDateTime} >={?@StartDate}";  // get it from crystal report expert option //
                            //    report1_letter.SetParameterValue("portfolio_id", _REPORTCNIC_letter);
                            //    //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
                            //    //  report1_letter.SetParameterValue("@PFNO", _REPORTCNIC);
                            //    report1_letter.SetParameterValue("@StartDate",letter_date);
                            //    //report1_letter.SetParameterValue("@EndDate", end_date);
                            //    ////comment
                            //    //report1_letter.SetParameterValue("@PFNO_new", _REPORTCNIC);
                            //    //report1_letter.SetParameterValue("@StartDate_new", start_date);
                            //    //report1_letter.SetParameterValue("@EndDate_new", end_date);

                            //    isExist_letter = File.Exists(check_path_letter);
                            //    if (isExist_letter == true) { continue; }

                            //    oStream1 = (MemoryStream)
                            //    report1_letter.ExportToStream(
                            //    CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                            //    filePathE_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
                            //    bool isExist_letter_ = File.Exists(filePathE_letter);
                            //    if (isExist_letter_)
                            //    {
                            //        continue;
                            //        // File.Delete(filePathE);
                            //    }
                            //    report1_letter.ExportToDisk(ExportFormatType.PortableDocFormat, filePathE_letter);
                            //    report1_letter.Close();
                            //    report1_letter.Dispose();
                            //      y = y + 1;

                            //}
                        
                        
                       // }
                      //  string filePath1 = @"E:\SOA\reports_soa\AccountStatement_DWH_PF_Portrait_Email_New_Daily.rpt";
                        //cr = Int32.Parse(label3.Text);
                        //cr = cr + 1;
                        //label3.Text = cr.ToString();
                //        Console.Write("\r{0}   ",cr.ToString() + "  generated  out of  "+  s.ToString());
                        //if (IsHandleCreated)
                        //{
                        //    // Always asynchronous, even on the UI thread already.  (Don't let it loop back here!)
                        //    BeginInvoke(new UpdateTextCallback(updateprogress));
                        //    return; // Fired-off asynchronously; let the current thread continue.
                        //    // WriteToForm will be called on the UI thread at some point in the near future.
                        //}
                        //else
                        //{
                        //    // Handle the error case, or do nothing.
                        //}

                        //if (IsHandleCreated)
                        //{
                        //    // Always synchronous.  (But you must watch out for cross-threading deadlocks!)
                        //    if (InvokeRequired)
                        //        Invoke(new UpdateTextCallback(updateprogress));
                        //    else
                        //        updateprogress(); // Call the method (or delegate) directly.
                        //    // Execution continues from here only once WriteToForm has completed and returned.
                        //}
                        //else
                        //{
                        //    // Handle the error case, or do nothing.
                        //}
                    
                        //  public delegate void UpdateTextCallback(string text);

                        //UpdateTextCallback delInstatnce = new UpdateTextCallback(UpdateLabel);
                        //this.Invoke(delInstatnce);

                        //  UpdateLabel(cr.ToString());
                       // Thread.Sleep(100);
                        
                        //cr = Int32.Parse(label1.Text);
                        //cr = cr + 1;
                        ////label1.Invoke((MethodInvoker)(() => label1.Text = cr.ToString()));   
                        ////label1.Text = cr.ToString();
                        //label1.BeginInvoke(new UpdateTextCallback(this.UpdateLabel),
                        //new object[] { cr.ToString() });
                    }

                }
                v = true;
                //watch.Stop();
               // string elapsedMs = watch.Elapsed.Duration().ToString();

                progressBar1.Value = progressBar1.Value+1;

                if (checkBox1.Checked == true)
                {

                    if (progressBar1.Value == 9) { MessageBox.Show("report generation  completed !"); }

                }
                else
                {
                    if (progressBar1.Value == 11) { MessageBox.Show("report generation  completed !"); }
                }


               // MessageBox.Show("report generation part 1 completed !"+ "reports generated " + y.ToString());
                break;                                           // to get out of the error exception loop after successful completion 
            }
            catch (Exception e) 
            {
                MessageBox.Show(" in thread function__1 "+ "__"+  _databasename1+"__" + e.ToString());
            }
            
         }


     }
        public void Thread2(int s1, int s, string[] items, bool v, int cr, string mesg, DateTime start_date, DateTime end_date,DateTime letter_date,string _servername, string _databasename, string _userid, string _password, string _databasename1, int total_count, string load_report, string load_report_letter_,string[]items_letter)
        {
            int b = 0;
            int y = 0;
            for (int h=0;h<20;h++)                    // used to keep the program running in case of exception occurence  
            {   
            try
            {
                //DateTime startdate = new DateTime(start_date.Year, start_date.Month, start_date.Day);
                //DateTime enddate = new DateTime(end_date.Year, end_date.Month, end_date.Day);                   
                 //DateTime startdate1 = dateTimePicker1.Value.Date;
                //DateTime enddate1 = dateTimePicker2.Value.Date;                                  
                if (v==false)
                {
                    for (int i = s1; i < s; i++)
                    {         
                        string _REPORTCNIC = items[i];              //= txt_cnicnum.Text;
                        _databasename1 = _REPORTCNIC;
                        
                       //  string filePathE = @"E:\SOA\daily_evening\";
                        
                        string filePathE = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\";
                        
                        
                        string check_path = filePathE + _REPORTCNIC + ".pdf";
                        bool isExist_ = File.Exists(check_path);


                        if (isExist_ == true && div_check == true)
                        {
                            int a = 0;
                         // File.Delete(check_path);
                          //  isExist_ = false;
                        }


                        
                        if (isExist_==false)
                        {
                            string filePath1 = load_report;
                            //filePath = @"C:\Users\misbah.haque\Desktop\almeezan portal\REPORT_BACKUP\New folder\AccountOpeningForm.rpt"; // location of rpt file 
                            // filePath = @"C:\Users\misbah.haque\Desktop\almeezan portal\ACCOUNT_OPENING\AccountOpeningForm.rpt"; // location of rpt file 
                            //filePath = Server.MapPath("~/Report/AccountOpeningForm.rpt");
                            // filePath = Server.MapPath("~/MisReports/Reports/AccountStatement_CRM.rpt");
                            //ReportDocument report = new ReportDocument();
                            //report.Load(filePath);

                           
                            ReportDocument report1 = new ReportDocument();
                            report1.Load(filePath1);
                            MemoryStream oStream;
                            //string _servername = ConfigurationManager.AppSettings["servername"];
                            //string _databasename = ConfigurationManager.AppSettings["databasename"];
                            //string _userid = ConfigurationManager.AppSettings["userid"];
                            //string _password = ConfigurationManager.AppSettings["password"];
                            //string _databasename1 = ConfigurationManager.AppSettings["databasename1"];
                            // _REPORTCNIC = "159951753";
                            report1.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            report1.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            //report1.Subreports[1].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            //report1.Subreports[2].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                            report1.RecordSelectionFormula = "({vw_Bulk_AS_Header.Portfolio_ID} = {?@PFNO})";  // get it from crystal report expert option //
                            report1.SetParameterValue("@PFNO", _REPORTCNIC);
                            //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
                            report1.SetParameterValue("@PFNO", _REPORTCNIC);
                            report1.SetParameterValue("@StartDate", start_date);
                            report1.SetParameterValue("@EndDate", end_date);
                            //comment
                            report1.SetParameterValue("@PFNO_new", _REPORTCNIC);
                            report1.SetParameterValue("@StartDate_new", start_date);
                            report1.SetParameterValue("@EndDate_new", end_date);
                            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            //    System.IO.MemoryStream mem = (System.IO.MemoryStream)report1.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                           
                            
                            oStream = (MemoryStream)
                            report1.ExportToStream(
                            CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                            //HttpContext.Current.Response.Clear();
                            //HttpContext.Current.Response.Buffer = true;
                            //HttpContext.Current.Response.ContentType = "Application/pdf";
                            //HttpContext.Current.Response.BinaryWrite(oStream.ToArray());

                            /////////////////////////////////////////////////////////////////////////////////////////////////

                            //Response.Clear();
                            //Response.Buffer = true;
                            //Response.ContentType = "Application/pdf";
                            //Response.BinaryWrite(oStream.ToArray());
                            // path of the folder where adobe file to be created  
                            //string UserFolder = Session["Username"].ToString();
                            //string path = "";
                            //if (!Directory.Exists(path))
                            //{
                            //    Directory.CreateDirectory(path);
                            //}
                            filePathE=filePathE + _REPORTCNIC + ".pdf";
                            //if (!Directory.Exists(filePathE))
                            //{
                            // Directory.CreateDirectory(filePathE);
                            //}
                           bool  isExist = File.Exists(filePathE);
                            if (isExist)
                            {
                                continue;
                               //File.Delete(filePathE);
                            }
                            report1.ExportToDisk(ExportFormatType.PortableDocFormat, filePathE);
                            // filePathE = Server.MapPath(path+"/"+a+".pdf");
                            //    filePathE = (path + "\\");
                            //   filePathE = filePathE + a + ".pdf";
                            //if (!Directory.Exists(filePathE))
                            //{
                            // Directory.CreateDirectory(filePathE);
                            //}
                            //  report.ExportToDisk(ExportFormatType.PortableDocFormat,filePath);
                            //  Response.End();
                            //oStream.Flush();
                            //oStream.Close();
                            //oStream.Dispose();
                            report1.Close();
                            report1.Dispose();
                        }

                            // bool monthly = true;
                            
   //                         if(items_letter.Contains(items[i]))
   //                          {

                            

   //                              string _REPORTCNIC_letter = items[i];              //= txt_cnicnum.Text;
                               
                                
   //                           //   string filePathE_letter = @"E:\SOA\daily_evening_letter\";
                                
   //                              string filePathE_letter = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\Letter\";
                                
                                
   //                              string check_path_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
   //                              bool isExist_letter = File.Exists(check_path_letter);
   //                              if (isExist_letter == true) { continue; }

   //                              string filePath1_letter = load_report_letter_;

   //                              ReportDocument report1_letter = new ReportDocument();

   //                              isExist_ = File.Exists(check_path_letter);
   //                              if (isExist_ == true) { continue; }

   //                              report1_letter.Load(filePath1_letter);
   //                              MemoryStream oStream1;
   //                              report1_letter.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
   //                             // report1_letter.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);

   //                              isExist_letter = File.Exists(check_path_letter);
   //                              if (isExist_letter == true) { continue; }

   //                              report1_letter.RecordSelectionFormula = "{InvestmentAcknowLetter.Portfolio_ID}={?portfolio_id}  and {InvestmentAcknowLetter.TradeDateTime} >={?@StartDate}";  // get it from crystal report expert option //
                                
                                  
   ////ssqlrpt = "({InvestmentAcknowLetter.Portfolio_ID} ='" & PorfolioID & "'"
   ////ssqlrpt = ssqlrpt & ") and ({InvestmentAcknowLetter.TradeDateTime} >= date(" & Year(StartDate_Letter) & "," & Month(StartDate_Letter) & "," & Day(StartDate_Letter) & ") and  {InvestmentAcknowLetter.TradeDateTime} <= date(" & Year(.txtEndDate) & "," & Month(.txtEndDate) & "," & Day(.txtEndDate) & "))"
   
   //                             report1_letter.SetParameterValue("portfolio_id", _REPORTCNIC_letter);
   //                              //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
   //                            //  report1_letter.SetParameterValue("@PFNO", _REPORTCNIC);
   //                              report1_letter.SetParameterValue("@StartDate",letter_date);
   //                              //report1_letter.SetParameterValue("@EndDate", end_date);
   //                              ////comment
   //                              //report1_letter.SetParameterValue("@PFNO_new", _REPORTCNIC);
   //                              //report1_letter.SetParameterValue("@StartDate_new", start_date);
   //                              //report1_letter.SetParameterValue("@EndDate_new", end_date);

   //                              isExist_letter = File.Exists(check_path_letter);
   //                              if (isExist_letter == true) { continue; }

   //                         oStream1 = (MemoryStream)
   //                         report1_letter.ExportToStream(
   //                         CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                                
   //                             filePathE_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
   //                              bool isExist_letter_ = File.Exists(filePathE_letter);
   //                              if (isExist_letter_)
   //                              {
   //                                  continue;
   //                                  // File.Delete(filePathE);
   //                              }
   //                              report1_letter.ExportToDisk(ExportFormatType.PortableDocFormat,filePathE_letter);
   //                              report1_letter.Close();
   //                              report1_letter.Dispose();
   //                          //    y = y + 1;

   //                              b = b + 1;

                            
   //                         }   //
                        
                        
                   
                    }
             //   }   //////////

                    
                    //for (int f = items.Count()-1;f >=0;f--)
                    //{

                    //    string _REPORTCNIC = items[f];              //= txt_cnicnum.Text;
                    //    string filePathE = @"E:\SOA\daily_evening\";
                    //    string check_path = filePathE + _REPORTCNIC + ".pdf";
                    //    bool isExist_ = File.Exists(check_path);
                    //    if (isExist_ == false)
                    //    {
                    //        string filePath1 = load_report;
                    //        //filePath = @"C:\Users\misbah.haque\Desktop\almeezan portal\REPORT_BACKUP\New folder\AccountOpeningForm.rpt"; // location of rpt file 
                    //        // filePath = @"C:\Users\misbah.haque\Desktop\almeezan portal\ACCOUNT_OPENING\AccountOpeningForm.rpt"; // location of rpt file 
                    //        //filePath = Server.MapPath("~/Report/AccountOpeningForm.rpt");
                    //        // filePath = Server.MapPath("~/MisReports/Reports/AccountStatement_CRM.rpt");
                    //        isExist_ = File.Exists(check_path);
                    //        if (isExist_ == true) { continue; }


                    //        ReportDocument report1 = new ReportDocument();
                    //        report1.Load(filePath1);
                    //        MemoryStream oStream;
                    //        //string _servername = ConfigurationManager.AppSettings["servername"];
                    //        //string _databasename = ConfigurationManager.AppSettings["databasename"];
                    //        //string _userid = ConfigurationManager.AppSettings["userid"];
                    //        //string _password = ConfigurationManager.AppSettings["password"];
                    //        //string _databasename1 = ConfigurationManager.AppSettings["databasename1"];
                    //        // _REPORTCNIC = "159951753";

                    //        isExist_ = File.Exists(check_path);
                    //        if (isExist_ == true) { continue; }

                    //        report1.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                    //        report1.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                    //        //report1.Subreports[1].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                    //        //report1.Subreports[2].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                    //        report1.RecordSelectionFormula = "({vw_Bulk_AS_Header.Portfolio_ID} = {?@PFNO})";  // get it from crystal report expert option //
                    //        report1.SetParameterValue("@PFNO", _REPORTCNIC);
                    //        //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
                    //        report1.SetParameterValue("@PFNO", _REPORTCNIC);
                    //        report1.SetParameterValue("@StartDate", start_date);
                    //        report1.SetParameterValue("@EndDate", end_date);
                    //        //comment
                    //        report1.SetParameterValue("@PFNO_new", _REPORTCNIC);
                    //        report1.SetParameterValue("@StartDate_new", start_date);
                    //        report1.SetParameterValue("@EndDate_new", end_date);
                    //        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    //        //    System.IO.MemoryStream mem = (System.IO.MemoryStream)report1.ExportToStream(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);                                                       
                    //        isExist_ = File.Exists(check_path);
                    //        if (isExist_ == true) { continue; }

                    //        oStream = (MemoryStream)
                    //        report1.ExportToStream(
                    //        CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);
                   
                            

                    //        /////////////////////////////////////////////////////////////////////////////////////////////////

                    //        //Response.Clear();
                    //        //Response.Buffer = true;
                    //        //Response.ContentType = "Application/pdf";
                    //        //Response.BinaryWrite(oStream.ToArray());
                    //        // path of the folder where adobe file to be created  
                    //        //string UserFolder = Session["Username"].ToString();                           
                    //        //if (!Directory.Exists(path))
                    //        //{
                    //        //    Directory.CreateDirectory(path);
                    //        //}
                    //        filePathE = filePathE + _REPORTCNIC + ".pdf";
                    //        //if (!Directory.Exists(filePathE))
                    //        //{
                    //        // Directory.CreateDirectory(filePathE);
                    //        //}
                    //        bool isExist = File.Exists(filePathE);
                    //        if (isExist)
                    //        {
                    //            continue;
                    //            //File.Delete(filePathE);
                    //        }
                    //        report1.ExportToDisk(ExportFormatType.PortableDocFormat, filePathE);
                    //        // filePathE = Server.MapPath(path+"/"+a+".pdf");                            
                    //        //if (!Directory.Exists(filePathE))
                    //        //{
                    //        // Directory.CreateDirectory(filePathE);
                    //        //}
                    //        //  report.ExportToDisk(ExportFormatType.PortableDocFormat,filePath);
                    //        //  Response.End();
                    //        //oStream.Flush();
                    //        //oStream.Close();
                    //        //oStream.Dispose();
                    //        report1.Close();
                    //        report1.Dispose();

                    //    }


                    //}   //

                    //for (int u = items.Count()-1; u>=0; u--)
                    //{

                    //    if (items_letter.Contains(items[u]))
                    //    {

                    //        string _REPORTCNIC_letter = items[u];              //= txt_cnicnum.Text;
                    //        string filePathE_letter = @"E:\SOA\daily_evening_letter\";
                    //        string check_path_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
                    //        bool isExist_letter = File.Exists(check_path_letter);
                    //        if (isExist_letter == true) 
                    //        { 
                    //            continue; 
                    //        }

                    //        string filePath1_letter = load_report_letter_;

                    //        ReportDocument report1_letter = new ReportDocument();

                    //        isExist_letter = File.Exists(check_path_letter);
                    //        if (isExist_letter == true) 
                    //        { 
                    //            continue; 
                    //        }

                    //        report1_letter.Load(filePath1_letter);
                    //        MemoryStream oStream1;
                    //        report1_letter.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                    //        // report1_letter.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);

                    //        isExist_letter = File.Exists(check_path_letter);
                    //        if (isExist_letter == true) { continue; }

                    //        report1_letter.RecordSelectionFormula = "{InvestmentAcknowLetter.Portfolio_ID}={?portfolio_id}  and {InvestmentAcknowLetter.TradeDateTime} >={?@StartDate}";  // get it from crystal report expert option //
                    //        report1_letter.SetParameterValue("portfolio_id", _REPORTCNIC_letter);
                    //        //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
                    //        //  report1_letter.SetParameterValue("@PFNO", _REPORTCNIC);
                    //        report1_letter.SetParameterValue("@StartDate", letter_date);
                    //        //report1_letter.SetParameterValue("@StartDate", start_date);
                    //        //report1_letter.SetParameterValue("@EndDate", end_date);
                    //        ////comment
                    //        //report1_letter.SetParameterValue("@PFNO_new", _REPORTCNIC);
                    //        //report1_letter.SetParameterValue("@StartDate_new", start_date);
                    //        //report1_letter.SetParameterValue("@EndDate_new", end_date);

                    //        isExist_letter = File.Exists(check_path_letter);
                    //        if (isExist_letter == true) 
                    //        { 
                    //            continue; 
                    //        }

                    //        oStream1 = (MemoryStream)
                    //        report1_letter.ExportToStream(
                    //        CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

                    //        filePathE_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
                    //        bool isExist_letter_ = File.Exists(filePathE_letter);
                    //        if (isExist_letter_)
                    //        {
                    //            continue;
                    //            // File.Delete(filePathE);
                    //        }
                    //        report1_letter.ExportToDisk(ExportFormatType.PortableDocFormat, filePathE_letter);
                    //        report1_letter.Close();
                    //        report1_letter.Dispose();
                    //        y = y + 1;

                    //    }
                    //}     //
                                    
                
                }

                    v = true;
                //watch.Stop();
                ////watch.Elapsed.Duration().ToString();
                //string elapsedMs = watch.Elapsed.Duration().ToString();
                    progressBar1.Value = progressBar1.Value + 1;

                    if (checkBox1.Checked == true)
                    {

                        if (progressBar1.Value == 9) { MessageBox.Show("report generation  completed !"); }

                    }
                    else
                    {
                        if (progressBar1.Value == 11) { MessageBox.Show("report generation  completed !"); }
                    }
              
              //  MessageBox.Show("report generation part" + mesg + "completed !" + "letter generated in sub loop " + y.ToString() + "letter generated in main loop"+b.ToString());
                break;           
            }
            catch (Exception e)
            {
                MessageBox.Show(" in thread function__" + mesg+  "__" + _databasename1+ "__"  + e.ToString()); 
            }
          
         }

      }
        public void Thread3(int s1, int s, string[] items, bool v, int cr, string mesg, DateTime start_date, DateTime end_date, DateTime letter_date, string _servername, string _databasename, string _userid, string _password, string _databasename1, int total_count, string load_report, string load_report_letter_, string[] items_letter)
        {
            int b = 0;
            int y = 0;
            for (int h = 0; h < 20; h++)                    // used to keep the program running in case of exception occurence  
            {
                try
                {
                    //DateTime startdate = new DateTime(start_date.Year, start_date.Month, start_date.Day);
                    //DateTime enddate = new DateTime(end_date.Year, end_date.Month, end_date.Day);                   
                    //DateTime startdate1 = dateTimePicker1.Value.Date;
                    //DateTime enddate1 = dateTimePicker2.Value.Date;                                  
                    if (v == false)
                    {
                        for (int i = s1; i < s; i++)
                        {
                            string _REPORTCNIC = items[i];              //= txt_cnicnum.Text;
                            _databasename1 = _REPORTCNIC;

                            // string filePathE = @"E:\SOA\daily_evening\";

                           string filePathE = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\";


                            string check_path = filePathE + _REPORTCNIC + ".pdf";
                            bool isExist_ = File.Exists(check_path);


                            if (isExist_ == true && div_check == true)
                            {
                                int a = 0;
                            //   File.Delete(check_path);
                                isExist_ = false;
                            }

                            // bool monthly = true;

                            if (items_letter.Contains(items[i]))
                            {



                                string _REPORTCNIC_letter = items[i];              //= txt_cnicnum.Text;


                                 //  string filePathE_letter = @"E:\SOA\daily_evening_letter\";

                              string filePathE_letter = @"X:\T24 BulkPrinting\Fetch Reports\Exported\Bulk\Letter\";


                                string check_path_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
                                bool isExist_letter = File.Exists(check_path_letter);
                                if (isExist_letter == true) { continue; }

                                string filePath1_letter = load_report_letter_;

                                ReportDocument report1_letter = new ReportDocument();

                                isExist_ = File.Exists(check_path_letter);
                                if (isExist_ == true) { continue; }

                                report1_letter.Load(filePath1_letter);
                                MemoryStream oStream1;
                                report1_letter.DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);
                                // report1_letter.Subreports[0].DataSourceConnections[0].SetConnection(_servername, _databasename, _userid, _password);

                                isExist_letter = File.Exists(check_path_letter);
                                if (isExist_letter == true) { continue; }

                                report1_letter.RecordSelectionFormula = "{InvestmentAcknowLetter.Portfolio_ID}={?portfolio_id}  and {InvestmentAcknowLetter.TradeDateTime} >={?@StartDate}";  // get it from crystal report expert option //


                                //ssqlrpt = "({InvestmentAcknowLetter.Portfolio_ID} ='" & PorfolioID & "'"
                                //ssqlrpt = ssqlrpt & ") and ({InvestmentAcknowLetter.TradeDateTime} >= date(" & Year(StartDate_Letter) & "," & Month(StartDate_Letter) & "," & Day(StartDate_Letter) & ") and  {InvestmentAcknowLetter.TradeDateTime} <= date(" & Year(.txtEndDate) & "," & Month(.txtEndDate) & "," & Day(.txtEndDate) & "))"

                                report1_letter.SetParameterValue("portfolio_id", _REPORTCNIC_letter);
                                //  report.SetParameterValue("@StartDate", DateTime.Now.AddYears(-1).ToShortDateString());
                                //  report1_letter.SetParameterValue("@PFNO", _REPORTCNIC);
                                report1_letter.SetParameterValue("@StartDate", letter_date);
                                //report1_letter.SetParameterValue("@EndDate", end_date);
                                ////comment
                                //report1_letter.SetParameterValue("@PFNO_new", _REPORTCNIC);
                                //report1_letter.SetParameterValue("@StartDate_new", start_date);
                                //report1_letter.SetParameterValue("@EndDate_new", end_date);

                                isExist_letter = File.Exists(check_path_letter);
                                if (isExist_letter == true) { continue; }

                                oStream1 = (MemoryStream)
                                report1_letter.ExportToStream(
                                CrystalDecisions.Shared.ExportFormatType.PortableDocFormat);

                                filePathE_letter = filePathE_letter + _REPORTCNIC_letter + "Letter.pdf";
                                bool isExist_letter_ = File.Exists(filePathE_letter);
                                if (isExist_letter_)
                                {
                                    continue;
                                    // File.Delete(filePathE);
                                }
                                report1_letter.ExportToDisk(ExportFormatType.PortableDocFormat, filePathE_letter);
                                report1_letter.Close();
                                report1_letter.Dispose();
                                //    y = y + 1;

                                b = b + 1;


                            }   //
                        
                        }
                      
                    }

                    v = true;
                    //watch.Stop();
                    ////watch.Elapsed.Duration().ToString();
                    //string elapsedMs = watch.Elapsed.Duration().ToString();
                    progressBar1.Value = progressBar1.Value + 1;
                    if (progressBar1.Value == 11) { MessageBox.Show("report generation  completed !" + mesg); }


                    //  MessageBox.Show("report generation part" + mesg + "completed !" + "letter generated in sub loop " + y.ToString() + "letter generated in main loop"+b.ToString());
                    break;
                }
                catch (Exception e)
                {
                    MessageBox.Show(" in thread function__" + mesg + "__" + _databasename1 + "__" + e.ToString());
                }

            }

        }



   }

    public class clsDBOperation5
    {
        private string _strConnectionString;
        public clsDBOperation5()
        {
            // _strConnectionString = ConfigurationManager.ConnectionStrings["ConnStringDb1"].ConnectionString;
            _strConnectionString = ConfigurationManager.ConnectionStrings["ConnStringDb1_ISMS_Live"].ConnectionString;
        }

        public SqlConnection _openConnection()
        {
            SqlConnection oCnn = new SqlConnection(_strConnectionString);
            oCnn.Close();
            oCnn.Open();

            return oCnn;
        }

        public SqlConnection _openConnection(string connection_amimnav)
        {
            SqlConnection oCnn = new SqlConnection(connection_amimnav);
            oCnn.Close();
            oCnn.Open();

            return oCnn;
        }

        public void _closeConnection(SqlConnection oCnn)
        {
            oCnn.Close();
            oCnn = null;
        }

        public Boolean ExecuteNonQuery(string strQuery)
        {
            SqlConnection oCnn = _openConnection();
            SqlCommand oCmd = new SqlCommand(strQuery, oCnn);
            oCmd.ExecuteNonQuery();
            oCnn.Close();
            //  _closeConnection(oCnn);
            return true;
        }

        public DataTable GetDataTable(string strQuery)
        {
            SqlConnection oCnn = _openConnection();
            SqlDataAdapter oDA = new SqlDataAdapter(strQuery, oCnn);
            oDA.SelectCommand.CommandType = CommandType.Text;
            oDA.SelectCommand.CommandText = strQuery;

            DataSet oDS = new DataSet();


            oDA.Fill(oDS);
            oCnn.Close();
            //_closeConnection(oCnn);
            return oDS.Tables[0];
        }

        public DataTable GetDataTable(string strQuery, string connect_amimnav)
        {
            SqlConnection oCnn = _openConnection(connect_amimnav);
            SqlDataAdapter oDA = new SqlDataAdapter(strQuery, oCnn);
            oDA.SelectCommand.CommandType = CommandType.Text;
            oDA.SelectCommand.CommandText = strQuery;

            DataSet oDS = new DataSet();


            oDA.Fill(oDS);
            oCnn.Close();
            //_closeConnection(oCnn);
            return oDS.Tables[0];
        }



        public bool purchase_record(string start_date, string end_date)
        {
            try
            {
                SqlConnection oCon = _openConnection();
                SqlCommand cmd = new SqlCommand("dbo.sp_soa_daily_purchase", oCon);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@From", "Al Meezan Investments<newsletters@almeezangroup.com>");
                cmd.Parameters.AddWithValue("@StartDate", start_date);
                cmd.Parameters.AddWithValue("@EndDate", end_date);
                //cmd.Parameters.AddWithValue("@Attachment", attachment);
                //cmd.Parameters.AddWithValue("@ImageContentID", "cid:fb");
                //cmd.Parameters.AddWithValue("@EmbeddedImage", "1");
                //cmd.Parameters.AddWithValue("@EntryDate", DateTime.Now.ToString("yyyy-MM-dd"));
                //string queryEmailPool = "INSERT INTO [172.16.1.42].[ISMS].[dbo].[EmailPool] (fromemail,toemail,subject,body,attachment,entrydate,flag,EmbededImage) VALUES (@fromEmail,@toEmail,@Subj,@Body,@Attach,@EntryDate,@Flag,@EmbedImg)";
                //SqlCommand cmdMsg = new SqlCommand(queryEmailPool, oCon);
                //cmdMsg = new SqlCommand(queryEmailPool, Connection.GetConnection());
                //cmdMsg.Parameters.AddWithValue("@fromEmail", "Al Meezan Investments <newsletters@almeezangroup.com>");
                //cmdMsg.Parameters.AddWithValue("@toEmail", email);

                //cmdMsg.Parameters.AddWithValue("@Body", msg);
                //cmdMsg.Parameters.AddWithValue("@Attach", "none");
                //cmdMsg.Parameters.AddWithValue("@EntryDate", DateTime.Now);
                //cmdMsg.Parameters.AddWithValue("@EmbedImg", 0);
                //cmdMsg.Parameters.AddWithValue("@Flag", 'N');
                //cmdMsg.ExecuteNonQuery();

                cmd.CommandTimeout = 300;

                int a = cmd.ExecuteNonQuery();
                oCon.Close();
                _closeConnection(oCon);

                //return true;

                if (a == 0)
                {
                    return false;     //not updated
                }
                else
                {
                    return true;//updated
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;

            }
        }

        public bool non_purchase_record(string start_date, string end_date)
        {
            try
            {
                SqlConnection oCon = _openConnection();
                SqlCommand cmd = new SqlCommand("dbo.sp_soa_daily_non_purchase", oCon);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@From", "Al Meezan Investments<newsletters@almeezangroup.com>");
                cmd.Parameters.AddWithValue("@StartDate", start_date);
                cmd.Parameters.AddWithValue("@EndDate", end_date);
                //cmd.Parameters.AddWithValue("@Attachment", attachment);
                //cmd.Parameters.AddWithValue("@ImageContentID", "cid:fb");
                //cmd.Parameters.AddWithValue("@EmbeddedImage", "1");
                //cmd.Parameters.AddWithValue("@EntryDate", DateTime.Now.ToString("yyyy-MM-dd"));
                //string queryEmailPool = "INSERT INTO [172.16.1.42].[ISMS].[dbo].[EmailPool] (fromemail,toemail,subject,body,attachment,entrydate,flag,EmbededImage) VALUES (@fromEmail,@toEmail,@Subj,@Body,@Attach,@EntryDate,@Flag,@EmbedImg)";
                //SqlCommand cmdMsg = new SqlCommand(queryEmailPool, oCon);
                //cmdMsg = new SqlCommand(queryEmailPool, Connection.GetConnection());
                //cmdMsg.Parameters.AddWithValue("@fromEmail", "Al Meezan Investments <newsletters@almeezangroup.com>");
                //cmdMsg.Parameters.AddWithValue("@toEmail", email);

                //cmdMsg.Parameters.AddWithValue("@Body", msg);
                //cmdMsg.Parameters.AddWithValue("@Attach", "none");
                //cmdMsg.Parameters.AddWithValue("@EntryDate", DateTime.Now);
                //cmdMsg.Parameters.AddWithValue("@EmbedImg", 0);
                //cmdMsg.Parameters.AddWithValue("@Flag", 'N');
                //cmdMsg.ExecuteNonQuery();

                cmd.CommandTimeout = 300;

                int a = cmd.ExecuteNonQuery();

                oCon.Close();
                _closeConnection(oCon);
                //  return true;

                if (a == 0)
                {
                    return false;     //not updated
                }
                else
                {
                    return true;//updated
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
        }

        public bool non_purchase_record_Mraf(string start_date, string end_date)
        {
            try
            {
                SqlConnection oCon = _openConnection();
                SqlCommand cmd = new SqlCommand("dbo.sp_soa_daily_MRAF", oCon);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@From", "Al Meezan Investments<newsletters@almeezangroup.com>");
                cmd.Parameters.AddWithValue("@StartDate", start_date);
                cmd.Parameters.AddWithValue("@EndDate", end_date);
                //cmd.Parameters.AddWithValue("@Attachment", attachment);
                //cmd.Parameters.AddWithValue("@ImageContentID", "cid:fb");
                //cmd.Parameters.AddWithValue("@EmbeddedImage", "1");
                //cmd.Parameters.AddWithValue("@EntryDate", DateTime.Now.ToString("yyyy-MM-dd"));
                //string queryEmailPool = "INSERT INTO [172.16.1.42].[ISMS].[dbo].[EmailPool] (fromemail,toemail,subject,body,attachment,entrydate,flag,EmbededImage) VALUES (@fromEmail,@toEmail,@Subj,@Body,@Attach,@EntryDate,@Flag,@EmbedImg)";
                //SqlCommand cmdMsg = new SqlCommand(queryEmailPool, oCon);
                //cmdMsg = new SqlCommand(queryEmailPool, Connection.GetConnection());
                //cmdMsg.Parameters.AddWithValue("@fromEmail", "Al Meezan Investments <newsletters@almeezangroup.com>");
                //cmdMsg.Parameters.AddWithValue("@toEmail", email);

                //cmdMsg.Parameters.AddWithValue("@Body", msg);
                //cmdMsg.Parameters.AddWithValue("@Attach", "none");
                //cmdMsg.Parameters.AddWithValue("@EntryDate", DateTime.Now);
                //cmdMsg.Parameters.AddWithValue("@EmbedImg", 0);
                //cmdMsg.Parameters.AddWithValue("@Flag", 'N');
                //cmdMsg.ExecuteNonQuery();

                cmd.CommandTimeout = 300;

                int a = cmd.ExecuteNonQuery();

                oCon.Close();
                _closeConnection(oCon);
                //  return true;

                if (a == 0)
                {
                    return false;     //not updated
                }
                else
                {
                    return true;//updated
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
        }

        public bool purchase_record_Mraf(string start_date, string end_date)
        {
            try
            {
                SqlConnection oCon = _openConnection();
                SqlCommand cmd = new SqlCommand("dbo.sp_soa_daily_MRAF_Purchase", oCon);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@From", "Al Meezan Investments<newsletters@almeezangroup.com>");
                cmd.Parameters.AddWithValue("@StartDate", start_date);
                cmd.Parameters.AddWithValue("@EndDate", end_date);
                //cmd.Parameters.AddWithValue("@Attachment", attachment);
                //cmd.Parameters.AddWithValue("@ImageContentID", "cid:fb");
                //cmd.Parameters.AddWithValue("@EmbeddedImage", "1");
                //cmd.Parameters.AddWithValue("@EntryDate", DateTime.Now.ToString("yyyy-MM-dd"));
                //string queryEmailPool = "INSERT INTO [172.16.1.42].[ISMS].[dbo].[EmailPool] (fromemail,toemail,subject,body,attachment,entrydate,flag,EmbededImage) VALUES (@fromEmail,@toEmail,@Subj,@Body,@Attach,@EntryDate,@Flag,@EmbedImg)";
                //SqlCommand cmdMsg = new SqlCommand(queryEmailPool, oCon);
                //cmdMsg = new SqlCommand(queryEmailPool, Connection.GetConnection());
                //cmdMsg.Parameters.AddWithValue("@fromEmail", "Al Meezan Investments <newsletters@almeezangroup.com>");
                //cmdMsg.Parameters.AddWithValue("@toEmail", email);

                //cmdMsg.Parameters.AddWithValue("@Body", msg);
                //cmdMsg.Parameters.AddWithValue("@Attach", "none");
                //cmdMsg.Parameters.AddWithValue("@EntryDate", DateTime.Now);
                //cmdMsg.Parameters.AddWithValue("@EmbedImg", 0);
                //cmdMsg.Parameters.AddWithValue("@Flag", 'N');
                //cmdMsg.ExecuteNonQuery();

                cmd.CommandTimeout = 300;

                int a = cmd.ExecuteNonQuery();

                oCon.Close();
                _closeConnection(oCon);
                //  return true;

                if (a == 0)
                {
                    return false;     //not updated
                }
                else
                {
                    return true;//updated
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return false;
            }
        }


        public DataTable GetDataTableFromExcel(string sMyPath)
        {

            string ConnectionString = "";
            ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + sMyPath + ";Extended Properties='Excel 12.0 xml;HDR=NO;'";
            OleDbConnection cn = new OleDbConnection(ConnectionString);
            cn.Open();
            DataTable dbSchema = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dbSchema == null || dbSchema.Rows.Count < 1)
            {
                throw new Exception("Error: Could not determine the name of the first worksheet.");
            }
            string WorkSheetName = "";
            foreach (DataRow drsh in dbSchema.Rows)
            {
                WorkSheetName = drsh["TABLE_NAME"].ToString();
                string[] arr = WorkSheetName.Split('$');
                if (arr.Length == 2 && arr[1].Length == 0)
                {
                    break;
                }
            }
            OleDbCommand excelCmd = new OleDbCommand();
            excelCmd = cn.CreateCommand();
            excelCmd.CommandText = "SELECT * FROM [" + WorkSheetName + "]";
            excelCmd.CommandType = CommandType.Text;
            // OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM [" + WorkSheetName + "]", cn);
            OleDbDataReader dr = excelCmd.ExecuteReader();
            dr.Read();
            DataTable dt = new DataTable(WorkSheetName);
            dt.Load(dr);

            excelCmd.Dispose();
            dr.Close();
            dr.Dispose();

            cn.Close();
            cn.Dispose();
            //da.Fill(dt);
            return dt;
        }


        public void fetch_trades_details(string start_date, string Query_date, string end_date)
        {

            try
            {
                // string trade_details = "INSERT INTO soa_daily_trades ";
                string trade_details = "";

                trade_details = trade_details + "EXEC SP_FetchCustomerTrade_Insert_Soa'" + start_date + "','" + Query_date + "','" + end_date + "'";

                string connetionString = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn = new SqlConnection(connetionString);
                cnn.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd = new SqlCommand(trade_details, cnn);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                cmd.CommandTimeout = 1500;
                SqlDataReader drCN = cmd.ExecuteReader();
                DataTable dtCN = new DataTable();
                dtCN.Load(drCN);
                int s = dtCN.Rows.Count;
                cnn.Close();



            }
            catch (Exception ex)
            {
                ex.ToString();
                MessageBox.Show(ex.ToString());

            }


        }

        public void remove_trade_details(string del_query)
        {



            try
            {

                // string query = "Delete from soa_daily_trades";
                string connetionString = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn = new SqlConnection(connetionString);
                cnn.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd = new SqlCommand(del_query, cnn);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                SqlDataReader drCN = cmd.ExecuteReader();
                DataTable dtCN = new DataTable();
                dtCN.Load(drCN);
                int s = dtCN.Rows.Count;
                cnn.Close();


            }
            catch (Exception ex)
            {
                ex.ToString();
                MessageBox.Show(ex.ToString());

            }




        }


        public void insert_customer_sp(string sp_query)
        {

            try
            {
                // string trade_details = "INSERT INTO soa_daily_trades ";
                string trade_details = "";

    trade_details 
     = trade_details + "insert into SP_Daily select * from customer_sp where portfolioid in ("+sp_query+")";

                string connetionString = ConfigurationManager.ConnectionStrings["acc_open_con_str"].ConnectionString;
                SqlConnection cnn;
                // connetionString = "Data Source=ServerName;Initial Catalog=DatabaseName;User ID=UserName;Password=Password";
                cnn = new SqlConnection(connetionString);
                cnn.Open();
                // MessageBox.Show ("Connection Open ! ");
                // cnn.Close();
                SqlCommand cmd = new SqlCommand(trade_details, cnn);    // to be corrected 
                // SqlCommand cmd1 = new SqlCommand("select top 1 * from Customer_IDs where ID_NO ='" + txtCnicNo + "'", oCnn);   
                cmd.CommandTimeout = 1500;
                SqlDataReader drCN = cmd.ExecuteReader();
                DataTable dtCN = new DataTable();
                dtCN.Load(drCN);
                int s = dtCN.Rows.Count;
                cnn.Close();
            }
            catch (Exception ex)
            {
                ex.ToString();
                MessageBox.Show(ex.ToString());

            }


        }





    }







}
