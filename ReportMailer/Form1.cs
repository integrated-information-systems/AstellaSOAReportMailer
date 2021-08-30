using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using Dapper;
using Dapper.Contrib.Extensions;
using ReportMailer.Models;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ReportMailer
{


    public partial class Form1 : Form
    {
        static public int CurrentCustomerIndex = 0;
        static public int TotalCustomerCount = 0;
        static List<OCRD> CustomersList = null;
        static public ProcessSelector CurrentProcess { get; set; }

        public static string SenderEmailId { get; } = ConfigurationManager.AppSettings["SenderEmailId"];
        public static string SenderEmailIdPwd { get; } = ConfigurationManager.AppSettings["SenderEmailIdPwd"];
        public static string CCEmailId { get; } = ConfigurationManager.AppSettings["CCEmailId"];
        public static string BCCEmailId { get; } = ConfigurationManager.AppSettings["BCCEmailId"];
        public static string ReportPrefix { get; } = ConfigurationManager.AppSettings["ReportPrefix"];
        public string path = ConfigurationManager.AppSettings["RptFilePath"] + "SyncError.txt";
        public Form1()
        {
            InitializeComponent();
            ClearSyncLog();

        }
        private void ClearSyncLog()
        {
            try
            {
                // Clear the error log file at initail stage
                if (File.Exists(path))
                {
                    File.WriteAllText(path, string.Empty);
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }
        }
        private void Form1_Load(object sender, EventArgs e)
        {


            try
            {
                // var idList = new[] { "C00194", "C00577", "C00552", "C00549", "C00451", "C00388" };
                using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["SAP_DB_ConnectionString"].ConnectionString))
                {
                    //Form1.CustomersList = db.GetAll<OCRD>().Where(x => x.U_send == "Y" && x.CardType == "C" && idList.Contains(x.CardCode)).ToList();
                    Form1.CustomersList = db.GetAll<OCRD>().Where(x => x.U_send == "Y" && x.CardType == "C").ToList();
                }
                WriteLog("Total records found : " + CustomersList.Count);
                TotalCustomerCount = CustomersList.Count - 1;

                // Switch_Process();

            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
                Application.Exit();
            }
        }

        private void Process_Switcher_Tick(object sender, EventArgs e)
        {
            Process_Switcher.Enabled = false;
            Switch_Process();

        }

        private void Switch_Process()
        {
            try
            {
                if (CurrentCustomerIndex == CustomersList.Count)
                {
                    CurrentProcess++;
                    CurrentCustomerIndex = 0;
                }


                switch (CurrentProcess)
                {
                    case ProcessSelector.Report1:

                        using (IDbConnection db = new SqlConnection(ConfigurationManager.ConnectionStrings["SAP_DB_ConnectionString"].ConnectionString))
                        {

                            DateTime toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                            toDate = toDate.AddSeconds(-1); //
                                                            //string toDate = "2021-06-30";

                            string CardCode = CustomersList[CurrentCustomerIndex].CardCode;
                            var procedure = "[IISsp_SOA]";
                            var values = new { @CardCodeFr = CardCode, @CardCodeTo = CardCode, @AgeDate = toDate };
                            var results = db.Query(procedure, values, commandType: CommandType.StoredProcedure).ToList();

                            WriteLog("Customer : " + CardCode);

                            if (results.Count() > 0)
                            {

                                if (CustomersList[CurrentCustomerIndex].E_Mail != null && CustomersList[CurrentCustomerIndex].E_Mail != string.Empty)
                                {

                                    ReportDetails reportDetails = new ReportDetails();
                                    reportDetails.CardCode = CustomersList[CurrentCustomerIndex].CardCode;
                                    reportDetails.CustomerEmail = CustomersList[CurrentCustomerIndex].E_Mail;

                                    WriteLog("report generation " + reportDetails.CardCode + " " + reportDetails.CustomerEmail);

                                    if (CardCode.Contains("/"))
                                    {
                                        CardCode = CardCode.Replace("/", "-");
                                    }
                                    reportDetails.ReportName = ReportPrefix + CardCode;
                                    Generate_Report(reportDetails);
                                }
                            }

                        }

                        break;
                    default:

                        Application.Exit();
                        break;
                }
                CurrentCustomerIndex++;
                Process_Switcher.Enabled = true;

            }
            catch (Exception ex)
            {
                WriteLog("Switch Process : " + ex.Message);
            }
        }

        private void SendReportMail(ReportDetails reportDetails)
        {
            string sourceDirectory = ConfigurationManager.AppSettings["RptFilePath"];
            string filePath = sourceDirectory + reportDetails.ReportFilename;

            DateTime toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
            toDate = toDate.AddSeconds(-1); //

            string lastmonthDate = toDate.ToString("dd MMMM yyyy");
            MailMessage mailMessage = new MailMessage
            {
                Subject = "Your Statement of Account from Astellas Pharma Singapore Pte Ltd is ready: " + reportDetails.ReportName,
                From = new MailAddress(SenderEmailId),
                IsBodyHtml = true,
            };
            try
            {
                WriteLog("sending Customer Mail:");

                StringBuilder mailBody = new StringBuilder();
                mailBody.Append("Dear Customer,<br/></br/>");
                mailBody.Append($"<p>Please find attached Electronic Statement of Account as of {lastmonthDate}.</p>" +
                    $"          <p>This statement is for your reconciliation and payment process.</p>" +
                    $"          <p>If you have any questions regarding the statement please contact us at " +
                    $"<a href='mailto: Finance_AP@sg.astellas.com'>Finance_AP@sg.astellas.com</a></p><br/>");

                mailBody.Append("Kind regards<br/> Astellas Pharma Singapore Pte Ltd");

                mailMessage.Body = mailBody.ToString();

                mailMessage.To.Add(reportDetails.CustomerEmail);

                if (CCEmailId != string.Empty)
                {
                    mailMessage.CC.Add(CCEmailId);
                }
                if (BCCEmailId != string.Empty)
                {
                    mailMessage.Bcc.Add(BCCEmailId);
                }
                try
                {
                    mailMessage.Attachments.Add(new Attachment(filePath));

                }
                catch (DirectoryNotFoundException DNfEx)
                {
                    WriteLog("Email Error DirectoryNotFoundException: " + DNfEx.Message);
                }
                catch (FileNotFoundException FNFEx)
                {
                    WriteLog("Email Error FileNotFoundException: " + FNFEx.Message);
                }

                using (SmtpClient smtpClient = new SmtpClient("164.162.116.227")
                {
                    Port = 25,
                    UseDefaultCredentials = false,
                    EnableSsl = false,
                    // TargetName = "STARTTLS/smtp.gmail.com",
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Timeout = 600000,
                    Credentials = new NetworkCredential(SenderEmailId, SenderEmailIdPwd)
                })
                {
                    smtpClient.Send(mailMessage);
                }



            }
            catch (Exception ex)
            {
                WriteLog("Error In Sending Email: " + reportDetails.CustomerEmail + ex.Message);
            }
            finally
            {

                if (mailMessage.Attachments != null)
                {
                    for (int i = mailMessage.Attachments.Count - 1; i >= 0; i--)
                    {
                        mailMessage.Attachments[i].Dispose();
                    }
                    mailMessage.Attachments.Clear();
                    mailMessage.Attachments.Dispose();
                }
                mailMessage.Dispose();
                mailMessage = null;

                try
                {
                    if (File.Exists(filePath))
                    {
                        File.Delete(filePath);
                    }
                }
                catch (Exception ex)
                {
                    WriteLog("Finally: " + ex.Message);
                }

            }

        }
        private void Generate_Report(ReportDetails reportDetails)
        {
            try
            {
                WriteLog("Generate Customer Report PDF:");

                DateTime toDate = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                toDate = toDate.AddSeconds(-1); //
                reportDetails.Todate = toDate;

                using (ReportDocument document = new ReportDocument())
                {

                    string reportFilePath = ConfigurationManager.AppSettings["RptFilePath"] + "SOA_ANSAC.rpt";
                    document.Load(reportFilePath);
                    document.SetDatabaseLogon(ConfigurationManager.AppSettings["DB_Username"], ConfigurationManager.AppSettings["DB_Password"], ConfigurationManager.AppSettings["DB_Server"], ConfigurationManager.AppSettings["Company_DB"]);
                    document.SetParameterValue("@AgeDate", reportDetails.Todate.ToString("yyyy-MM-dd HH:mm:ss"));
                    document.SetParameterValue("CardCodeFr@Select * from OCRD where cardtype = 'C' order by cardcode", reportDetails.CardCode);
                    document.SetParameterValue("CardCodeTo@Select * from OCRD where cardtype = 'C' order by cardcode", reportDetails.CardCode);

                    ExportOptions exportOptions = document.ExportOptions;

                    exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;

                    exportOptions.ExportDestinationType = ExportDestinationType.DiskFile;

                    exportOptions.DestinationOptions = new DiskFileDestinationOptions();

                    string FileNameGenerated = reportDetails.ReportName + reportDetails.Todate.ToString("yyyyMMddhhmmss") + ".pdf";

                    reportDetails.ReportFilename = FileNameGenerated;
                    DiskFileDestinationOptions diskFileDestinationOptions = (DiskFileDestinationOptions)document.ExportOptions.DestinationOptions;

                    diskFileDestinationOptions.DiskFileName = ConfigurationManager.AppSettings["RptFilePath"] + FileNameGenerated;

                    document.Export();
                }

                SendReportMail(reportDetails);
            }
            catch (Exception ex)
            {
                WriteLog("Error In Report Generation : " + ex.Message);
            }

        }
        private void RestartProcessSwitch(string result)
        {
            this.Text = result;
            CurrentCustomerIndex++;
            Switch_Process();
        }
        private void Report_Generated(object sender, RunWorkerCompletedEventArgs e)
        {
            try
            {
                if (e.Error == null)
                {
                    //this.Invoke((Action)(() => {
                    //    this.RestartProcessSwitch(e.Result.ToString());
                    //}));
                }
                else
                {
                    this.Invoke((Action)(() => {
                        this.WriteLog(e.Error.Message);
                    }));
                }


            }
            catch (Exception ex)
            {
                WriteLog("Report : " + ex.Message);
            }
        }

        void WriteLog(string ErrMsg)
        {
            try
            {
                string text = ErrMsg;

                if (File.Exists(path))
                {
                    // Create a file to write to.
                    using (StreamWriter sw = File.AppendText(path))
                    {
                        sw.WriteLine(text);
                    }
                }
                else
                {
                    File.Create(path);
                }
            }
            catch (Exception ex)
            {
                WriteLog(ex.Message);
            }

        }


    }
}
