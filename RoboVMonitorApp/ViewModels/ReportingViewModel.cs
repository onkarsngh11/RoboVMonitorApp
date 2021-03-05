using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using RoboVMonitorApp.Models;
using System.Diagnostics;
using System.Windows.Input;
using RoboVMonitorApp.Commands;
using System.Data;
using System.Drawing;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security.Cryptography.X509Certificates;
using System.Web.Script.Serialization;
using System.Net.Security;
using Microsoft.Office.Interop.Excel;
using excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Automation;
using Microsoft.Win32;
using System.Threading;
using SHDocVw;
using mshtml;
using System.Xml.Linq;
using Microsoft.Vbe.Interop;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Windows;

namespace RoboVMonitorApp.ViewModels
{
    public class ReportingViewModel
    {
        #region VARIABLES
        //oces ticket related variables
        public string status = string.Empty;
        public string resolveddate = string.Empty;
        public string downloadurl = string.Empty;
        public string TicketNumber = string.Empty;
        string id = null;
        public string usertype = string.Empty;
        public string action = string.Empty;
        public string ExistingNew = string.Empty;
        public string ExistingExisting = string.Empty;
        public string ExistingReactivate = string.Empty;
        public string ExistingPR = string.Empty;
        public string NewTickets = "0";
        public string ExistingTickets = "0";
        public string ReactivateTickets = "0";
        public string PRTickets = "0";
        private int NewTicketsCount;
        private int ExistingTicketsCount;
        private int ReactivateTicketsCount;
        private int PRTicketsCount;
        List<string> reverseTNs = new List<string>();
        List<string> newTNs = new List<string>();

        //sr ticket related variables
        public string SRNewUser = "0";
        public string SRExistingUser = "0";
        public string SRReactivateUser = "0";
        public string SRPasswordReset = "0";
        public string SPSRNewUser = "0";
        public string SPSRExistingUser = "0";
        public string SPSRReactivateUser = "0";
        public string SPSRPasswordReset = "0";
        public string SPNewTickets = "0";
        public string SPExistingTickets = "0";
        public string SPReactivateTickets= "0";
        public string SPPRTickets = "0";


        public string ExistingSRNew = string.Empty;
        public string ExistingSRExisting = string.Empty;
        public string ExistingSRReactivate = string.Empty;
        public string ExistingSRPR = string.Empty;
        public int FlagforSRTickets = 0;
        private int SRNewUserCount;
        private int SRPasswordResetCount;
        private int SRExistingUserCount;
        private int SRReactivateUserCount;
        //IPA758 related variables
        private int MatchPassCount = 0;
        private int MatchFailCount = 0;
        private int NoMatchPassCount = 0;
        private int NoMatchFailCount = 0;
        private int DiscrepancyPassCount = 0;
        private int DiscrepancyFailCount = 0;
        private int MatchFlag = 0;
        private string MatchFailReason;
        private int NoMatchFlag = 0;
        private string NoMatchFailReason;
        private int DiscrepancyFlag = 0;
        private string DiscrepancyFailReason;
        private string ActualMatchFailReason;
        private string ActualNoMatchFailReason;
        private string ActualDiscrepancyFailReason;
        private string ActualMatchFailReason1 = string.Empty;
        private string ActualMatchFailReason2 = string.Empty;
        private string ActualNoMatchFailReason1 = string.Empty;
        private string ActualNoMatchFailReason2 = string.Empty;
        private string ActualDiscrepanyFailReason1 = string.Empty;
        private string ActualDiscrepancyFailReason2 = string.Empty;
        //Baan related variables
        private string BaanPassTickets = "0";
        private string BaanFailTickets = "0";
        private int BaanPassTicketsCount;
        private int BaanFailTicketsCount;
        //IPA199 related variables
        private int IPA199PassTicketsCount;
        private int IPA199FailTicketsCount;
        //Mixed variables
        public string AlreadyExistingTickets = string.Empty;
        uint DownloadprocessID;
        int processID = 0;
        public int flag = 0;
        public bool isSignIn = false;
        public string pathDownload;
        public string filetofind;
        public string FileName;
        public string ServerName;
        string logpath = @"‪C:\Users\jsonawk\Desktop\LogAST.txt";
        private int Days;
        private string FeildValue;
        private string SelectedUCName;
        DateTime FromDate;
        DateTime ToDate;
        private int ExtractJiraTicketsFromFilterFlag = 0;
        #endregion

        #region Objects
        IWebBrowser2 iwb2;
        IHTMLDocument2 ihd2;
        Process process;
        IntPtr hwndPSCRM;
        excel.Application excelApp = new excel.Application();
        excel.Application SexcelApp = new excel.Application();
        TimeSpan datesDiff;
        public List<string> LastTicketNumbers = null;
        List<string> TNs = new List<string>();
        internal List<string> UCNameList = new List<string>();
        Dictionary<string, string> statusValue = new Dictionary<string, string>();
        System.Data.DataTable dt758 = new System.Data.DataTable();
        System.Data.DataTable dtBaan = new System.Data.DataTable();
        System.Data.DataTable dtBENA = new System.Data.DataTable();
        System.Data.DataTable dt199 = new System.Data.DataTable();
        internal System.Data.DataTable dt = new System.Data.DataTable();
        DateTime TicketsonThatDay = new DateTime();
        #endregion

        #region Imported Dlls
        [DllImport("user32.Dll")]
        public static extern int EnumWindows(IECallBack x, int y);
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("User32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        #endregion

        #region Configurables and Objects
        private string UCNames = ConfigurationManager.AppSettings["UCNames"];
        private static string username = ConfigurationManager.AppSettings["username"].ToString();
        private static string password = ConfigurationManager.AppSettings["password"].ToString();
        private static string EnviormentLG = ConfigurationManager.AppSettings["EnviormentLG"].ToString();
        private string Filter = ConfigurationManager.AppSettings["Filter"].ToString();
        private static string Post_URL = ConfigurationManager.AppSettings["Post_URL_" + EnviormentLG].ToString();
        private static string Favourite_URL = ConfigurationManager.AppSettings["Favourite_URL_" + EnviormentLG].ToString();
        private static string basic = ConfigurationManager.AppSettings["Basic"].ToString();
        private static string appjson = ConfigurationManager.AppSettings["Application_json"].ToString();
        private string Attachment_URL = ConfigurationManager.AppSettings["Attachment_URL"].ToString();
        private static string ResultSize = ConfigurationManager.AppSettings["ResultSize"];
        private string TrackerFileName = ConfigurationManager.AppSettings["TrackerFileNAme"].ToString();
        private string TrackerPath = ConfigurationManager.AppSettings["TrackerPath"].ToString();
        private string BENAFeildName = ConfigurationManager.AppSettings["BENAFeildName"].ToString();
        private string BENAColumns = ConfigurationManager.AppSettings["BENAColumns"].ToString();
        private string BENAFeildValue = ConfigurationManager.AppSettings["BENAFeildValue"].ToString();
        private string IPA758FeildName = ConfigurationManager.AppSettings["IPA758FeildName"].ToString();
        private string IPA758Columns = ConfigurationManager.AppSettings["IPA758Columns"].ToString();
        private string IPA758FeildValue = ConfigurationManager.AppSettings["IPA758FeildValue"].ToString();
        private string BaanColumns = ConfigurationManager.AppSettings["BaanColumns"].ToString();
        private string BaanFeildName = ConfigurationManager.AppSettings["BaanFeildName"].ToString();
        private string BaanFeildValue = ConfigurationManager.AppSettings["BaanFeildValue"].ToString();
        private string IPA199Columns = ConfigurationManager.AppSettings["IPA199Columns"].ToString();
        private string IPA199FeildName = ConfigurationManager.AppSettings["IPA199FeildName"].ToString();
        private string IPA199FeildValue = ConfigurationManager.AppSettings["IPA199FeildValue"].ToString();
        private string AssistedgeTrackerPath = ConfigurationManager.AppSettings["AssistedgeTrackerPath"].ToString();
        private string TemplateSheetName = ConfigurationManager.AppSettings["TemplateSheetName"].ToString();
        private string AssistedgeTrackerName = ConfigurationManager.AppSettings["AssistedgeTrackerName"].ToString();
        #endregion

        #region Property Variables
        private ReportingModel _ReportingControls;
        private bool FullDateRange;
        private bool ChkSpTracker;
        private DateTime SPresolveddate;

        public ReportingModel ReportingControls
        {
            get { return _ReportingControls; }
            set { _ReportingControls = value; }
        }
        #endregion

        public ReportingViewModel()
        {
            dt758.Columns.Add("Category", typeof(string));
            dt758.Columns.Add("Pass Count", typeof(string));
            dt758.Columns.Add("Fail Count", typeof(string));
            dt758.Columns.Add("Total Count", typeof(string));
            dtBaan.Columns.Add("BAAN Pass Count", typeof(string));
            dtBaan.Columns.Add("BAAN Fail Count", typeof(string));
            dtBaan.Columns.Add("BAAN Total Count", typeof(string));
            dt199.Columns.Add("IPA 199 Pass Count", typeof(string));
            dt199.Columns.Add("IPA 199 Fail Count", typeof(string));
            dt199.Columns.Add("IPA 199 Total Count", typeof(string));
            dtBENA.Columns.Add("New User Tickets", typeof(string));
            dtBENA.Columns.Add("Existing User Tickets", typeof(string));
            dtBENA.Columns.Add("Password Reset User Tickets", typeof(string));
            dtBENA.Columns.Add("Reactivate User Tickets", typeof(string));
            foreach (var item in UCNames.Split(','))
            {
                UCNameList.Add(item);
            }
            _ReportingControls = new ReportingModel(UCNameList, SelectedUCName);
            ExtractCommand = new ReportingExtractCommand(this);
            UpdateCommand = new ReportingUpdateCommand(this);
            UpdateCommand2 = new ReportingUpdateCommand2(this);
        }

        public ICommand ExtractCommand
        {
            get;
            private set;
        }

        public ICommand UpdateCommand
        {
            get;
            internal set;
        }

        public ICommand UpdateCommand2
        {
            get;
            internal set;
        }

        public bool CanExtract { get { if (ReportingControls.UCName != null && ReportingControls.ToDate != null && ReportingControls.FromDate <= ReportingControls.ToDate) { return true; } else { return false; } } }

        public bool CanUpdate { get { if (ReportingControls.FromDate != null && ReportingControls.ToDate != null && ReportingControls.ToDate <= DateTime.Now.Date && ReportingControls.FromDate <= ReportingControls.ToDate) { return true; } else { return false; } } }

        public bool CanUpdate2 { get { if (ReportingControls.FromDate != null && ReportingControls.ToDate != null && ReportingControls.ToDate <= DateTime.Now.Date && ReportingControls.FromDate <= ReportingControls.ToDate) { return true; } else { return false; } } }

        public void ExtractData()
        {
            FromDate = _ReportingControls.FromDate;
            ToDate = _ReportingControls.ToDate;
            SelectedUCName = _ReportingControls.UCName;
            FullDateRange = _ReportingControls.FullDateRange;
            if (SelectedUCName == "IPA 758")
            {
                dt758.Clear();
                GetKibanaDetailfor758();
                dt758.Rows.Add("Match", MatchPassCount, MatchFailCount, MatchPassCount + MatchFailCount);
                dt758.Rows.Add("No Match", NoMatchPassCount, NoMatchFailCount, NoMatchPassCount + NoMatchFailCount);
                dt758.Rows.Add("Discrepancy", DiscrepancyPassCount, DiscrepancyFailCount, DiscrepancyPassCount + DiscrepancyFailCount);
                dt758.Rows.Add("Total", MatchPassCount + NoMatchPassCount + DiscrepancyPassCount, MatchFailCount + NoMatchFailCount + DiscrepancyFailCount, MatchPassCount + DiscrepancyPassCount + NoMatchPassCount + MatchFailCount + NoMatchFailCount + DiscrepancyFailCount);
                _ReportingControls.UCDetails = dt758;
            }
            else if (SelectedUCName == "UC 502")
            {
                dtBaan.Clear();
                GetKibanaDetailForBAAN();
                CountTickets();
                dtBaan.Rows.Add(BaanPassTicketsCount, BaanFailTicketsCount, BaanPassTicketsCount + BaanFailTicketsCount);
                _ReportingControls.UCDetails = dtBaan;
            }
            else if (SelectedUCName == "BENA")
            {
                dtBENA.Clear();
                GetKibanaDetailforBENA();
                CountTickets();
                dtBENA.Rows.Add(SRNewUserCount, SRExistingUserCount, SRPasswordResetCount, SRReactivateUserCount);
                _ReportingControls.UCDetails = dtBENA;
            }
            else if (SelectedUCName == "IPA 199")
            {
                dt199.Clear();
                GetKibanaDetailFor199();
                dt199.Rows.Add(IPA199PassTicketsCount, IPA199FailTicketsCount, IPA199PassTicketsCount + IPA199FailTicketsCount);
                _ReportingControls.UCDetails = dt199;
            }
            else
            {
                MessageBox.Show("Data not ready yet.");
            }
        }
        
        public void UpdateTracker()
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to update Assistedge Tracker?", "Confirmation", MessageBoxButton.YesNo);
            if (result == MessageBoxResult.Yes)
            {
                try
                {
                    FromDate = ReportingControls.FromDate;
                    ToDate = ReportingControls.ToDate;
                    datesDiff = ToDate.Subtract(FromDate);
                    Days = datesDiff.Days;
                    int temp = Days;
                    GetKibanaDetailforBENA();
                    GetKibanaDetailfor758();
                    GetKibanaDetailForBAAN();
                    GetTicketDetailsFromJira();
                    CountTickets();
                    CheckandWriteTickets();
                    foreach (Process pr in Process.GetProcesses())
                    {
                        try
                        {
                            if (pr.Id != 0)
                            {
                                if (pr.Id.Equals(processID))
                                {
                                    pr.Kill();
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                    MessageBox.Show("Tracker Updated!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error occurred in UpdateTracker \r\n" + ex);
                }
            }
            else
            {
                MessageBox.Show("No steps taken.");
            }
        }

        private void GetKibanaDetailforBENA()
        {
            try
            {
                SRNewUser = string.Empty;
                SRExistingUser = string.Empty;
                SRReactivateUser = string.Empty;
                SRPasswordReset = string.Empty;
                FromDate = ReportingControls.FromDate;
                ToDate = ReportingControls.ToDate;
                datesDiff = ToDate.Subtract(FromDate);
                ChkSpTracker = _ReportingControls.ChkSpTracker;
                Days = datesDiff.Days;
                int temp = Days;
                FeildValue = BENAFeildValue;
                System.Data.DataTable dtTransaction = new System.Data.DataTable();
                List<string> SRTicketdt = new List<string>();
                dtTransaction.TableName = BENAFeildName;
                CreateJSON objCreateJSON = new CreateJSON();
                if (!ChkSpTracker)
                {
                    for (int i = 0; i <= Days; i++)
                    {
                        TicketsonThatDay = FromDate.AddDays(i);
                        if (Days >= i)
                        {

                            string newIndex = "rpa-trans-" + TicketsonThatDay.ToString("yyyy.MM.dd");
                            ServerName = "j051m201:9200";
                            string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, BENAFeildName, FeildValue, Convert.ToInt32(ResultSize));
                            if (json != string.Empty)
                            {
                                ExtractBENATickets(json, dtTransaction);
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                }
                else
                {
                    SRNewUser = "0";
                    SRExistingUser = "0";
                    SRPasswordReset = "0";
                    SRReactivateUser = "0";
                    string newIndex = "rpa-trans-" + TicketsonThatDay.ToString("yyyy.MM.dd");
                    ServerName = "j051m201:9200";
                    string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, BENAFeildName, FeildValue, Convert.ToInt32(ResultSize));
                    if (json != string.Empty)
                    {
                        ExtractBENATickets(json, dtTransaction);
                        FormSpTicketData();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred in GetKibanaDetailforBENA\r\n" + ex);
            }
        }

        private void ExtractBENATickets(string json,System.Data.DataTable dtTransaction)
        {
            dtTransaction.Merge(JsonToDataTable(json, BENAFeildName, BENAColumns));
            for (int j = 0; j < dtTransaction.Rows.Count; j++)
            {
                if (dtTransaction.Rows[j][4].ToString().Contains("SR") && dtTransaction.Rows[j][9].ToString() != "NA")
                {
                    if (SRNewUser == "0" && dtTransaction.Rows[j][9].ToString() == "BENA_NewUser")
                    {
                        SRNewUser = dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    else if (dtTransaction.Rows[j][9].ToString() == "BENA_NewUser" && !SRNewUser.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)) && !AlreadyExistingTickets.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)))
                    {
                        SRNewUser = SRNewUser + " " + dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    if (SRExistingUser == "0" && dtTransaction.Rows[j][9].ToString() == "BENA_ExistingUser")
                    {
                        SRExistingUser = dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    else if (dtTransaction.Rows[j][9].ToString() == "BENA_ExistingUser" && !SRExistingUser.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)) && !AlreadyExistingTickets.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)))
                    {
                        SRExistingUser = SRExistingUser + " " + dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    if (SRReactivateUser == "0" && dtTransaction.Rows[j][9].ToString() == "BENA_ReactivateUser")
                    {
                        SRReactivateUser = dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    else if (dtTransaction.Rows[j][9].ToString() == "BENA_ReactivateUser" && !SRReactivateUser.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)) && !AlreadyExistingTickets.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)))
                    {
                        SRReactivateUser = SRReactivateUser + " " + dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    if (SRPasswordReset == "0" && dtTransaction.Rows[j][9].ToString() == "BENA_ResetPassword")
                    {
                        SRPasswordReset = dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                    else if (dtTransaction.Rows[j][9].ToString() == "BENA_ResetPassword" && !SRPasswordReset.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)) && !AlreadyExistingTickets.Contains(dtTransaction.Rows[j][4].ToString().Substring(0, 10)))
                    {
                        SRPasswordReset = SRPasswordReset + " " + dtTransaction.Rows[j][4].ToString().Substring(0, 10);
                    }
                }
            }
        }

        private void GetKibanaDetailfor758()
        {
            try
            {
                MatchPassCount = 0;
                MatchFailCount = 0;
                NoMatchFailCount = 0;
                NoMatchPassCount = 0;
                DiscrepancyFailCount = 0;
                DiscrepancyPassCount = 0;
                FeildValue = IPA758FeildValue;
                List<string> SRTicketdt = new List<string>();
                CreateJSON objCreateJSON = new CreateJSON();
                DateTime TicketsonThatDay = new DateTime();
                datesDiff = ToDate.Subtract(FromDate);
                Days = datesDiff.Days;
                for (int i = 0; i <= Days; i++)
                {
                    TicketsonThatDay = FromDate.AddDays(i);
                    if (FullDateRange == false)
                    {
                        string newIndex = "rpa-trans-" + ToDate.ToString("yyyy.MM.dd");
                        ServerName = "j051m201:9200";
                        string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, IPA758FeildName, FeildValue, Convert.ToInt32(ResultSize));

                        if (json != string.Empty)
                        {
                            System.Data.DataTable dtTransaction = new System.Data.DataTable();
                            dtTransaction.TableName = IPA758FeildName;
                            dtTransaction.Merge(JsonToDataTable(json, BENAFeildName, IPA758Columns));
                            for (int j = 0; j < dtTransaction.Rows.Count; j++)
                            {
                                if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("pass"))
                                {
                                    ++MatchPassCount;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("fail"))
                                {
                                    ++MatchFailCount;
                                    MatchFailReason = dtTransaction.Rows[j][11].ToString();
                                    if (MatchFlag == 0 && !MatchFailReason.Equals(dtTransaction.Rows[j][11].ToString()))
                                    {
                                        MatchFailReason = dtTransaction.Rows[j][11].ToString();
                                        MatchFlag++;
                                    }
                                    if (!Regex.IsMatch(MatchFailReason, "\\bObject\\b", RegexOptions.IgnoreCase))
                                    {
                                        ActualMatchFailReason1 = "MFGPro Timed out";
                                    }
                                    else
                                    {
                                        ActualMatchFailReason2 = "IE Error";
                                    }
                                    if (ActualMatchFailReason2 == "") { ActualMatchFailReason = ActualMatchFailReason1; }
                                    else
                                    {
                                        ActualMatchFailReason = ActualMatchFailReason1 + " " + ActualMatchFailReason2;
                                    }
                                    ActualMatchFailReason = ActualMatchFailReason1 + " " + ActualMatchFailReason2;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("no match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("pass"))
                                {
                                    ++NoMatchPassCount;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("no match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("fail"))
                                {
                                    ++NoMatchFailCount;
                                    NoMatchFailReason = "," + dtTransaction.Rows[j][11].ToString();
                                    if (!Regex.IsMatch(NoMatchFailReason, "\\bObject\\b", RegexOptions.IgnoreCase))
                                    {
                                        ActualNoMatchFailReason1 = "MFGPro Timed out";
                                    }
                                    else
                                    {
                                        ActualNoMatchFailReason2 = "IE Error";
                                    }
                                    if (ActualNoMatchFailReason2 == "") { ActualNoMatchFailReason = ActualNoMatchFailReason1; }
                                    else
                                    {
                                        ActualNoMatchFailReason = ActualNoMatchFailReason1 + " " + ActualNoMatchFailReason2;
                                    }
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Contains("discrepancia") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("pass"))
                                {
                                    ++DiscrepancyPassCount;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Contains("discrepancia") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("fail"))
                                {
                                    ++DiscrepancyFailCount;
                                    DiscrepancyFailReason = dtTransaction.Rows[j][11].ToString();
                                    if (!Regex.IsMatch(DiscrepancyFailReason, "\\bObject\\b", RegexOptions.IgnoreCase))
                                    {
                                        ActualDiscrepanyFailReason1 = "MFGPro Timed out";
                                    }
                                    else
                                    {
                                        ActualDiscrepancyFailReason2 = "IE Error";
                                    }
                                    if (ActualDiscrepancyFailReason2 == "") { ActualDiscrepancyFailReason = ActualDiscrepanyFailReason1; }
                                    else
                                    {
                                        ActualDiscrepancyFailReason = ActualDiscrepanyFailReason1 + " " + ActualDiscrepancyFailReason2;
                                    }
                                }
                                else
                                {

                                }
                            }
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    else
                    {
                        string newIndex = "rpa-trans-" + TicketsonThatDay.ToString("yyyy.MM.dd");
                        ServerName = "j051m201:9200";
                        string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, IPA758FeildName, FeildValue, Convert.ToInt32(ResultSize));

                        if (json != string.Empty)
                        {
                            System.Data.DataTable dtTransaction = new System.Data.DataTable();
                            dtTransaction.TableName = IPA758FeildName;
                            dtTransaction.Merge(JsonToDataTable(json, BENAFeildName, IPA758Columns));
                            for (int j = 0; j < dtTransaction.Rows.Count; j++)
                            {
                                if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("pass"))
                                {
                                    ++MatchPassCount;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("fail"))
                                {
                                    ++MatchFailCount;
                                    MatchFailReason = dtTransaction.Rows[j][11].ToString();
                                    if (MatchFlag == 0 && !MatchFailReason.Equals(dtTransaction.Rows[j][11].ToString()))
                                    {
                                        MatchFailReason = dtTransaction.Rows[j][11].ToString();
                                        MatchFlag++;
                                    }
                                    if (!Regex.IsMatch(MatchFailReason, "\\bObject\\b", RegexOptions.IgnoreCase))
                                    {
                                        ActualMatchFailReason1 = "MFGPro Timed out";
                                    }
                                    else
                                    {
                                        ActualMatchFailReason2 = "IE Error";
                                    }
                                    if (ActualMatchFailReason2 == "") { ActualMatchFailReason = ActualMatchFailReason1; }
                                    else
                                    {
                                        ActualMatchFailReason = ActualMatchFailReason1 + " " + ActualMatchFailReason2;
                                    }
                                    ActualMatchFailReason = ActualMatchFailReason1 + " " + ActualMatchFailReason2;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("no match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("pass"))
                                {
                                    ++NoMatchPassCount;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Equals("no match") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("fail"))
                                {
                                    ++NoMatchFailCount;
                                    NoMatchFailReason = "," + dtTransaction.Rows[j][11].ToString();
                                    if (!Regex.IsMatch(NoMatchFailReason, "\\bObject\\b", RegexOptions.IgnoreCase))
                                    {
                                        ActualNoMatchFailReason1 = "MFGPro Timed out";
                                    }
                                    else
                                    {
                                        ActualNoMatchFailReason2 = "IE Error";
                                    }
                                    if (ActualNoMatchFailReason2 == "") { ActualNoMatchFailReason = ActualNoMatchFailReason1; }
                                    else
                                    {
                                        ActualNoMatchFailReason = ActualNoMatchFailReason1 + " " + ActualNoMatchFailReason2;
                                    }
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Contains("discrepancia") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("pass"))
                                {
                                    ++DiscrepancyPassCount;
                                }
                                else if (dtTransaction.Rows[j][10].ToString().ToLower().Contains("discrepancia") && dtTransaction.Rows[j][9].ToString().ToLower().Contains("fail"))
                                {
                                    ++DiscrepancyFailCount;
                                    DiscrepancyFailReason = dtTransaction.Rows[j][11].ToString();
                                    if (!Regex.IsMatch(DiscrepancyFailReason, "\\bObject\\b", RegexOptions.IgnoreCase))
                                    {
                                        ActualDiscrepanyFailReason1 = "MFGPro Timed out";
                                    }
                                    else
                                    {
                                        ActualDiscrepancyFailReason2 = "IE Error";
                                    }
                                    if (ActualDiscrepancyFailReason2 == "") { ActualDiscrepancyFailReason = ActualDiscrepanyFailReason1; }
                                    else
                                    {
                                        ActualDiscrepancyFailReason = ActualDiscrepanyFailReason1 + " " + ActualDiscrepancyFailReason2;
                                    }
                                }
                                else
                                {

                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred in GetKibanaDetailforIPA758\r\n" + ex);
            }
        }

        private void GetKibanaDetailForBAAN()
        {
            try
            {
                BaanPassTickets = string.Empty;
                BaanFailTickets = string.Empty;
                FeildValue = BaanFeildValue;
                System.Data.DataTable dtTransaction = new System.Data.DataTable();
                dtTransaction.TableName = IPA758FeildName;
                CreateJSON objCreateJSON = new CreateJSON();
                DateTime TicketsonThatDay = new DateTime();
                datesDiff = ToDate.Subtract(FromDate);
                Days = datesDiff.Days;
                for (int i = 0; i <= Days; i++)
                {
                    TicketsonThatDay = FromDate.AddDays(i);
                    if (FullDateRange == false)
                    {
                        string newIndex = "rpa-trans-" + ToDate.ToString("yyyy.MM.dd");
                        ServerName = "j051m201:9200";
                        string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, IPA758FeildName, FeildValue, Convert.ToInt32(ResultSize));
                        if (json != string.Empty)
                        {
                            dtTransaction.Merge(JsonToDataTable(json, BENAFeildName, BaanColumns));
                            for (int j = 0; j < dtTransaction.Rows.Count; j++)
                            {
                                if (BaanPassTickets == "0" && dtTransaction.Rows[j][6].ToString() == "Pass" && !BaanPassTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanPassTickets = dtTransaction.Rows[j][1].ToString();
                                }
                                else if (BaanPassTickets != "0" && dtTransaction.Rows[j][6].ToString() == "Pass" && !BaanPassTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanPassTickets += dtTransaction.Rows[j][1].ToString() + " ";
                                }
                                if (BaanFailTickets == "0" && dtTransaction.Rows[j][6].ToString() == "Fail" && !BaanFailTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanFailTickets = dtTransaction.Rows[j][1].ToString();
                                }
                                else if (BaanFailTickets != "0" && dtTransaction.Rows[j][6].ToString() == "Fail" && !BaanFailTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanFailTickets += dtTransaction.Rows[j][1].ToString() + " ";
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                        break;
                    }
                    else
                    {
                        string newIndex = "rpa-trans-" + TicketsonThatDay.ToString("yyyy.MM.dd");
                        ServerName = "j051m201:9200";
                        string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, IPA758FeildName, FeildValue, Convert.ToInt32(ResultSize));
                        if (json != string.Empty)
                        {
                            dtTransaction.Merge(JsonToDataTable(json, BENAFeildName, BaanColumns));
                            for (int j = 0; j < dtTransaction.Rows.Count; j++)
                            {
                                if (BaanPassTickets == "0" && dtTransaction.Rows[j][6].ToString() == "Pass" && !BaanPassTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanPassTickets = dtTransaction.Rows[j][1].ToString();
                                }
                                else if (BaanPassTickets != "0" && dtTransaction.Rows[j][6].ToString() == "Pass" && !BaanPassTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanPassTickets += dtTransaction.Rows[j][1].ToString() + " ";
                                }
                                if (BaanFailTickets == "0" && dtTransaction.Rows[j][6].ToString() == "Fail" && !BaanFailTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanFailTickets = dtTransaction.Rows[j][1].ToString();
                                }
                                else if (BaanFailTickets != "0" && dtTransaction.Rows[j][6].ToString() == "Fail" && !BaanFailTickets.Contains(dtTransaction.Rows[j][1].ToString()))
                                {
                                    BaanFailTickets += dtTransaction.Rows[j][1].ToString() + " ";
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred in GetKibanaDetailforBAAN \r\n" + ex);
            }
        }

        private void GetKibanaDetailFor199()
        {
            try
            {
                IPA199FailTicketsCount = 0;
                IPA199PassTicketsCount = 0;
                FeildValue = IPA199FeildValue;
                CreateJSON objCreateJSON = new CreateJSON();
                DateTime TicketsonThatDay = new DateTime();
                datesDiff = ToDate.Subtract(FromDate);
                Days = datesDiff.Days;
                for (int i = 0; i <= Days; i++)
                {
                    TicketsonThatDay = FromDate.AddDays(i);
                    if (FullDateRange == false)
                    {
                        System.Data.DataTable dtTransaction = new System.Data.DataTable();
                        dtTransaction.TableName = IPA199FeildName;
                        string newIndex = "rpa-trans-" + ToDate.ToString("yyyy.MM.dd");
                        ServerName = "j051m201:9200";
                        string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, IPA199FeildName, FeildValue, Convert.ToInt32(ResultSize));
                        if (json != string.Empty)
                        {
                            dtTransaction.Merge(JsonToDataTable(json, IPA199FeildName, IPA199Columns));
                            for (int j = 0; j < dtTransaction.Rows.Count; j++)
                            {
                                if (dtTransaction.Rows[j][8].ToString() == "Pass")
                                {
                                    ++IPA199PassTicketsCount;
                                }
                                else
                                {
                                    ++IPA199FailTicketsCount;
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                        break;
                    }
                    else
                    {
                        System.Data.DataTable dtTransaction = new System.Data.DataTable();
                        dtTransaction.TableName = IPA199FeildName;
                        string newIndex = "rpa-trans-" + TicketsonThatDay.ToString("yyyy.MM.dd");
                        ServerName = "j051m201:9200";
                        string json = objCreateJSON.GetJson(ServerName.Split(':')[0], ServerName.Split(':')[1], newIndex, IPA758FeildName, FeildValue, Convert.ToInt32(ResultSize));
                        if (json != string.Empty)
                        {
                            dtTransaction.Merge(JsonToDataTable(json, IPA199FeildName, IPA199Columns));
                            for (int j = 0; j < dtTransaction.Rows.Count; j++)
                            {
                                if (dtTransaction.Rows[j][8].ToString() == "Pass")
                                {
                                    ++IPA199PassTicketsCount;
                                }
                                else
                                {
                                    ++IPA199FailTicketsCount;
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred in GetKibanaDetailforBAAN \r\n" + ex);
            }
        }

        public void GetTicketDetailsFromJira()
        {
            try
            {
                
                ChkSpTracker = _ReportingControls.ChkSpTracker;
                if (!ChkSpTracker)
                {
                    #region AST
                    NewTickets = "0";
                    ExistingTickets = "0";
                    ReactivateTickets = "0";
                    PRTickets = "0";
                    TNs = GetAllIssuesFromFilter(Filter);
                    ExtractExistingTickets(out AlreadyExistingTickets);
                    for (int i = 0; i < TNs.Count; i++)
                    {
                        statusValue = GetIssueDetails(TNs[i], out id);
                        status = statusValue["name"];
                        resolveddate = statusValue["resolutiondate"];
                        if ((status.Contains("Resolved") || status.Contains("Closed")))
                        {
                            if (DateTime.Parse(resolveddate).Date <= ToDate.Date && DateTime.Parse(resolveddate).Date >= FromDate.Date)
                            {
                                
                                if (!AlreadyExistingTickets.Contains(TNs[i]))                   //comment if all the tickets for date range are to be extracted
                                {
                                    TicketNumber = TNs[i];
                                    string filedetails = id + "/" + TNs[i] + "_BEEU_Results.xlsx";
                                    filetofind = TNs[i] + "_BEEU_Results.xlsx";
                                    downloadurl = string.Format("{0}/{1}", Attachment_URL, filedetails);
                                    if (id != null)
                                    {
                                        Launch();
                                        SignIn();
                                        Thread.Sleep(1000);
                                        ReadFilePath();
                                        Thread.Sleep(1000);
                                        ExtractUserType();      //also extracts action
                                    }
                                    else
                                    {
                                        usertype = "New";
                                    }
                                    FormOCESTicketData();
                                }
                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            if (status.Contains("Progress") || status.Contains("Assign"))
                            {
                                continue;
                            }
                            else break;
                        }
                    }
                    #endregion
                }
                else
                {
                    DateTime SPFromDate = new DateTime();
                    DateTime SPFromDatereference = new DateTime();
                    int resolvedateflag = 0;
                    NewTickets = "0";
                    ExistingTickets = "0";
                    ReactivateTickets = "0";
                    PRTickets = "0";
                    if (ExtractJiraTicketsFromFilterFlag == 0)
                    {
                        TNs = GetAllIssuesFromFilter(Filter);
                        for (int i = 0; i < TNs.Count; i++)
                        {
                            statusValue = GetIssueDetails(TNs[i], out id);
                            status = statusValue["name"];
                            resolveddate = statusValue["resolutiondate"];
                            reverseTNs.Add(TNs[i]);
                            if (FromDate > DateTime.Parse(resolveddate) && status != "In Progress" && !status.Contains("Closed"))
                            {
                                SPFromDatereference = DateTime.Parse(resolveddate);
                                break;
                            }
                            else
                            {
                                SPFromDate = DateTime.Parse(resolveddate);
                                continue;
                            }
                        }
                        for (int k = reverseTNs.Count; k > 0; k--)
                        {
                            newTNs.Add(reverseTNs[k - 1]);
                        }
                        ExtractJiraTicketsFromFilterFlag = 1;
                    }
                    for (int i = 0; i < newTNs.Count; i++)
                    {
                        statusValue = GetIssueDetails(newTNs[i], out id);
                        status = statusValue["name"];
                        resolveddate = statusValue["resolutiondate"];
                        if ((status.Contains("Resolved") || status.Contains("Closed")))
                        {
                            if (DateTime.Parse(resolveddate).Date <= TicketsonThatDay.Date)
                            {
                                if (DateTime.Parse(resolveddate).Date == TicketsonThatDay.Date)
                                {
                                    TicketNumber = newTNs[i];
                                    string filedetails = id + "/" + newTNs[i] + "_BEEU_Results.xlsx";
                                    filetofind = newTNs[i] + "_BEEU_Results.xlsx";
                                    downloadurl = string.Format("{0}/{1}", Attachment_URL, filedetails);
                                    if (id != null)
                                    {
                                        Launch();
                                        Thread.Sleep(500);
                                        SignIn();
                                        Thread.Sleep(500);
                                        ReadFilePath();
                                        Thread.Sleep(1000);
                                        ExtractUserType();      //also extracts action
                                    }
                                    else
                                    {
                                        usertype = "New";
                                    }
                                    if (usertype == "New")
                                    {
                                        if (NewTickets == "0")
                                        {
                                            NewTickets = TicketNumber;
                                        }
                                        else
                                        {
                                            NewTickets += " " + TicketNumber;
                                        }
                                    }
                                    else if ((usertype.Contains("Existing") && !action.Contains("PasswordReset") && !action.Contains("Reactivate")))
                                    {
                                        if (ExistingTickets == "0")
                                        {
                                            ExistingTickets = TicketNumber;
                                        }
                                        else
                                        {
                                            ExistingTickets += " " + TicketNumber;
                                        }
                                    }
                                    else if (usertype.Contains("Existing") && action.Contains("Reactivate"))
                                    {
                                        if (ReactivateTickets == "0")
                                        {
                                            ReactivateTickets = TicketNumber;
                                        }
                                        else
                                        {
                                            ReactivateTickets += " " + TicketNumber;
                                        }
                                    }
                                    else if (usertype.Contains("Existing") && action.Contains("PasswordReset"))
                                    {
                                        if (PRTickets == "0")
                                        {
                                            PRTickets = TicketNumber;
                                        }
                                        else
                                        {
                                            PRTickets += " " + TicketNumber;
                                        }
                                    }
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                break;                                
                            }
                        }
                        else
                        {
                            if (status.Contains("Progress") || status.Contains("Assign"))
                            {
                                continue;
                            }
                            else break;
                        }
                    }
                    FormSpTicketData();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error occurred in GetTicketDetailsFromJIRA \r\n" + ex);
            }
        }

        public static List<string> GetAllIssuesFromFilter(string filterName)
        {
            List<string> JIRATickets = new List<string>();
            try
            {
                string searchURL = GetFavouriteFilter(filterName);
                searchURL = searchURL + "&startAt=0&maxResults=100";
                string result = string.Empty;
                string postUrl = Convert.ToString(Post_URL);

                HttpClient client = new HttpClient();
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                client.BaseAddress = new System.Uri(postUrl);


                byte[] cred = UTF8Encoding.UTF8.GetBytes(username + ":" + password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(basic, Convert.ToBase64String(cred));
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(appjson));

                HttpResponseMessage response = client.GetAsync(searchURL).Result;
                if (response.IsSuccessStatusCode)
                {
                    result = response.Content.ReadAsStringAsync().Result;
                    System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    js.MaxJsonLength = Int32.MaxValue;
                    FilterObject filter = js.Deserialize<FilterObject>(result);

                    for (int i = 0; i < filter.issues.Count; i++)
                    {
                        JIRATickets.Add(filter.issues[i].key.ToString());
                    }
                }
                else
                {
                    result = response.Content.ReadAsStringAsync().Result;
                }
            }
            catch (Exception ex)
            {
                return JIRATickets;
            }
            return JIRATickets;
        }

        public static string GetFavouriteFilter(string filterName)
        {
            string id = "-1";
            string searchURL = string.Empty;
            try
            {
                string result = string.Empty;
                string postUrl = Convert.ToString(Post_URL);

                HttpClient client = new HttpClient();
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                client.BaseAddress = new System.Uri(postUrl);

                byte[] cred = UTF8Encoding.UTF8.GetBytes(username + ":" + password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue(basic, Convert.ToBase64String(cred));
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(appjson));

                HttpResponseMessage response = client.GetAsync(Favourite_URL).Result;
                if (response.IsSuccessStatusCode)
                {
                    result = response.Content.ReadAsStringAsync().Result;
                    System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    AllFavouriteFilter[] filter = js.Deserialize<AllFavouriteFilter[]>(result);

                    for (int i = 0; i < filter.Length; i++)
                    {
                        if (filter[i].name.ToString().ToLower().Equals(filterName.ToLower()))
                        {
                            id = filter[i].id.ToString();
                            searchURL = filter[i].searchUrl.ToString();
                            break;
                        }
                    }
                }
                else
                {
                    result = response.Content.ReadAsStringAsync().Result;
                }
                return searchURL;
            }
            catch (Exception ex)
            {
                string result = ex.InnerException.ToString();
                return searchURL;
            }
        }

        public static Dictionary<string, string> GetIssueDetails(string TicketNumber, out string id)
        {
            id = null;
            Dictionary<string, string> statusValue = new Dictionary<string, string>();
            try
            {
                string result = string.Empty;
                string postUrl = Convert.ToString(ConfigurationManager.AppSettings["URL_PROD"]);

                HttpClient client = new HttpClient();
                ServicePointManager.ServerCertificateValidationCallback = delegate (object s, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
                client.BaseAddress = new System.Uri(postUrl);

                byte[] cred = UTF8Encoding.UTF8.GetBytes(username + ":" + password);
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue((basic), Convert.ToBase64String(cred));
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue(appjson));

                HttpResponseMessage response = client.GetAsync(Convert.ToString(ConfigurationManager.AppSettings["URL_PROD"]) + Convert.ToString(ConfigurationManager.AppSettings["RestApi"]) + TicketNumber + "/").Result;//To be configured
                if (response.IsSuccessStatusCode)
                {
                    result = response.Content.ReadAsStringAsync().Result;
                    System.Web.Script.Serialization.JavaScriptSerializer js = new System.Web.Script.Serialization.JavaScriptSerializer();
                    FavouriteField fieldsobject = js.Deserialize<FavouriteField>(result);
                    Newtonsoft.Json.Linq.JObject root = Newtonsoft.Json.Linq.JObject.Parse(result);
                    string Mainjitems=null;
                    foreach (JToken childtoken in root.Children())
                    {
                        if (((JProperty)childtoken).Path == "fields")
                        {
                            foreach (JToken childtokenhits in childtoken.Children())
                            {
                                if (((JObject)childtokenhits).Path == "fields")
                                {
                                    foreach (JToken childtokenhits1 in childtokenhits.Children())
                                    {
                                        if (((JProperty)childtokenhits1).Path == "fields.resolutiondate")
                                        {
                                            Mainjitems = ((JProperty)childtokenhits1).Value.ToString();
                                        }
                                        if (Mainjitems != null)
                                            break;
                                    }
                                }
                                if (Mainjitems != null)
                                    break;
                            }
                        }
                        if (Mainjitems != null)
                            break;
                    }
                    statusValue.Add("name", fieldsobject.fields.status.name.ToString());
                    for (int i = 0; i < fieldsobject.fields.attachment.Count; i++)
                    {
                        if (fieldsobject.fields.attachment[i].filename.Contains(".xlsx"))
                        {
                            id = fieldsobject.fields.attachment[i].id;
                            break;
                        }
                    }
                    if (Mainjitems != null && Mainjitems!="")
                    {
                        statusValue.Add("resolutiondate", Mainjitems.Substring(0, 10));
                    }
                    else
                    {
                        statusValue.Add("resolutiondate", fieldsobject.fields.updated.ToString().Substring(0, 10));
                    }
                }
                else
                {
                    id = null;
                    result = null;
                }
            }
            catch (Exception ex)
            {
                id = null;
                throw;
            }
            return statusValue;
        }

        public void Launch()
        {
            foreach (Process pr in Process.GetProcesses())
            {
                try
                {
                    if (pr.Id != 0)
                    {
                        if (pr.Id.Equals(processID))
                        {
                            pr.Kill();
                        }
                    }
                }
                catch (Exception ex)
                {
                }
            }
            process = Process.Start(@"iexplore.exe", "-private -nomerge http://ocoejira.johnsoncontrols.com/login.jsp?os_destination=%2Fsecure%2FDashboard.jspa");
            Thread.Sleep(2500);
            hwndPSCRM = process.Handle;
            processID = process.Id;
            iwb2 = new InternetExplorer();
            (iwb2 as InternetExplorer).NewProcess += Ie_NewProcess;
            (iwb2 as InternetExplorer).DocumentComplete += Ie_DocumentComplete;
            (iwb2 as InternetExplorer).Navigate2("http://ocoejira.johnsoncontrols.com/login.jsp?os_destination=%2Fsecure%2FDashboard.jspa");
            Thread.Sleep(1500);
        }

        public void SignIn()
        {
            try
            {
                ihd2 = iwb2.Document as IHTMLDocument2;                
                HTMLBody hb = ihd2.body as HTMLBody;
                HTMLInputElement id = FindInputControl(null, "login-form-username", ihd2);
                HTMLInputElement pwd = FindInputControl(null, "login-form-password", ihd2);
                if (id != null && pwd != null && !string.IsNullOrEmpty(username) && !string.IsNullOrEmpty(password))
                {
                    id.value = "";
                    id.value = username;
                    pwd.value = "";
                    pwd.value = password;
                    HTMLInputElement btnLogin = FindInputElementbyType("submit", ihd2);
                    if (btnLogin != null)
                    {
                        btnLogin.click();
                        Thread.Sleep(1500);
                    }
                }
                (iwb2 as InternetExplorer).Navigate2(downloadurl);
                Thread.Sleep(3000);
                IntPtr DownloadIEhandle = IntPtr.Zero;
                for (int i = 0; i < 20; i++)
                {
                    DownloadIEhandle = FindWindow("#32770", "Internet Explorer");
                    if (DownloadIEhandle != IntPtr.Zero)
                    {
                        break;
                    }
                    Thread.Sleep(500);
                }
                GetWindowThreadProcessId(DownloadIEhandle, out DownloadprocessID);
                AutomationElement iedopen = AutomationElement.FromHandle(DownloadIEhandle);
                PropertyCondition iedoPC1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                PropertyCondition iedoPC2 = new PropertyCondition(AutomationElement.NameProperty, "Save");
                AutomationElement iedClickopen = iedopen.FindFirst(TreeScope.Descendants, new AndCondition(iedoPC1, iedoPC2));
                Win32.PostMessage((IntPtr)iedClickopen.Current.NativeWindowHandle, Win32.BN_CLICKED, 0, 0);
                Thread.Sleep(2000);

            }
            catch (Exception ex)
            {

            }
        }

        public string ReadFilePath()
        {
            try
            {
                String path = @"Software\Microsoft\Internet Explorer\Main";    //Code to get Browser Download Location
                RegistryKey rKey = Registry.CurrentUser.OpenSubKey(path);
                if (rKey != null)
                    pathDownload = (String)rKey.GetValue("Default Download Directory");
                if (String.IsNullOrEmpty(pathDownload))
                    pathDownload = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\downloads";
                var directory = new DirectoryInfo(pathDownload);
                var fileNameList = directory.GetFiles().OrderByDescending(f => f.LastWriteTime);
                Thread.Sleep(500);
                if (fileNameList == null)
                {
                    fileNameList = directory.GetFiles().OrderByDescending(f => f.LastWriteTime);
                }
                foreach (var fileName in fileNameList)
                {
                    if (Convert.ToString(fileName).Contains(TicketNumber))
                    {
                        FileName = fileName.ToString();
                        break;
                    }
                }
                pathDownload = pathDownload + "\\" + FileName;
            }
            catch (Exception ex)
            {
            }
            return pathDownload;
        }

        public void ExtractUserType()
        {
            try
            {
                Workbook wb = excelApp.Workbooks.Open(pathDownload);
                Thread.Sleep(1000);
                Worksheet sh = null;
                bool isSheetPresent = false;
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet != null && !string.IsNullOrEmpty(sheet.Name) && sheet.Name.Equals("Sheet1"))
                    {
                        isSheetPresent = true;
                        sh = sheet;
                    }
                    if (isSheetPresent)
                    {
                        List<string> missingFields = new List<string>();
                        Range userData = sh.get_Range("B2", "D2");
                        if (!string.IsNullOrEmpty(Convert.ToString(userData.Value2[1, 1])))
                        {
                            usertype = Convert.ToString(userData.Value2[1, 1]);
                            action = Convert.ToString(userData.Value2[1, 3]);
                        }
                    }
                }
                wb.Close();
            }
            catch (Exception ex)
            {
            }
        }

        public string ExtractExistingTickets(out string AlreadyExistingTickets)
        {
            try
            {
                Workbook wb = excelApp.Workbooks.Open(TrackerPath);
                foreach (Worksheet sheet in wb.Sheets)
                {
                    if (sheet != null && !string.IsNullOrEmpty(sheet.Name) && sheet.Name.Equals("AE Use Case"))
                    {
                        Range userData = sheet.get_Range("C25", "C35");
                        ExistingNew = Convert.ToString(userData[1][1].Value);
                        ExistingExisting = Convert.ToString(userData[2][1].Value);
                        ExistingReactivate = Convert.ToString(userData[3][1].Value);
                        ExistingPR = Convert.ToString(userData[4][1].Value);
                        ExistingSRNew = Convert.ToString(userData[7][1].Value);
                        ExistingSRExisting = Convert.ToString(userData[8][1].Value);
                        ExistingSRReactivate = Convert.ToString(userData[9][1].Value);
                        ExistingSRPR = Convert.ToString(userData[10][1].Value);
                        break;
                    }
                }
                AlreadyExistingTickets = ExistingNew + ExistingExisting + ExistingReactivate + ExistingPR + ExistingSRNew + ExistingSRExisting + ExistingSRReactivate + ExistingSRPR;
            }
            catch (Exception ex)
            {
                AlreadyExistingTickets = null;
            }
            return AlreadyExistingTickets;
        }

        private void Ie_DocumentComplete(object pDisp, ref object URL)
        {
            try
            {
                if (URL.ToString().Contains("http://ocoejira.johnsoncontrols.com/Secure/Dashboard"))
                {


                }
            }
            catch (Exception ex)
            {

            }
        }

        private void Ie_NewProcess(int lCauseFlag, object pWB2, ref bool Cancel)
        {
            try
            {
                iwb2 = (SHDocVw.InternetExplorer)pWB2;
                (iwb2 as InternetExplorer).DocumentComplete += new DWebBrowserEvents2_DocumentCompleteEventHandler(Ie_DocumentComplete);
                iwb2.ToolBar = 0;
            }
            catch (Exception ex)
            {

            }
        }

        private HTMLInputElement FindInputElementbyType(string type, IHTMLDocument2 ihd2)
        {
            HTMLInputElement hie = null;
            try
            {
                HTMLBody hb = ihd2.body as HTMLBody;
                IHTMLElementCollection icoll = hb.getElementsByTagName("input");
                if (icoll != null && icoll.length != 0)
                {
                    foreach (IHTMLElement element in icoll)
                    {
                        try
                        {
                            if (element != null && element is HTMLInputElement)
                            {
                                HTMLInputElement inputelement = element as HTMLInputElement;
                                if (!string.IsNullOrEmpty(type))
                                {
                                    if (inputelement.type != null)
                                    {
                                        if (inputelement.type.ToLower().Contains(type.ToLower()))
                                        {
                                            hie = inputelement;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return hie;
        }

        public HTMLInputElement FindInputControl(string controlName, string controlId, IHTMLDocument2 ihd2)
        {
            HTMLInputElement hie = null;
            try
            {
                HTMLBody hb = ihd2.body as HTMLBody;
                IHTMLElementCollection icoll = hb.getElementsByTagName("input");
                if (icoll != null && icoll.length != 0)
                {
                    foreach (IHTMLElement element in icoll)
                    {
                        try
                        {
                            if (element.id != null && element is HTMLInputElement)
                            {
                                HTMLInputElement inputelement = element as HTMLInputElement;
                                if (!string.IsNullOrEmpty(controlName))
                                {
                                    if (inputelement.name != null)
                                    {
                                        if (inputelement.name.ToLower().Contains(controlName.ToLower()))
                                        {
                                            hie = inputelement;
                                        }
                                    }
                                }
                                else if (!string.IsNullOrEmpty(controlId))
                                {
                                    if (inputelement.id != null)
                                    {
                                        if (inputelement.id.ToLower().Contains(controlId.ToLower()))
                                        {
                                            hie = inputelement;
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return hie;
        }

        private void FormOCESTicketData()
        {
            if (!AlreadyExistingTickets.Contains(TicketNumber))
            {
                if (usertype == "New")
                {
                    if (NewTickets == "0")
                    {
                        NewTickets = TicketNumber;
                    }
                    else
                    {
                        NewTickets += " " + TicketNumber;
                    }
                }
                else if ((usertype.Contains("Existing") && !action.Contains("PasswordReset") && !action.Contains("Reactivate")))
                {
                    if (ExistingTickets == "0")
                    {
                        ExistingTickets = TicketNumber;
                    }
                    else
                    {
                        ExistingTickets += " " + TicketNumber;
                    }
                }
                else if (usertype.Contains("Existing") && action.Contains("Reactivate"))
                {
                    if (ReactivateTickets == "0")
                    {
                        ReactivateTickets = TicketNumber;
                    }
                    else
                    {
                        ReactivateTickets += " " + TicketNumber;
                    }
                }
                else if (usertype.Contains("Existing") && action.Contains("PasswordReset"))
                {
                    if (PRTickets == "0")
                    {
                        PRTickets = TicketNumber;
                    }
                    else
                    {
                        PRTickets += " " + TicketNumber;
                    }
                }
            }
        }

        public System.Data.DataTable JsonToDataTable(string json, string tableName, string ColumnNames)
        {
            bool columnsCreated = false;
            Newtonsoft.Json.Linq.JArray Mainjitems = null;
            System.Data.DataTable dt = new System.Data.DataTable(tableName);
            try
            {
                Newtonsoft.Json.Linq.JObject root = Newtonsoft.Json.Linq.JObject.Parse(json);
                foreach (JToken childtoken in root.Children())
                {
                    if (((JProperty)childtoken).Path == "hits")
                    {
                        foreach (JToken childtokenhits in childtoken.Children())
                        {
                            if (((JObject)childtokenhits).Path == "hits")
                            {
                                foreach (JToken childtokenhits1 in childtokenhits.Children())
                                {
                                    if (((JProperty)childtokenhits1).Path == "hits.hits")
                                    {
                                        foreach (JToken childtokenhits2 in childtokenhits1.Children())
                                        {
                                            if (((JArray)childtokenhits2).Path == "hits.hits")
                                            {
                                                Mainjitems = (JArray)childtokenhits2;
                                                break;
                                            }
                                        }
                                    }
                                    if (Mainjitems != null)
                                        break;
                                }
                            }
                            if (Mainjitems != null)
                                break;
                        }
                    }
                    if (Mainjitems != null)
                        break;
                }
                Newtonsoft.Json.Linq.JObject item = default(Newtonsoft.Json.Linq.JObject);
                Newtonsoft.Json.Linq.JToken jtoken = default(Newtonsoft.Json.Linq.JToken);

                for (int i = 0; i <= Mainjitems.Count - 1; i++)
                {
                    item = (Newtonsoft.Json.Linq.JObject)Mainjitems[i];
                    Dictionary<string, string> additionalcolumns = new Dictionary<string, string>();
                    foreach (Newtonsoft.Json.Linq.JToken _childone in item.Children())
                    {
                        if (ColumnNames.Contains(((JProperty)_childone).Name.ToString()))
                        {
                            additionalcolumns.Add(((JProperty)_childone).Name.ToString(), Convert.ToString(((Newtonsoft.Json.Linq.JProperty)_childone).Value));
                        }
                    }

                    if (((JObject)item).Path == "hits.hits[" + i + "]")
                    {
                        foreach (JToken childtokenhits1 in item.Children())
                        {
                            if (((JProperty)childtokenhits1).Path == "hits.hits[" + i + "]._source")
                            {
                                jtoken = childtokenhits1.First;
                                // Create the new row, put the values into the columns then add the row to the DataTable
                                DataRow dr = dt.NewRow();
                                if (!columnsCreated)
                                {
                                    foreach (string _key in additionalcolumns.Keys)
                                    {
                                        dt.Columns.Add(new DataColumn(_key));
                                    }
                                    foreach (JToken childtokenhits2 in jtoken.Children())
                                    {
                                        if (ColumnNames.Contains(((Newtonsoft.Json.Linq.JProperty)childtokenhits2).Name.ToString()))
                                            dt.Columns.Add(new DataColumn(((Newtonsoft.Json.Linq.JProperty)childtokenhits2).Name.ToString()));
                                    }
                                    columnsCreated = true;
                                }

                                foreach (string _key in additionalcolumns.Keys)
                                {
                                    dr[_key] = Convert.ToString(additionalcolumns[_key]);
                                }
                                foreach (JToken childtokenhits2 in jtoken.Children())
                                {
                                    if (ColumnNames.Contains(((Newtonsoft.Json.Linq.JProperty)childtokenhits2).Name.ToString()))
                                        dr[((Newtonsoft.Json.Linq.JProperty)childtokenhits2).Name.ToString()] = Convert.ToString(((Newtonsoft.Json.Linq.JProperty)childtokenhits2).Value);
                                }
                                dt.Rows.Add(dr);
                                break;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
            }
            return dt;

        }

        private void CountTickets()
        {
            NewTicketsCount = NewTickets.Split(' ')[0] == "0" ? 0 : NewTickets.Split(' ').Count();
            ExistingTicketsCount = ExistingTickets.Split(' ')[0] == "0" ? 0 : ExistingTickets.Split(' ').Count();
            PRTicketsCount = PRTickets.Split(' ')[0] == "0" ? 0 : PRTickets.Split(' ').Count();
            ReactivateTicketsCount = ReactivateTickets.Split(' ')[0] == "0" ? 0 : ReactivateTickets.Split(' ').Count();
            SRNewUserCount = SRNewUser.Split(' ')[0] == "0" || SRNewUser == "" ? 0 : SRNewUser.Split(' ').Count();
            SRExistingUserCount = SRExistingUser.Split(' ')[0] == "0" || SRExistingUser == "" ? 0 : SRExistingUser.Split(' ').Count();
            SRPasswordResetCount = SRPasswordReset.Split(' ')[0] == "0" || SRPasswordReset == "" ? 0 : SRPasswordReset.Split(' ').Count();
            SRReactivateUserCount = SRReactivateUser.Split(' ')[0] == "0" || SRReactivateUser == "" ? 0 : SRReactivateUser.Split(' ').Count();
            BaanPassTicketsCount = BaanPassTickets.Split(' ')[0] == "0" || BaanPassTickets == "" ? 0 : BaanPassTickets.Substring(0, BaanPassTickets.Length - 1).Split(' ').Count();
            BaanFailTicketsCount = BaanFailTickets.Split(' ')[0] == "0" || BaanFailTickets == "" ? 0 : BaanFailTickets.Substring(0, BaanFailTickets.Length - 1).Split(' ').Count();
        }

        public void CheckandWriteTickets()
        {
            try
            {
                IntPtr ASTHandle = IntPtr.Zero;
                for (int i = 0; i < 10; i++)
                {
                    ASTHandle = FindWindow("#32770", "Microsoft Excel");
                    if (ASTHandle != IntPtr.Zero)
                    {
                        break;
                    }
                    Thread.Sleep(100);
                }
                if ((ASTHandle != IntPtr.Zero))
                {
                    AutomationElement ASTyes = AutomationElement.FromHandle(ASTHandle);
                    PropertyCondition ASTMEPC1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                    PropertyCondition ASTMEPC2 = new PropertyCondition(AutomationElement.NameProperty, "Yes");
                    AutomationElement iedClickopen = ASTyes.FindFirst(TreeScope.Descendants, new AndCondition(ASTMEPC1, ASTMEPC2));
                    Win32.PostMessage((IntPtr)iedClickopen.Current.NativeWindowHandle, Win32.BN_CLICKED, 0, 0);
                    Thread.Sleep(100);
                }
                else
                {
                    Workbook wb = excelApp.Workbooks.Open(TrackerPath);
                    foreach (excel.Worksheet sheet in wb.Sheets)
                    {
                        if (sheet != null && !string.IsNullOrEmpty(sheet.Name) && sheet.Name.Equals("AE Use Case"))
                        {
                            if (FlagforSRTickets == 0)
                            {
                                sheet.Cells[25, 3].Value = "0";
                                sheet.Cells[25, 3].Value = NewTickets;
                                sheet.Cells[26, 3].Value = "0";
                                sheet.Cells[26, 3].Value = ExistingTickets;
                                sheet.Cells[27, 3].Value = "0";
                                sheet.Cells[27, 3].Value = ReactivateTickets;
                                sheet.Cells[28, 3].Value = "0";
                                sheet.Cells[28, 3].Value = PRTickets;
                                sheet.Cells[32, 3].Value = "0";
                                sheet.Cells[32, 3].Value = SRNewUser;
                                sheet.Cells[33, 3].Value = "0";
                                sheet.Cells[33, 3].Value = SRExistingUser;
                                sheet.Cells[34, 3].Value = "0";
                                sheet.Cells[34, 3].Value = SRReactivateUser;
                                sheet.Cells[35, 3].Value = "0";
                                sheet.Cells[35, 3].Value = SRPasswordReset;

                                sheet.Cells[44, 4].Value = "0";
                                sheet.Cells[44, 4].Value = MatchPassCount;
                                sheet.Cells[45, 4].Value = "0";
                                sheet.Cells[45, 4].Value = NoMatchPassCount;
                                sheet.Cells[46, 4].Value = "0";
                                sheet.Cells[46, 4].Value = DiscrepancyPassCount;
                                sheet.Cells[44, 5].Value = "0";
                                sheet.Cells[44, 5].Value = MatchFailCount;
                                sheet.Cells[45, 5].Value = "0";
                                sheet.Cells[45, 5].Value = NoMatchFailCount;
                                sheet.Cells[46, 5].Value = "0";
                                sheet.Cells[46, 5].Value = DiscrepancyFailCount;
                                sheet.Cells[44, 6].Value = 0;
                                sheet.Cells[44, 6].Value = ActualMatchFailReason;
                                sheet.Cells[45, 6].Value = 0;
                                sheet.Cells[45, 6].Value = ActualNoMatchFailReason;
                                sheet.Cells[46, 6].Value = 0;
                                sheet.Cells[46, 6].Value = ActualDiscrepancyFailReason;

                                sheet.Cells[48, 4].Value = 0;
                                sheet.Cells[48, 4].Value = BaanPassTicketsCount;
                                sheet.Cells[48, 5].Value = 0;
                                sheet.Cells[48, 5].Value = BaanFailTicketsCount;
                                sheet.Cells[49, 4].Value = 0;
                                sheet.Cells[49, 4].Value = BaanPassTickets;
                                sheet.Cells[49, 5].Value = 0;
                                sheet.Cells[49, 5].Value = BaanFailTickets;

                                int lastColumn = sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                                for (int j = 1; j <= lastColumn; j++)
                                {
                                    if (sheet.Cells[4, j].Value == null)
                                    {
                                        sheet.Cells[4, j].Value = "0";
                                        sheet.Cells[5, j].Value = NewTicketsCount;
                                        sheet.Cells[6, j].Value = ExistingTicketsCount;
                                        sheet.Cells[7, j].Value = ReactivateTicketsCount;
                                        sheet.Cells[8, j].Value = PRTicketsCount;
                                        sheet.Cells[9, j].Value = "0";
                                        sheet.Cells[11, j].Value = "1";
                                        sheet.Cells[12, j].Value = SRNewUserCount;
                                        sheet.Cells[13, j].Value = SRExistingUserCount;
                                        sheet.Cells[14, j].Value = SRReactivateUserCount;
                                        sheet.Cells[15, j].Value = SRPasswordResetCount;
                                        sheet.Cells[16, j].Value = "0";
                                        sheet.Cells[17, j].Value = "0";
                                        sheet.Cells[18, j].Value = "0";
                                        sheet.Cells[19, j].Value = "0";
                                        break;
                                    }
                                    FlagforSRTickets = 1;
                                }
                            }
                        }
                    }
                }
                foreach (Workbook item in excelApp.Workbooks)
                {
                    try
                    {
                        if (item.Name.Contains(TrackerFileName))
                        {
                            item.Save();
                            item.Close();
                        }
                        else
                        {
                            item.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        #region Update tracker in Sharepoint
        public void UpdateTracker2()
        {
            MessageBoxResult result = MessageBox.Show("Are you sure you want to update Assistedge Tracker?", "Confirmation", MessageBoxButton.YesNo);
            if (result==MessageBoxResult.Yes)
            {
                FromDate = ReportingControls.FromDate;
                ToDate = ReportingControls.ToDate;
                datesDiff = ToDate.Subtract(FromDate);
                Workbook wb = SexcelApp.Workbooks.Open(AssistedgeTrackerPath);
                Thread.Sleep(500);
                Worksheet sh = null;
                bool isSheetPresent = false;
                for (int i = 0; i <= Days; i++)
                {
                    TicketsonThatDay = FromDate.AddDays(i);
                    foreach (Worksheet sheet in wb.Sheets)
                    {
                        if (sheet != null && !string.IsNullOrEmpty(sheet.Name) && sheet.Name.Equals(TemplateSheetName))
                        {
                            isSheetPresent = true;
                            sh = sheet;
                        }
                        if (isSheetPresent)
                        {
                            GetTicketDetailsFromJira();
                            GetKibanaDetailforBENA();
                            Range userDatafrom = sh.get_Range("A1", "F50");
                            Worksheet newWorksheet;
                            newWorksheet = (Worksheet)wb.Worksheets.Add();
                            newWorksheet.Name = TicketsonThatDay.ToString("dd") + TicketsonThatDay.ToString("MMM");
                            Range userDatato = newWorksheet.get_Range("A1", "F50");
                            userDatafrom.Copy(userDatato);
                            int lastColumn = newWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
                            CountTickets();
                            for (int j = 1; j <= lastColumn; j++)
                            {
                                if (sheet.Cells[4, j].Value == null)
                                {
                                    newWorksheet.Cells[3, j].Value = "0";
                                    newWorksheet.Cells[4, j].Value = NewTicketsCount;
                                    newWorksheet.Cells[5, j].Value = ExistingTicketsCount;
                                    newWorksheet.Cells[6, j].Value = ReactivateTicketsCount;
                                    newWorksheet.Cells[7, j].Value = PRTicketsCount;
                                    newWorksheet.Cells[8, j].Value = "0";
                                    newWorksheet.Cells[9, j].Value = "1";
                                    newWorksheet.Cells[10, j].Value = "0";
                                    newWorksheet.Cells[11, j].Value = SRNewUserCount;
                                    newWorksheet.Cells[12, j].Value = SRExistingUserCount;
                                    newWorksheet.Cells[13, j].Value = SRReactivateUserCount;
                                    newWorksheet.Cells[14, j].Value = SRPasswordResetCount;
                                    newWorksheet.Cells[15, j].Value = "0";
                                    newWorksheet.Cells[16, j].Value = "0";
                                    newWorksheet.Cells[17, j].Value = "0";
                                    newWorksheet.Cells[18, j].Value = "0";
                                    break;
                                }
                            }
                            newWorksheet.Cells[23, 3].Value = "0";
                            newWorksheet.Cells[23, 3].Value = SPNewTickets;
                            newWorksheet.Cells[24, 3].Value = "0";
                            newWorksheet.Cells[24, 3].Value = SPExistingTickets;
                            newWorksheet.Cells[25, 3].Value = "0";
                            newWorksheet.Cells[25, 3].Value = SPReactivateTickets;
                            newWorksheet.Cells[26, 3].Value = "0";
                            newWorksheet.Cells[26, 3].Value = SPPRTickets;
                            newWorksheet.Cells[30, 3].Value = "0";
                            newWorksheet.Cells[30, 3].Value = SPSRNewUser;
                            newWorksheet.Cells[31, 3].Value = "0";
                            newWorksheet.Cells[31, 3].Value = SPSRExistingUser;
                            newWorksheet.Cells[32, 3].Value = "0";
                            newWorksheet.Cells[32, 3].Value = SPSRReactivateUser;
                            newWorksheet.Cells[33, 3].Value = "0";
                            newWorksheet.Cells[33, 3].Value = SPSRPasswordReset;
                        }
                        isSheetPresent = false;
                    }
                }
                foreach (Workbook item in SexcelApp.Workbooks)
                {
                    try
                    {
                        if (item.Name.Contains(AssistedgeTrackerName))
                        {
                            item.Save();
                            item.Close();
                        }
                        else
                        {
                            item.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }
                MessageBox.Show("Tracker Updated.");
            }
            else
            {
                MessageBox.Show("No steps taken");
            }
        }

        private void FormSpTicketData()
        {
            SPSRNewUser = SRNewUser;
            SPSRExistingUser = SRExistingUser;
            SPSRReactivateUser = SRReactivateUser;
            SPSRPasswordReset = SRPasswordReset;
            SPNewTickets = NewTickets;
            SPExistingTickets = ExistingTickets;
            SPReactivateTickets = ReactivateTickets;
            SPPRTickets = PRTickets;
        }

        #endregion

        #region Classes for Jira eTicket extraction
        public class FilterObject
        {
            public List<FavouriteFilterIssues> issues { get; set; }
        }

        public class FavouriteFilterIssues
        {
            public string key { get; set; }
        }

        public class FavouriteField
        {
            public Fields fields { get; set; }
            public TicketStatus status { get; set; }
            public Attachment attachment { get; set; }
        }

        public class Fields
        {
            public string updated;
            public TicketStatus status { get; set; }
            public List<Attachment> attachment { get; set; }
        }

        public class TicketStatus
        {
            public string name { get; set; }
            public int id { get; set; }
        }

        public class Attachment
        {
            public string id { get; set; }
            public string filename { get; set; }
        }

        class AllFavouriteFilter
        {
            public string id { get; set; }
            public string name { get; set; }
            public string searchUrl { get; set; }
        }
        #endregion
    }
}
