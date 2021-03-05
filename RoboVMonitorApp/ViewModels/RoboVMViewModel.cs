using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Text.RegularExpressions;
using System.Windows;
using RoboVMonitorApp.Models;
using System.Windows.Input;
using RoboVMonitorApp.Commands;
using System.Timers;

namespace RoboVMonitorApp.ViewModels
{
    public class RoboVMViewModel : Window
    {
        #region Imported Dlls
        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("User32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        #endregion

        #region Configurable variables
        private string A4201002_password = ConfigurationManager.AppSettings["A4201002_password"];
        private string A3000156_password = ConfigurationManager.AppSettings["A3000156_password"];
        private string A3051025_password = ConfigurationManager.AppSettings["A3051025_password"];
        private string A3000096_password = ConfigurationManager.AppSettings["A3000096_password"];
        private string A3000102_password = ConfigurationManager.AppSettings["A3000102_password"];
        private string jsonawk_password = ConfigurationManager.AppSettings["jsonawk_password"];
        private string jnaths_password = ConfigurationManager.AppSettings["jnaths_password"];
        private string processName_RoboView = ConfigurationManager.AppSettings["processName_RoboView"];
        private string processName_RPAAgent = ConfigurationManager.AppSettings["processName_RPAAgent"];
        private string sessionStatuses = ConfigurationManager.AppSettings["sessionStatuses"];
        private string MachineNames = ConfigurationManager.AppSettings["MachineNames"];
        private string AccountNames = ConfigurationManager.AppSettings["AccountNames"];
        private string DomainNames = ConfigurationManager.AppSettings["DomainNames"];
        private string SEPath = ConfigurationManager.AppSettings["SEPath"];
        private string ConfigPath = ConfigurationManager.AppSettings["ConfigPath"];
        private string ConfigFileName = ConfigurationManager.AppSettings["configFileName"];
        private string[] ConfigFile;
        #endregion

        #region Robo related variables
        private string RoboSEprocessStatus;
        private string RobotSEPath;
        private string RobotName = string.Empty;
        private string[] AccountKeywords;
        string[] lines;
        private string SessionValues;
        private string SessionUserName;
        private string Password;
        private int RoboSEprocessID = 0;
        private string newOwnerNameValue;
        private int RPAAgentprocessID = 0;
        private string RPAAgentOwnerName="Not Found";
        private string RoboSEOwnerName;
        internal string SelectedMachineName;
        internal string SelectedAccountName;
        internal string SelectedDomainName;
        internal string AllExistingRobotNames = string.Empty;
        private string CPUUtilization;
        private string Memory;
        private string SignedInStatus;
        private string[] SignedInUser;
        #endregion

        #region Objects
        private string SignedInUsers = string.Empty;
        internal DataTable dt = new DataTable();
        Process cmdProcess;
        Process[] RPAAgentprocesses;
        Process[] RoboSEprocesses;
        Dictionary<string, string> SessionDict = new Dictionary<string, string>();
        Dictionary<string, string> RPAAgentprocessStatusDict = new Dictionary<string, string>();
        Dictionary<string, string> RobotNameDict = new Dictionary<string, string>();
        internal List<string> MachineList = new List<string>();
        internal List<string> AllRobotsPresent = new List<string>();
        internal List<string> AccountList = new List<string>();
        internal List<string> DomainList = new List<string>();
        System.Timers.Timer timerFindSecurityRDC = null;
        int timerFindSecurityRDCCount = 0;
        #endregion


        private RoboVMModel _RoboVMControls;
        private bool CPUCheck;

        public RoboVMModel RoboVMControls
        {
            get { return _RoboVMControls; }
            set { _RoboVMControls = value; }
        }

        public bool CanCheck
        {
            get
            {
                if (RoboVMControls.SelectedMachineName != null && RoboVMControls.SelectedMachineName != string.Empty)
                { return true; }
                else
                { return false; }
            }
        }

        public bool CanConnect
        {
            get
            {
                if (RoboVMControls.SelectedMachineName != null && RoboVMControls.SelectedMachineName != string.Empty && RoboVMControls.SelectedDomainName != null && RoboVMControls.SelectedDomainName != string.Empty && RoboVMControls.SelectedAccountName != null && RoboVMControls.SelectedAccountName != string.Empty)
                { return true; }
                else
                { return false; }
            }
        }

        #region Commands
        public ICommand ConnectCommand { get; set; }
        public ICommand CheckCommand { get; set; }
        public ICommand ReportingCommand { get; set; }
        #endregion

        public RoboVMViewModel()
        {
            if (dt.Rows.Count == 0)
            {
                dt.Clear();
                dt.Columns.Add("All Users",typeof(string));
                dt.Columns.Add("Sign In Status", typeof(string));
                dt.Columns.Add("Session Status",typeof(string));
                dt.Columns.Add("RPAAgent Status",typeof(string));
                dt.Columns.Add("Active Robot Names", typeof(string));
                dt.Columns.Add("All Robots present", typeof(string));
                dt.Columns.Add("CPU Utilization", typeof(string));
                dt.Columns.Add("Available Memory", typeof(string));
                foreach (var MachineName in MachineNames.Split(','))
                {
                    MachineList.Add(MachineName);
                }
                foreach (var AccountName in AccountNames.Split(','))
                {
                    AccountList.Add(AccountName);
                }
                foreach (var DomainName in DomainNames.Split(','))
                {
                    DomainList.Add(DomainName);
                }
                _RoboVMControls = new RoboVMModel(MachineList, AccountList, DomainList,SelectedMachineName,SelectedAccountName,SelectedDomainName,dt);
                ConnectCommand = new RoboVMConnectCommand(this);
                CheckCommand = new RoboVMCheckCommand(this);
                ReportingCommand = new RoboVMReportingCommand(this);
            }
            timerFindSecurityRDC = new System.Timers.Timer();
            timerFindSecurityRDC.AutoReset = false;
            timerFindSecurityRDC.Interval = 2000;
            timerFindSecurityRDC.Elapsed += new System.Timers.ElapsedEventHandler(timerFindSecurityRDC_Elapsed);
        }

        public void VMConnect(string SelectedVMName, string SelectedAccount, string SelectedDomain)
        {
            try
            {
                SelectedAccountName = SelectedAccount;
                SelectedDomainName = SelectedDomain;
                SelectedMachineName = SelectedVMName;
                IntPtr rdhWnd = IntPtr.Zero;
                IntPtr rdShWnd = IntPtr.Zero;
                IntPtr rdVMhWnd = IntPtr.Zero;
                // Launching mstsc window
                Process process = Process.Start("mstsc.exe");
                Thread.Sleep(1000);
                for (int loopCnt = 0; loopCnt < 20; loopCnt++)
                {
                    rdhWnd = (IntPtr)FindWindow("#32770", "Remote Desktop Connection");

                    if (rdhWnd != IntPtr.Zero)
                    {
                        break;
                    }
                    Thread.Sleep(500);
                }

                // setting machine name field as empty field.

                AutomationElement rdMachine = AutomationElement.FromHandle(rdhWnd);
                PropertyCondition rdMC1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                PropertyCondition rdMC2 = new PropertyCondition(AutomationElement.ClassNameProperty, "Edit");
                AutomationElement rdEdit = rdMachine.FindFirst(TreeScope.Descendants, new AndCondition(rdMC1, rdMC2));
                Win32.SendMessage((IntPtr)rdEdit.Current.NativeWindowHandle, Win32.WM_SETTEXT, 0, "");
                Win32.SendMessage((IntPtr)rdEdit.Current.NativeWindowHandle, Win32.WM_SETTEXT, 0, SelectedVMName);


                // Clicking Connect button 
                timerFindSecurityRDCCount = 0;
                timerFindSecurityRDC.Start();
                AutomationElement rdConnect = AutomationElement.FromHandle(rdhWnd);
                PropertyCondition rdCC3 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                PropertyCondition rdCC4 = new PropertyCondition(AutomationElement.NameProperty, "Connect");
                AutomationElement rdClickConnect = rdConnect.FindFirst(TreeScope.Descendants, new AndCondition(rdCC3, rdCC4));
                Win32.PostMessage((IntPtr)rdClickConnect.Current.NativeWindowHandle, Win32.BN_CLICKED, 0, 0);
                Thread.Sleep(1000);
                
                //Thread.Sleep(5000);                                           //uncomment if u want to close remote desktop and set the timing in thread.sleep (when too many remote desktops have to be connected and disconnected.)
                ////Close remote desktop connection
                //AutomationElement rdClose = AutomationElement.FromHandle(rdVMhWnd);
                //rdClose.SetFocus();
                //mstscprocessId = GetWindowThreadProcessId((IntPtr)rdClose.Current.NativeWindowHandle, out ProcessID);
                //foreach (Process pr in Process.GetProcesses())
                //{
                //    try
                //    {
                //        if (pr.Id != 0)
                //        {
                //            if (pr.Id.Equals((int)ProcessID))
                //            {
                //                pr.Kill();
                //            }
                //        }
                //    }
                //    catch (Exception ex)
                //    {
                //    }
                //}
            }
            catch (Exception ex)
            {
            }
        }

        private void timerFindSecurityRDC_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            try
            {
                IntPtr rdhWnd = IntPtr.Zero;
                IntPtr rdShWnd = IntPtr.Zero;
                IntPtr rdVMhWnd = IntPtr.Zero;

                //increment counter
                timerFindSecurityRDCCount++;

                //stop the timer
                timerFindSecurityRDC.Stop();
                rdShWnd = (IntPtr)FindWindow("#32770", "Windows Security");
                
                //CLicking use another account
                if (rdShWnd != IntPtr.Zero)
                {
                    AutomationElement rdSelectAccount = AutomationElement.FromHandle(rdShWnd);
                    PropertyCondition rdSAC1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.ListItem);
                    PropertyCondition rdSAC2 = new PropertyCondition(AutomationElement.NameProperty, "Use another account");
                    AutomationElement rdSelected = rdSelectAccount.FindFirst(TreeScope.Descendants, new AndCondition(rdSAC1, rdSAC2));
                    SelectionItemPattern LineItemsSelect = (SelectionItemPattern)rdSelected.GetCurrentPattern(SelectionItemPattern.Pattern);
                    LineItemsSelect.Select();

                    //entering username

                    AutomationElement rdSecurity1 = AutomationElement.FromHandle(rdShWnd);
                    PropertyCondition rdSC1 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                    PropertyCondition rdSC2 = new PropertyCondition(AutomationElement.NameProperty, "User name");
                    AutomationElement rdUsername = rdSecurity1.FindFirst(TreeScope.Descendants, new AndCondition(rdSC1, rdSC2));
                    Win32.SendMessage((IntPtr)rdUsername.Current.NativeWindowHandle, Win32.WM_SETTEXT, 0, "");
                    Win32.SendMessage((IntPtr)rdUsername.Current.NativeWindowHandle, Win32.WM_SETTEXT, 0, SelectedDomainName + "\\" + SelectedAccountName);

                    //entering password
                    Password = GetPassword(SelectedAccountName);
                    AutomationElement rdSecurity = AutomationElement.FromHandle(rdShWnd);
                    PropertyCondition rdSC3 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
                    PropertyCondition rdSC4 = new PropertyCondition(AutomationElement.NameProperty, "Password");
                    AutomationElement rdPassword = rdSecurity.FindFirst(TreeScope.Descendants, new AndCondition(rdSC3, rdSC4));
                    Win32.SendMessage((IntPtr)rdPassword.Current.NativeWindowHandle, Win32.WM_SETTEXT, 0, "");
                    Win32.SendMessage((IntPtr)rdPassword.Current.NativeWindowHandle, Win32.WM_SETTEXT, 0, Password);

                    //clicking ok after passing credentials

                    AutomationElement rdOk = AutomationElement.FromHandle(rdShWnd);
                    PropertyCondition rdOC3 = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button);
                    PropertyCondition rdOC4 = new PropertyCondition(AutomationElement.NameProperty, "OK");
                    AutomationElement rdClickOk = rdOk.FindFirst(TreeScope.Descendants, new AndCondition(rdOC3, rdOC4));
                    Win32.SendMessage((IntPtr)rdClickOk.Current.NativeWindowHandle, Win32.BN_CLICKED, 0, 0);
                }
                else
                {
                    timerFindSecurityRDC.Start();
                }
               
            }
            catch (Exception Ex)
            {
            }
        }

        public void VMDetails()
        {
            try
            {
                SelectedMachineName = _RoboVMControls.SelectedMachineName;
                SelectedAccountName = _RoboVMControls.SelectedAccountName;
                SelectedDomainName = _RoboVMControls.SelectedDomainName;
                CPUCheck = _RoboVMControls.CPUCheck;
                #region intialize objects
                dt.Clear();
                RobotNameDict.Clear();
                SessionDict.Clear();
                RPAAgentprocessStatusDict.Clear();
                SignedInUsers = string.Empty;
                RobotName = string.Empty;
                newOwnerNameValue = string.Empty;
                AllExistingRobotNames = string.Empty;
                #endregion

                RoboSEprocesses = Process.GetProcessesByName(processName_RoboView, SelectedMachineName);
                RPAAgentprocesses = Process.GetProcessesByName(processName_RPAAgent, SelectedMachineName);

                ExtractSessionStatus();         //Sets SignedIn Users and Session Status
                
                GetRPAAgentStatusandRobotNames();

                if (SignedInUsers == "")
                {
                    SignedInUser = new string[] { "" };
                }
                else
                {
                    SignedInUser = SignedInUsers.Split(' ');
                }
                for (int j = 0; j < AccountList.Count; j++)
                { 
                    if (SignedInUsers.ToLower().Contains(AccountList[j].ToLower()) && SignedInUsers != "")
                    {
                        SignedInStatus = "Yes";
                        AllExistingRobotNames = string.Empty;
                        GetExistingRobotNames(AccountList[j]);
                        GetPerformanceData(SelectedMachineName);

                        if (RPAAgentprocessStatusDict.ContainsKey(AccountList[j]) && RobotNameDict.ContainsKey(AccountList[j]) && SessionDict.ContainsKey(AccountList[j]))
                        {
                            dt.Rows.Add(AccountList[j], SignedInStatus, SessionDict[AccountList[j]], RPAAgentprocessStatusDict[AccountList[j]], RobotNameDict[AccountList[j]], AllExistingRobotNames, CPUUtilization, Memory);
                        }
                        else if (RPAAgentprocessStatusDict.ContainsKey(AccountList[j]) && !RobotNameDict.ContainsKey(AccountList[j]) && SessionDict.ContainsKey(AccountList[j]))
                        {
                            dt.Rows.Add(AccountList[j], SignedInStatus, SessionDict[AccountList[j]], RPAAgentprocessStatusDict[AccountList[j]], "No Robots were found running", AllExistingRobotNames, CPUUtilization, Memory);
                        }
                        else if (!RPAAgentprocessStatusDict.ContainsKey(AccountList[j]) && SessionDict.ContainsKey(AccountList[j]))
                        {
                            dt.Rows.Add(AccountList[j], SignedInStatus, SessionDict[AccountList[j]], "Not Running", "No Robots were found running", AllExistingRobotNames, CPUUtilization, Memory);
                        }
                        else if((RPAAgentprocesses.Length == 0))
                        {
                            dt.Rows.Add(AccountList[j], SignedInStatus, SessionDict[AccountList[j]], "Not Found", "No Robots were found running", AllExistingRobotNames, CPUUtilization, Memory);
                        }
                        else
                        {
                            dt.Rows.Add(AccountList[j], SignedInStatus, SessionDict[AccountList[j]], "Not Found", "No Robots were found running", AllExistingRobotNames, CPUUtilization, Memory);
                        }
                    }
                    else
                    {
                        SignedInStatus = "No";
                        GetExistingRobotNames(AccountList[j]);
                        GetPerformanceData(SelectedMachineName);
                        if (AllExistingRobotNames != "Robots might be placed on different directory.")
                        dt.Rows.Add(AccountList[j], SignedInStatus, "Not Logged In", "Not Found","No robots running", AllExistingRobotNames, CPUUtilization, Memory);
                    }
                }
                if (dt.Rows.Count == 0)
                {
                    dt.Rows.Add("No user connected", SignedInStatus, "Not Logged In", "Not Found", "No robots running", AllExistingRobotNames, CPUUtilization, Memory);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("User is not registered on that machine /r/n or "+ ex.Message);
            }
        }

        private void GetRPAAgentStatusandRobotNames()
        {
            ConnectionOptions options = new ConnectionOptions();
            options.Username = "jsonawk";
            options.Password = jsonawk_password;
            options.EnablePrivileges = true;
            options.Authority = "ntlmdomain:jci.com";
            ManagementScope scope = new ManagementScope("\\\\" + SelectedMachineName + "\\root\\cimv2", options);

            foreach (var process in RPAAgentprocesses)
            {
                if (process.ProcessName == "RPAAgentView")
                {
                    RPAAgentprocessID = process.Id;
                    GetRPAandSEProcessDetails(process, SelectedMachineName, Password,scope);
                    if (RPAAgentprocessID != 0)
                    {
                        RPAAgentprocessStatusDict.Add(RPAAgentOwnerName, "Running");
                    }
                    else
                    {
                        RPAAgentprocessStatusDict.Add(RPAAgentOwnerName, "Not Running");
                    }
                }
            }
            foreach (var process in RoboSEprocesses)
            {
                if (process.ProcessName == "RoboSE")
                {
                    RoboSEprocessID = process.Id;
                    RoboSEprocessStatus = RoboSEprocessID == 0 ? "Not Running" : "Running";
                    GetRPAandSEProcessDetails(process, SelectedMachineName, Password,scope);
                    var temp = RobotSEPath;
                    var searchName = temp.Split('\\');
                    var tempName = searchName[searchName.Length - 2];
                    RobotName += tempName + ",";
                    if (RobotNameDict.ContainsKey(RoboSEOwnerName))
                    {
                        var tempOwnerValue = RobotNameDict[RoboSEOwnerName];
                        newOwnerNameValue = tempOwnerValue + "," + tempName;
                        RobotNameDict.Remove(RoboSEOwnerName);
                        RobotNameDict.Add(RoboSEOwnerName, newOwnerNameValue);
                    }
                    else
                    {
                        RobotNameDict.Add(RoboSEOwnerName, tempName);
                    }
                }
            }
        }
        
        private void ExtractSessionStatus()
        {
            cmdProcess = Process.Start("cmd.exe", "/c qwinsta /server:" + SelectedMachineName + ">SessionReport.txt");
            Thread.Sleep(500);
            lines = File.ReadAllLines(Environment.CurrentDirectory + "/SessionReport.txt");
            foreach (string line in lines)
            {
                line.Replace(' ', ',');
                AccountKeywords = line.Split(' ');
                for (int i = 0; i < AccountKeywords.Length; i++)
                {
                    if (AccountKeywords[i] == "A4201002" || AccountKeywords[i] == "A3000102" || AccountKeywords[i] == "A3000156" || AccountKeywords[i] == "A3051025" || AccountKeywords[i] == "A3000096" || AccountKeywords[i].ToLower() == "jsonawk" || AccountKeywords[i].ToLower() == "jbhagan" || AccountKeywords[i].ToLower() == "jpalanj" || AccountKeywords[i].ToLower() == "jnaths" || AccountKeywords[i].ToLower() == "jshankgo")
                    {
                        SessionUserName = AccountKeywords[i].ToString();
                        SignedInUsers += AccountKeywords[i].ToString() + " ";
                        for (int j = 0; j < AccountKeywords.Length; j++)
                        {
                            if (sessionStatuses.Contains(AccountKeywords[j].ToString()) && AccountKeywords[j] != string.Empty)
                            {
                                SessionValues = AccountKeywords[j].ToString();
                                break;
                            }
                        }
                        if (SessionUserName != null && SessionValues != null)
                        {
                            SessionDict.Add(SessionUserName, SessionValues);
                        }
                    }
                }
            }
            if (SignedInUsers!="")
            {
                if (SignedInUsers[SignedInUsers.Length - 1] == ',')
                {
                    SignedInUsers = SignedInUsers.Substring(0, SignedInUsers.Length - 1);
                }
            }
        }

        private void GetExistingRobotNames(string SignedInUser)
        {
            AllExistingRobotNames = string.Empty;
            if (File.Exists("\\\\" + SelectedMachineName + "\\" + SEPath + SignedInUser + ConfigPath))
            {
                ConfigFile = Directory.GetFiles("\\\\" + SelectedMachineName + "\\" + SEPath + SignedInUser + "\\Configurations\\");
                foreach (string item in ConfigFile)
                {
                    string[] temp = item.Split('\\');
                    if (temp[temp.Length - 1] == ConfigFileName)
                    {
                        Regex r = new Regex(@"\w+(?=\</Name>)"); //@"\b(?<word>\w+)\s+(\k<word>)\b"
                        string text = File.ReadAllText(item);
                        MatchCollection mc = r.Matches(text);               //Extracts the robot names from config.xml
                        foreach (var robotname in mc)
                        {
                            if (mc.Count == 1)
                            {
                                AllExistingRobotNames += robotname;
                            }
                            else
                            {
                                AllExistingRobotNames += robotname + ",";
                            }
                        }
                    }
                }
                if (AllExistingRobotNames != "")
                {
                    if (AllExistingRobotNames[AllExistingRobotNames.Length - 1] == ',')
                    {
                        AllExistingRobotNames = AllExistingRobotNames.Substring(0, AllExistingRobotNames.Length - 1);
                    }
                }
                else
                {
                    AllExistingRobotNames = "Robots are yet to be configured.";
                }
            }
            else
            {
                AllExistingRobotNames = "Robots might be placed on different directory.";
            }
        }

        private void GetPerformanceData(string SelectedVMName)
        {
            PerformanceCounter cpuCounter;
            cpuCounter = new PerformanceCounter();
            cpuCounter.MachineName = SelectedVMName;
            cpuCounter.CategoryName = "Processor";
            cpuCounter.CounterName = "% Processor Time";
            cpuCounter.InstanceName = "_Total";
            cpuCounter.NextValue();
            if (CPUCheck == true)
            {
                Thread.Sleep(1000);
            }
            CPUUtilization = cpuCounter.NextValue() + "%";
            PerformanceCounter MemoryCounter;
            MemoryCounter = new PerformanceCounter();
            MemoryCounter.MachineName = SelectedVMName;
            MemoryCounter.CategoryName = "Memory";
            MemoryCounter.CounterName = "Available MBytes";
            Memory = MemoryCounter.NextValue() + "MB";
        }

        public void GetRPAandSEProcessDetails(Process process, string SelectedVMName, string Password,ManagementScope scope )
        {
            try
            {
                var RPAAgentquery = new SelectQuery("Select * From Win32_Process Where ProcessID = " + RPAAgentprocessID);
                var RoboSEquery = new SelectQuery("Select * From Win32_Process Where ProcessID = " + RoboSEprocessID);
                scope.Connect();
                ManagementObjectSearcher RoboSEsearcher = new ManagementObjectSearcher(scope, RoboSEquery);
                ManagementObjectCollection RoboSEprocessList = RoboSEsearcher.Get();
                foreach (ManagementObject item in RoboSEprocessList)
                {
                    object path = item["ExecutablePath"];
                    if (path != null)
                    {
                        RobotSEPath = path.ToString();
                    }
                    string[] argList = new string[] { string.Empty, string.Empty };
                    int returnVal = Convert.ToInt32(item.InvokeMethod("GetOwner", argList));
                    RoboSEOwnerName = argList[0];
                }
                ManagementObjectSearcher RPAAgentsearcher = new ManagementObjectSearcher(scope, RPAAgentquery);
                ManagementObjectCollection RPAAgentprocessList = RPAAgentsearcher.Get();
                foreach (ManagementObject item in RPAAgentprocessList)
                {
                    string[] argList = new string[] { string.Empty, string.Empty };
                    int returnVal = Convert.ToInt32(item.InvokeMethod("GetOwner", argList));
                    RPAAgentOwnerName = argList[0];
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show("Error in GetRoboSEProcessDetails method");
            }
        }

        //public void GetRPAAgentProcessDetails(Process process, string SelectedVMName, string Password )
        //{
        //    try
        //    {
        //        var RPAAgentquery = new SelectQuery("Select * From Win32_Process Where ProcessID = " + RPAAgentprocessID);
        //        ConnectionOptions options = new ConnectionOptions();
        //        options.Username = "jsonawk";
        //        options.Password = jsonawk_password;
        //        options.EnablePrivileges = true;
        //        options.Authority = "ntlmdomain:jci.com";
        //        ManagementScope scope = new ManagementScope("\\\\" + SelectedVMName + "\\root\\cimv2", options);
        //        scope.Connect();
        //        ManagementObjectSearcher RPAAgentsearcher = new ManagementObjectSearcher(scope, RPAAgentquery);
        //        ManagementObjectCollection RPAAgentprocessList = RPAAgentsearcher.Get();
        //        foreach (ManagementObject item in RPAAgentprocessList)
        //        {
        //            string[] argList = new string[] { string.Empty, string.Empty };
        //            int returnVal = Convert.ToInt32(item.InvokeMethod("GetOwner", argList));
        //            RPAAgentOwnerName = argList[0];
        //            Thread.Sleep(500);
        //        }
        //    }
        //    catch (Exception Ex)
        //    {
        //        MessageBox.Show("Error in GetRPAAgentProcessDetails method");
        //    }
        //}

        public string GetPassword(string SelectedAccountName)
        {
            try
            {
                foreach(string accountname in AccountList)
                {
                    if (accountname == SelectedAccountName)
                    {
                        Password = ConfigurationManager.AppSettings[SelectedAccountName + "_password"].ToString();
                        return Password;
                    }
                }
                return string.Empty;
            }
            catch
            {
                MessageBox.Show("Password not listed in app config file.\r\nPlease enter password according to account belonging to that particular machine.");
                return string.Empty;
            }
        }

    }
}
