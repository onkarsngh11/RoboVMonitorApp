using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RoboVMonitorApp.Models
{
    public class RoboVMModel : INotifyPropertyChanged
    {
        private List<string> _MachinesList;
        public List<string> MachinesList
        {
            get { return _MachinesList; }
            set
            {
                _MachinesList = value;
            }
        }
        private List<string> _AccountsList;

        public List<string> AccountsList
        {
            get { return _AccountsList; }
            set { _AccountsList = value; }
        }
        private List<string> _DomainList;

        public List<string> DomainsList
        {
            get { return _DomainList; }
            set { _DomainList = value; }
        }
        private DataTable _VMDetails;

        public DataTable VMDetails
        {
            get { return _VMDetails; }
            set { _VMDetails = value; }
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public RoboVMModel(List<string> MachineList,List<string> AccountList,List<string> DomainList,string MachineName,string AccountName,string DomainName,DataTable dt)
        {
            MachinesList = MachineList;
            AccountsList = AccountList;
            DomainsList = DomainList;
            SelectedMachineName = MachineName;
            SelectedAccountName = AccountName;
            SelectedDomainName = DomainName;
            VMDetails = dt;
        }
        private string _SelectedMachineName;

        public string SelectedMachineName
        {
            get { return _SelectedMachineName; }
            set { _SelectedMachineName = value; }
        }
        private string _SelectedAccountName;

        public string SelectedAccountName
        {
            get { return _SelectedAccountName; }
            set { _SelectedAccountName = value; }
        }
        private string _SelectedDomainName;

        public string SelectedDomainName
        {
            get { return _SelectedDomainName; }
            set { _SelectedDomainName = value; }
        }

        private bool _CPUCheck;

        public bool CPUCheck
        {
            get { return _CPUCheck; }
            set { _CPUCheck = value; }
        }

    }
}
