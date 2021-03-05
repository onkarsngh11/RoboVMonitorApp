using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp.Models
{
    public class ReportingModel : INotifyPropertyChanged
    {        
        private List<string> _UCList;
        public List<string> UCList
        {
            get { return _UCList; }
            set { _UCList = value;
                OnPropertyChanged("UCList");
            }
        }

        private DateTime fromDate=DateTime.Now.Date;
        public DateTime FromDate
        {
            get { return fromDate; }
            set { fromDate = value;}
        }
        private DateTime toDate=DateTime.Now.Date;
        public DateTime ToDate
        {
            get { return toDate; }
            set { toDate = value; }
        }

        private string ucName;
        public string UCName
        {
            get { return ucName; }
            set { ucName = value; OnPropertyChanged("UCName"); }
        }

        private DataTable _UCDetails;
        public DataTable UCDetails
        {
            get { return _UCDetails; }
            set { _UCDetails = value; OnPropertyChanged("UCDetails");} 
        }

        private bool fullDateRange;
        public bool FullDateRange
        {
            get { return fullDateRange; }
            set { fullDateRange = value; }
        }

        private bool chkSpTracker;
        public bool ChkSpTracker
        {
            get { return chkSpTracker; }
            set { chkSpTracker = value; }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public ReportingModel(List<string> UCNamesList,string SelectedUCName)
        {
            UCList = UCNamesList;
            UCName = SelectedUCName;
        }
    }
}
