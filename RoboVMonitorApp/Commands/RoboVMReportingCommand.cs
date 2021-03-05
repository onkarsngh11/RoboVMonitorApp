using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp.Commands
{
    class RoboVMReportingCommand:ICommand
    {
        private RoboVMViewModel _viewModel;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            ReportingSection rs = new ReportingSection();
            rs.Show();
        }
        public RoboVMReportingCommand(RoboVMViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
