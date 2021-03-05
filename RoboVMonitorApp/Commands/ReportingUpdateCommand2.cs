using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp.Commands
{
    class ReportingUpdateCommand2:ICommand
    {
        private ReportingViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return _viewModel.CanUpdate2;
        }

        public void Execute(object parameter)
        {
            _viewModel.UpdateTracker2();
        }
        public ReportingUpdateCommand2(ReportingViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
