using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp.Commands
{
    class ReportingUpdateCommand : ICommand
    {
        private ReportingViewModel _viewModel;

        public event EventHandler CanExecuteChanged;

        public bool CanExecute(object parameter)
        {
            return _viewModel.CanUpdate;
        }

        public void Execute(object parameter)
        {
            _viewModel.UpdateTracker();
        }
        public ReportingUpdateCommand(ReportingViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
