using RoboVMonitorApp.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace RoboVMonitorApp.Commands
{
    public class ReportingExtractCommand : ICommand
    {
        private ReportingViewModel _viewModel;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
        
        public bool CanExecute(object parameter)
        {
            return _viewModel.CanExtract;
        }

        public void Execute(object parameter)
        {
            _viewModel.ExtractData();
        }
        public ReportingExtractCommand(ReportingViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
