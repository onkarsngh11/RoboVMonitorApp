using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp.Commands
{
    class RoboVMConnectCommand: ICommand
    {
        private RoboVMViewModel _viewModel;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public bool CanExecute(object parameter)
        {
            return _viewModel.CanConnect;
        }

        public void Execute(object parameter)
        {
            _viewModel.VMConnect(_viewModel.RoboVMControls.SelectedMachineName, _viewModel.RoboVMControls.SelectedAccountName, _viewModel.RoboVMControls.SelectedDomainName);
        }
        public RoboVMConnectCommand(RoboVMViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
