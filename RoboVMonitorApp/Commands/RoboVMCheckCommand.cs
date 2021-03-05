using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp.Commands
{
    class RoboVMCheckCommand: ICommand
    {
        private RoboVMViewModel _viewModel;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public RoboVMViewModel reportingViewModel { get; set; }

        public bool CanExecute(object parameter)
        {
            return _viewModel.CanCheck;
        }

        public void Execute(object parameter)
        {
            _viewModel.VMDetails();
        }
        public RoboVMCheckCommand(RoboVMViewModel viewModel)
        {
            _viewModel = viewModel;
        }
    }
}
