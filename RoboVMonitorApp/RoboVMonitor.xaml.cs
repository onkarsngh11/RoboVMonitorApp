using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Management;
using System.Management.Instrumentation;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using RoboVMonitorApp.ViewModels;

namespace RoboVMonitorApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary> 
    public partial class RoboVMonitor : Window
    {
        public RoboVMonitor()
        {
            InitializeComponent();
            DataContext = new RoboVMViewModel();
            this.Closed += OnClosed;
        }
        private void OnClosed(object sender, EventArgs e)
        {
            
            Environment.Exit(0);
        }
    }
}
