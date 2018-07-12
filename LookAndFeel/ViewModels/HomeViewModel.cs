using GalaSoft.MvvmLight.Command;
using LookAndFeel.Views;
using LookAndFeel.Vistas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows;

namespace LookAndFeel.ViewModels
{
    class HomeViewModel
    {
        public HomeViewModel()
        {
            CreateComexCommand();
            CreateComeXSorianaCommand();
            CreateFraguaCommand();
            CreateCDFCommand();
            CreateChedraguiCommand();
            CreateHEBCommand();
            CreateOxxoCommand();
            CreateWalmartCommand();
            CreateCostcoCommand();
            CreateFreskoCommand();
            CreateCancelCommand();
        }

        #region Commands

        #region Command Chedraui
        public ICommand Chedraui
        { get; internal set; }

        private bool CanExecuteChedraguiCommand()
        { return true; }

        private void CreateChedraguiCommand()
        { Chedraui = new RelayCommand(ChedraguiProcess, CanExecuteChedraguiCommand); }

        public void ChedraguiProcess()
        { vwChedragui che = new vwChedragui(); che.ShowDialog(); }
        #endregion

        #region Command Fragua
        public ICommand Fragua
        { get; internal set; }

        private bool CanExecuteGuardarCommand()
        { return true; }

        private void CreateFraguaCommand()
        { Fragua = new RelayCommand(FraguaProcess, CanExecuteGuardarCommand); }

        public void FraguaProcess()
        { vwFragua che = new vwFragua(); che.ShowDialog(); }
        #endregion

        #region Command Fresko
        public ICommand Fresko
        { get; internal set; }

        private bool CanExecuteFreskoCommand()
        { return true; }

        private void CreateFreskoCommand()
        { Fresko = new RelayCommand(FreskoProcess, CanExecuteFreskoCommand); }

        public void FreskoProcess()
        { vwFresko che = new vwFresko(); che.ShowDialog(); }
        #endregion

        #region Command CDF
        public ICommand CDF
        { get; internal set; }

        private bool CanExecuteCDFCommand()
        { return true; }

        private void CreateCDFCommand()
        { CDF = new RelayCommand(CDFProcess, CanExecuteCDFCommand); }

        public void CDFProcess()
        { vwCDF che = new vwCDF(); che.ShowDialog(); }
        #endregion #region Command Fragua

        #region Command ComeXSoriana
        public ICommand ComeXSoriana
        { get; internal set; }

        private bool CanExecuteComeXSorianaCommand()
        { return true; }

        private void CreateComeXSorianaCommand()
        { ComeXSoriana = new RelayCommand(ComeXSorianaProcess, CanExecuteComeXSorianaCommand); }

        public void ComeXSorianaProcess()
        { vwSoriana che = new vwSoriana(); che.ShowDialog(); }
        #endregion

        #region Command HEB
        public ICommand HEB
        { get; internal set; }

        private bool CanExecuteHEBCommand()
        { return true; }

        private void CreateHEBCommand()
        { HEB = new RelayCommand(HEBProcess, CanExecuteHEBCommand); }

        public void HEBProcess()
        { vwHEB che = new vwHEB(); che.ShowDialog(); }
        #endregion

        #region Command Oxxo
        public ICommand Oxxo
        { get; internal set; }

        private bool CanExecuteOxxoCommand()
        { return true; }

        private void CreateOxxoCommand()
        { Oxxo = new RelayCommand(OxxoProcess, CanExecuteOxxoCommand); }

        public void OxxoProcess()
        { vwOXXO che = new vwOXXO(); che.ShowDialog(); }
        #endregion

        #region Command Walmart
        public ICommand Walmart
        { get; internal set; }

        private bool CanExecuteWalmartCommand()
        { return true; }

        private void CreateWalmartCommand()
        { Walmart = new RelayCommand(WalmartProcess, CanExecuteWalmartCommand); }

        public void WalmartProcess()
        { vwWalmart che = new vwWalmart(); che.ShowDialog(); }
        #endregion

        #region Command Costco
        public ICommand Costco
        { get; internal set; }

        private bool CanExecuteCostcoCommand()
        { return true; }

        private void CreateCostcoCommand()
        { Costco = new RelayCommand(CostcoProcess, CanExecuteCostcoCommand); }

        public void CostcoProcess()
        { vwCostco che = new vwCostco(); che.ShowDialog(); }
        #endregion

        #region Command Comex
        public ICommand Comex
        { get; internal set; }

        private bool CanExecuteComexCommand()
        { return true; }

        private void CreateComexCommand()
        { Comex = new RelayCommand(ComexProcess, CanExecuteComexCommand); }

        public void ComexProcess()
        { vwComex che = new vwComex(); che.ShowDialog(); }
        #endregion

        #region Command Salir

        public ICommand CancelCommand
        {
            get; internal set;
        }

        private bool CanExecuteCancelCommand()
        {
            return true;
        }

        private void CreateCancelCommand()
        {
            CancelCommand = new RelayCommand(CancelarCommand, CanExecuteCancelCommand);
        }

        public void CancelarCommand()
        {
            Application.Current.Windows.OfType<Window>().SingleOrDefault(w => w.IsActive).Close();
        }
        #endregion

        #endregion

    }
}
