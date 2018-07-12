using GalaSoft.MvvmLight.Command;
using LookAndFeel.Vistas;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

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

        }
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

        private void CreateComexCommand()
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

    }
}
