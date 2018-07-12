using GalaSoft.MvvmLight.Command;
using LookAndFeel.Conexiones;
using LookAndFeel.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace LookAndFeel.ViewModels
{
    class FraguaViewModel
    {
        Credenciales creden;
        conds condb = new conds();

        String Tabla = "TblFragua";
        public FraguaViewModel()
        {
            creden = new Credenciales
            {
                Usuario = condb.ds.Tables[Tabla].Rows[0][0].ToString(),
                Contrasenia = condb.ds.Tables[Tabla].Rows[0][1].ToString()
            };
            CreateCancelCommand();
            CreateGuardarCommand();
            CreateBeginProcessCommand();
        }

        public String usuario
        {
            get { return creden.Usuario; }
            set { creden.Usuario = value; }
        }

        public String contrasenia
        {
            get { return creden.Contrasenia; }
            set { creden.Contrasenia = value; }
        }

        DateTime FechaI = DateTime.Now;
        public DateTime FechaInicial
        {
            get { return FechaI; }
            set { FechaI = value; }
        }

        DateTime FechaF = DateTime.Now;
        public DateTime FechaFinal
        {
            get { return FechaF; }
            set { FechaF = value; }
        }

        private WindowState StateWindows = WindowState.Normal;
        public WindowState win
        {
            get { return StateWindows; }
            set { StateWindows = value; }
        }

        #region Command Save
        public ICommand GuardarCommand
        {
            get; internal set;
        }

        private bool CanExecuteGuardarCommand()
        {
            bool taco = !string.IsNullOrEmpty(usuario) || !string.IsNullOrEmpty(contrasenia);
            return true;
            // return false;
        }

        private void CreateGuardarCommand()
        {
            GuardarCommand = new RelayCommand(GuardarRegistros, CanExecuteGuardarCommand);
        }

        public void GuardarRegistros()
        {
            bool Save;
            if (String.IsNullOrEmpty(usuario) || String.IsNullOrEmpty(contrasenia))
            { Save = false; MessageBox.Show("Revise que el campo de usuario o contraseña se encuentren vacios", "AVISO", MessageBoxButton.OK, MessageBoxImage.Information); }
            else
            {
                try
                {
                    condb.ds.Tables[Tabla].Rows[0][0] = creden.Usuario;
                    condb.ds.Tables[Tabla].Rows[0][1] = creden.Contrasenia;
                    condb.SetInfo();
                    Save = true;
                    MessageBox.Show("La información se guardo correctamente", "AVISO", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (System.Exception ex) { Save = false; }
            }
        }
        #endregion

        #region Command Begin
        public ICommand BeginProcessCommand
        {
            get; internal set;
        }

        private bool CanExecuteBeginProcessCommand()
        {
            return true;
        }

        private void CreateBeginProcessCommand()
        {
            BeginProcessCommand = new RelayCommand(BeginProcess, CanExecuteBeginProcessCommand);
        }

        public void BeginProcess()
        {
            Pruebas_clase7.Clases.Fragua Proceso = new Pruebas_clase7.Clases.Fragua();
            try
            {
                Proceso.funcionPrincipal(usuario, contrasenia);
                MessageBox.Show("Proceso finalizado");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ocurrio un problema,\n" + ex.Message);
            }
        }
        #endregion

        #region Command Cancel
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
    }
}
