using LookAndFeel.Conexiones;
using LookAndFeel.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace LookAndFeel.Vistas
{
    /// <summary>
    /// Lógica de interacción para vwSoriana.xaml
    /// </summary>
    public partial class vwSoriana : Window
    {
        public vwSoriana()
        {
            InitializeComponent();
            this.DataContext = new SorianaViewModel();
        }
        private void header_MouseDown(object sender, MouseButtonEventArgs e)
        { try { this.DragMove(); } catch { } }
    }
}
