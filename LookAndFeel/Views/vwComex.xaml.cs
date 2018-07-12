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

namespace LookAndFeel.Views
{
    /// <summary>
    /// Interaction logic for vwComex.xaml
    /// </summary>
    public partial class vwComex : Window
    {
        public vwComex()
        {
            InitializeComponent();
            this.DataContext = new ComexViewModel();
        }

        private void header_MouseDown(object sender, MouseButtonEventArgs e)
        { try { this.DragMove(); } catch { } }
    }
}
