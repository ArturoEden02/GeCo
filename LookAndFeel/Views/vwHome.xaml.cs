using LookAndFeel.ViewModels;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Effects;
using System.Windows.Media;

namespace LookAndFeel.Vistas
{
    /// <summary>
    /// Lógica de interacción para vwHome.xaml
    /// </summary>
    public partial class vwHome : Window
    {
        public vwHome()
        {
            InitializeComponent();
            DataContext = new HomeViewModel();
        }

        private void header_MouseDown(object sender, MouseButtonEventArgs e)
        { try { this.DragMove(); } catch { } }

        #region CDF
        private void MouseEnter(object sender, MouseEventArgs e)
        {
            Image myButton = sender as Image;
            DropShadowBitmapEffect mydrop = new DropShadowBitmapEffect();
            Color color = new Color();
            color.ScR = 0;
            color.ScG = 0;
            color.ScB = 0;
            mydrop.Color = color;
            mydrop.Direction = 200;
            mydrop.ShadowDepth = 20;
            mydrop.Softness = 10;
            mydrop.Opacity = 1;
            myButton.BitmapEffect = mydrop;
        }

        private void MouseLeave(object sender, MouseEventArgs e)
        {
            Image myButton = sender as Image;
            DropShadowBitmapEffect mydrop = new DropShadowBitmapEffect();
            Color color = new Color();
            color.ScA = 1;
            color.ScR = 255;
            color.ScG = 255;
            color.ScB = 255;
            mydrop.Color = color;
            mydrop.Direction = 200;
            mydrop.ShadowDepth = 20;
            mydrop.Softness = 10;
            mydrop.Opacity = .5;
            myButton.BitmapEffect = mydrop;
        }
        #endregion

        #region Boton Cancelar
        private void btnCancelar_MouseEnter(object sender, MouseEventArgs e)
        {
            SolidColorBrush mySolidColorBrush = new SolidColorBrush();
            mySolidColorBrush = (SolidColorBrush)(new BrushConverter().ConvertFrom("#E81123"));
            Button btn = sender as Button;
            btn.Background = mySolidColorBrush;
            btn.BorderBrush = mySolidColorBrush;
        }

        private void btnCancelar_MouseLeave(object sender, MouseEventArgs e)
        {
            SolidColorBrush mySolidColorBrush = new SolidColorBrush();
            mySolidColorBrush = (SolidColorBrush)(new BrushConverter().ConvertFrom("#004790"));
            Button btn = sender as Button;
            btn.Background = mySolidColorBrush;
            btn.BorderBrush = mySolidColorBrush;
        }
        #endregion

    }
}