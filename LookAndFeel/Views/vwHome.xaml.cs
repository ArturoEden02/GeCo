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

        bool CDF = false;
        private void imgCDF_MouseEnter(object sender, MouseEventArgs e)
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
        private void imgCDF_MouseLeave(object sender, MouseEventArgs e)
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

        bool ched = false;
        private void imgChedraui_MouseLeave(object sender, MouseEventArgs e)
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
        private void imgChedraui_MouseEnter(object sender, MouseEventArgs e)
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

        bool wal = false;
        private void Walmart_MouseEnter(object sender, MouseEventArgs e)
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
        private void Walmart_MouseLeave(object sender, MouseEventArgs e)
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

        bool sori = false;
        private void Soriana_MouseEnter(object sender, MouseEventArgs e)
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
        private void Soriana_MouseLeave(object sender, MouseEventArgs e)
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

        bool oxx = false;
        private void oxxo_MouseEnter(object sender, MouseEventArgs e)
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
        private void oxxo_MouseLeave(object sender, MouseEventArgs e)
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

        bool he = false;
        private void imgHeb_MouseEnter(object sender, MouseEventArgs e)
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
        private void imgHeb_MouseLeave(object sender, MouseEventArgs e)
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

        bool fres = false;
        private void ImgFresko_MouseEnter(object sender, MouseEventArgs e)
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
        private void ImgFresko_MouseLeave(object sender, MouseEventArgs e)
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

        bool Frag = false;
        private void Fragua_MouseEnter(object sender, MouseEventArgs e)
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
        private void Fragua_MouseLeave(object sender, MouseEventArgs e)
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

        bool Cos = false;
        private void Costco_MouseEnter(object sender, MouseEventArgs e)
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
        private void Costco_MouseLeave(object sender, MouseEventArgs e)
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

        bool Com = false;
        private void Comex_MouseEnter(object sender, MouseEventArgs e)
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
        private void Comex_MouseLeave(object sender, MouseEventArgs e)
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
    }
}