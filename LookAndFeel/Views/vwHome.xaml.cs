using LookAndFeel.Conexiones;
using LookAndFeel.ViewModels;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using WpfAnimatedGif;

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
            if (!CDF)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/CDF.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(imgCDF, image);
                CDF = true;
            }
        }

        private void imgCDF_MouseLeave(object sender, MouseEventArgs e)
        {
            if (CDF)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/CDF.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(imgCDF, image);
                CDF = false;
            }
        }
        bool ched = false;
        private void imgChedraui_MouseLeave(object sender, MouseEventArgs e)
        {
            if (ched)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Chedraui.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(imgChedraui, image);
                ched = false;
            }
        }

        private void imgChedraui_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!ched)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Chedraui.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(imgChedraui, image);
                ched = true;
            }
        }
        bool wal = false;
        private void Walmart_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!wal)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Walmart.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Walmart, image);
                wal = true;
            }
        }

        private void Walmart_MouseLeave(object sender, MouseEventArgs e)
        {
            if (wal)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Wallmart.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Walmart, image);
                wal = false;
            }
        }
        bool sori = false;
        private void Soriana_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!sori)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Soriana-Comex.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Soriana, image);
                sori = true;
            }
        }

        private void Soriana_MouseLeave(object sender, MouseEventArgs e)
        {
            if (sori)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Soriana.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Soriana, image);
                sori = false;
            }
        }
        bool oxx = false;
        private void oxxo_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!oxx)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Oxxo.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(oxxo, image);
                oxx = true;
            }
        }

        private void oxxo_MouseLeave(object sender, MouseEventArgs e)
        {
            if (oxx)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Oxxo.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(oxxo, image);
                oxx = false;
            }
        }
        bool he = false;
        private void imgHeb_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!he)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/HEB.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(imgHeb, image);
                he = true;
            }
        }

        private void imgHeb_MouseLeave(object sender, MouseEventArgs e)
        {
            if (he)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/HEB.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(imgHeb, image);
                he = false;
            }
        }
        bool fres = false;
        private void ImgFresko_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!fres)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Fresko.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(ImgFresko, image);
                fres = true;
            }
        }

        private void ImgFresko_MouseLeave(object sender, MouseEventArgs e)
        {
            if (fres)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Fresko.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(ImgFresko, image);
                fres = false;
            }
        }
        bool Frag = false;
        private void Fragua_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!Frag)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Fragua.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Fragua, image);
                Frag = true;
            }
        }

        private void Fragua_MouseLeave(object sender, MouseEventArgs e)
        {
            if (Frag)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Fragua.jpg");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Fragua, image);
                Frag = false;
            }
        }

        bool Cos = false;
        private void Costco_MouseEnter(object sender, MouseEventArgs e)
        {
            if (!Cos)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Gif/Costco.gif");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Costco, image);
                Cos = true;
            }
        }

        private void Costco_MouseLeave(object sender, MouseEventArgs e)
        {
            if (Cos)
            {
                var image = new BitmapImage();
                image.BeginInit();
                image.UriSource = new Uri("pack://application:,,,/Picture/Costco.png");
                image.EndInit();
                ImageBehavior.SetAnimatedSource(Costco, image);
                Cos = false;
            }
        }
    }
}
