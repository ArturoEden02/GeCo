namespace Pruebas_clase7.Clases
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Excel = Microsoft.Office.Interop.Excel;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using System.IO;
    using System.Threading;
    using System.Collections.ObjectModel;
    using System.Windows;

    class Chedraui
    {
        IWebDriver driver;
        String rutaDescargas = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
        String rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        public bool FuncionPrincipal(string usuario, string contraseña, DateTime FI, DateTime FF)
        {
            bool exito = true;
            string[] rutaArchivoRetenidos = new string[0];
            string[] rutaArchivoSaldo = new string[0];
            string[] rutaArchivoHistorico = new string[0];
            DateTime Inicio = DateTime.Now;
            if (FI > FF.AddDays(-90))
            {

                ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
                driverService.HideCommandPromptWindow = true;
                driver = new ChromeDriver(driverService, new ChromeOptions());
                driver.Navigate().GoToUrl("https://portal-financiero.chedraui.com.mx/#/account/login");
                bool bandera = false;
                do
                {
                    try
                    {
                        driver.FindElement(By.Name("username"));
                        bandera = false;
                    }
                    catch
                    {
                        bandera = true;
                        Thread.Sleep(500);
                    }
                } while (bandera != false);
                bandera = false;
                driver.FindElement(By.Name("username")).Clear();
                driver.FindElement(By.Name("username")).SendKeys(usuario);
                driver.FindElement(By.Name("password")).Clear();
                driver.FindElement(By.Name("password")).SendKeys(contraseña);
                Thread.Sleep(1000);
                ReadOnlyCollection<IWebElement> buttons = driver.FindElements(By.TagName("button"));
                for (int a = 0; a < buttons.Count; a++)
                {
                    string texto = buttons[a].Text;
                    if (texto == "INICIAR SESIÓN")
                    {
                        buttons[a].Click();
                        a = buttons.Count;
                    }
                }
                do
                {
                    try
                    {
                        driver.FindElement(By.Id("side-menu"));
                        bandera = false;
                        Thread.Sleep(3000);
                    }
                    catch
                    {
                        bandera = true;
                        Thread.Sleep(500);
                    }
                } while (bandera != false);
                IWebElement menu = driver.FindElement(By.Id("side-menu"));
                ReadOnlyCollection<IWebElement> menus = menu.FindElements(By.TagName("li"));
                int numeroMenu = 0;
                Thread.Sleep(3000);
                for (int a = 0; a < menus.Count; a++)
                {
                    string texto = menus[a].Text;
                    if (texto == "Estado de Cuenta")
                    {
                        menus[a].Click();
                        numeroMenu = a;
                        a = menus.Count;
                    }
                }
                Thread.Sleep(1000);
                ReadOnlyCollection<IWebElement> submenus = menus[numeroMenu].FindElements(By.TagName("li"));
                for (int a = 0; a < submenus.Count; a++)
                {
                    string texto = submenus[a].Text;
                    if (texto == "Saldo en Cuenta")
                    {
                        submenus[a].Click();
                        a = submenus.Count;
                    }
                }

                do
                {
                    try
                    {
                        IWebElement alerta = driver.FindElement(By.Id("WaitProgress"));
                        string propiedad = alerta.GetAttribute("style");
                        if (propiedad == "display: none;")
                        {
                            bandera = true;

                        }
                        else
                        {
                            bandera = false;
                            Thread.Sleep(500);
                        }
                    }
                    catch
                    {
                    }
                } while (bandera != false);

                rutaArchivoSaldo = descargarSaldo(FI, FF);
                Thread.Sleep(2000);
                for (int a = 0; a < submenus.Count; a++)
                {
                    string texto = submenus[a].Text;
                    if (texto == "Documentos Retenidos")
                    {
                        submenus[a].Click();
                        a = submenus.Count;
                    }
                }

                do
                {
                    try
                    {
                        IWebElement alerta = driver.FindElement(By.Id("WaitProgress"));
                        string propiedad = alerta.GetAttribute("style");
                        if (propiedad == "display: none;")
                        {
                            bandera = true;

                        }
                        else
                        {
                            bandera = false;
                            Thread.Sleep(500);
                        }
                    }
                    catch
                    {
                    }
                } while (bandera != false);

                rutaArchivoRetenidos = descargarSaldo(FI, FF);
                Thread.Sleep(2000);
                for (int a = 0; a < submenus.Count; a++)
                {
                    string texto = submenus[a].Text;
                    if (texto == "Movimientos Históricos")
                    {
                        submenus[a].Click();
                        a = submenus.Count;
                    }
                }

                do
                {
                    try
                    {
                        IWebElement alerta = driver.FindElement(By.Id("WaitProgress"));
                        string propiedad = alerta.GetAttribute("style");
                        if (propiedad == "display: none;")
                        {
                            bandera = true;

                        }
                        else
                        {
                            bandera = false;
                            Thread.Sleep(500);
                        }
                    }
                    catch
                    {
                    }
                } while (bandera != false);

                rutaArchivoHistorico = descargarSaldo(FI, FF);
                for (int a = 0; a < menus.Count; a++)
                {
                    string texto = menus[a].Text;
                    if (texto == "Salir")
                    {
                        menus[a].Click();
                        a = menus.Count;
                    }
                }
                Thread.Sleep(1000);
                buttons = driver.FindElements(By.TagName("button"));
                foreach (IWebElement item in buttons)
                {
                    string texto = item.Text;
                    if (texto == "ACEPTAR")
                    {
                        item.Click();
                        break;
                    }
                }
                bandera = false;
                do
                {
                    try
                    {
                        driver.FindElement(By.Name("username"));
                        bandera = false;
                    }
                    catch
                    {
                        bandera = true;
                        Thread.Sleep(500);
                    }
                } while (bandera != false);
                driver.Close();
                driver.Quit();
                if (rutaArchivoSaldo.Length > 0 && rutaArchivoRetenidos.Length > 0 & rutaArchivoHistorico.Length > 0)
                {
                    consolidarArchivos(rutaArchivoSaldo[0], rutaArchivoRetenidos[0], rutaArchivoHistorico[0]);
                }
                else if (rutaArchivoSaldo.Length > 0 && rutaArchivoRetenidos.Length > 0 & rutaArchivoHistorico.Length == 0)
                {
                    consolidarArchivos(rutaArchivoSaldo[0], rutaArchivoRetenidos[0], "Saldo en Cuenta", "Documentos Retenidos");
                }
                else if (rutaArchivoSaldo.Length > 0 && rutaArchivoRetenidos.Length == 0 & rutaArchivoHistorico.Length > 0)
                {
                    consolidarArchivos(rutaArchivoSaldo[0], rutaArchivoHistorico[0], "Saldo en Cuenta", "Movimientos Historicos");
                }
                else if (rutaArchivoSaldo.Length > 0 && rutaArchivoRetenidos.Length == 0 & rutaArchivoHistorico.Length == 0)
                {
                    string rut = moverArchivo(rutaArchivoSaldo[0], "Saldo en cuenta");
                    Excel.Application miExcel = new Excel.Application();
                    miExcel.Visible = true;
                    miExcel.DisplayAlerts = false;
                    Excel.Workbook libro = miExcel.Workbooks.Open(rut);
                    ((Excel.Worksheet)miExcel.Sheets[1]).Select();
                    Excel.Worksheet SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
                    SheetExcel.Columns.EntireColumn.AutoFit();
                    Excel.Range rango = SheetExcel.get_Range("1:1");
                    rango.Select();
                    rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
                    rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                    rango.Font.Bold = true;
                    SheetExcel.Name = "Saldo en Cuenta";
                    libro.Save();
                }
                else if (rutaArchivoSaldo.Length == 0 && rutaArchivoRetenidos.Length == 0 & rutaArchivoHistorico.Length > 0)
                {
                    string rut = moverArchivo(rutaArchivoHistorico[0], "Movientos Historicos");
                    Excel.Application miExcel = new Excel.Application();
                    miExcel.Visible = true;
                    miExcel.DisplayAlerts = false;
                    Excel.Workbook libro = miExcel.Workbooks.Open(rut);
                    ((Excel.Worksheet)miExcel.Sheets[1]).Select();
                    Excel.Worksheet SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
                    SheetExcel.Columns.EntireColumn.AutoFit(); ;
                    Excel.Range rango = SheetExcel.get_Range("1:1");
                    rango.Select();
                    rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
                    rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                    rango.Font.Bold = true;
                    SheetExcel.Name = "Movimientos Historicos";
                    libro.Save();
                }
                else if (rutaArchivoSaldo.Length == 0 && rutaArchivoRetenidos.Length > 0 & rutaArchivoHistorico.Length == 0)
                {
                    string rut = moverArchivo(rutaArchivoRetenidos[0], "Documentos Retenidos");
                    Excel.Application miExcel = new Excel.Application();
                    miExcel.Visible = true;
                    miExcel.DisplayAlerts = false;
                    Excel.Workbook libro = miExcel.Workbooks.Open(rut);
                    ((Excel.Worksheet)miExcel.Sheets[1]).Select();
                    Excel.Worksheet SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
                    SheetExcel.Columns.EntireColumn.AutoFit();
                    Excel.Range rango = SheetExcel.get_Range("1:1");
                    rango.Select();
                    rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
                    rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
                    rango.Font.Bold = true;
                    SheetExcel.Name = "Documentos Retenidos";
                    libro.Save();
                }
                else if (rutaArchivoSaldo.Length == 0 && rutaArchivoRetenidos.Length > 0 & rutaArchivoHistorico.Length > 0)
                {
                    consolidarArchivos(rutaArchivoRetenidos[0], rutaArchivoHistorico[0], "Documentos Retenidos", "Movimientos Historicos");
                }
                else if (rutaArchivoSaldo.Length == 0 && rutaArchivoRetenidos.Length == 0 & rutaArchivoHistorico.Length == 0)
                {
                }
            }
            else
            {
                MessageBox.Show("La consulta no puede ser mayor a 90 dias");
            }
            return exito;
        }

        public string[] NombresArchivosNuevos(string[] FilesBefore, string[] FilesAfter)
        {
            string[] Archivos = new string[10];
            foreach (string xl in FilesBefore)
                FilesAfter = Array.FindAll(FilesAfter, s => !s.Equals(xl));
            return (FilesAfter);
        }

        public string moverArchivo(string rutaArchivo, string Saldos)
        {
            string ruta = "";
            if (!Directory.Exists(rutaEscritorio + @"\Archivos Generados\")) Directory.CreateDirectory(rutaEscritorio + @"\Archivos Generados\");
            string nombre = "Chedraui Estado de Cuenta " + nombreAleatorio() + ".xlsx";
            File.Move(rutaArchivo, rutaEscritorio + @"\Archivos Generados\" + nombre);
            ruta = rutaEscritorio + @"\Archivos Generados\" + nombre;
            return ruta;
        }

        public void consolidarArchivos(string ruta1, string ruta2, string nombre1, string nombre2)
        {
            Excel.Application miExcel = new Excel.Application();
            miExcel.Visible = true;
            miExcel.DisplayAlerts = false;
            Excel.Workbook libro = miExcel.Workbooks.Open(ruta1);
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            Excel.Worksheet SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
            SheetExcel.Columns.EntireColumn.AutoFit();
            Excel.Workbook librof = miExcel.Workbooks.Add();
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            Excel.Worksheet PestañaF = (Excel.Worksheet)librof.ActiveSheet;
            Excel.Range rango = PestañaF.get_Range("A:Z");
            rango.Select();
            rango.NumberFormat = "@";
            copiarArchivos(SheetExcel, PestañaF);
            PestañaF.Name = nombre1;
            PestañaF.Columns.EntireColumn.AutoFit();
            rango = PestañaF.get_Range("1:1");
            rango.Select();
            rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
            rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            rango.Font.Bold = true;
            libro.Close();

            libro = miExcel.Workbooks.Open(ruta2);
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
            SheetExcel.Columns.EntireColumn.AutoFit();
            librof.Worksheets.Add();
            librof.Activate();
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            PestañaF = (Excel.Worksheet)librof.ActiveSheet;
            rango = PestañaF.get_Range("A:Z");
            rango.Select();
            rango.NumberFormat = "@";
            copiarArchivos(SheetExcel, PestañaF);
            PestañaF.Name = nombre2;
            PestañaF.Columns.EntireColumn.AutoFit();
            rango = PestañaF.get_Range("1:1");
            rango.Select();
            rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
            rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            rango.Font.Bold = true;
            libro.Close();

            if (!Directory.Exists(rutaEscritorio + @"\Archivos Generados\")) Directory.CreateDirectory(rutaEscritorio + @"\Archivos Generados\");
            string nombre = "Chedraui Estado de Cuenta " + nombreAleatorio() + ".xlsx";
            librof.SaveAs(rutaEscritorio + @"\Archivos Generados\" + nombre);
            File.Delete(ruta1);
            File.Delete(ruta2);
        }

        public void consolidarArchivos(string ruta1, string ruta2, string ruta3)
        {

            Excel.Application miExcel = new Excel.Application();
            miExcel.Visible = true;
            miExcel.DisplayAlerts = false;
            Excel.Workbook libro = miExcel.Workbooks.Open(ruta1);
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            Excel.Worksheet SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
            SheetExcel.Columns.EntireColumn.AutoFit();
            Excel.Workbook librof = miExcel.Workbooks.Add();
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            Excel.Worksheet PestañaF = (Excel.Worksheet)librof.ActiveSheet;
            Excel.Range rango = PestañaF.get_Range("A:Z");
            rango.Select();
            rango.NumberFormat = "@";
            copiarArchivos(SheetExcel, PestañaF);
            PestañaF.Name = "Saldo en Cuenta";
            PestañaF.Columns.EntireColumn.AutoFit();
            rango = PestañaF.get_Range("1:1");
            rango.Select();
            rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
            rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            rango.Font.Bold = true;
            libro.Close();

            libro = miExcel.Workbooks.Open(ruta2);
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
            SheetExcel.Columns.EntireColumn.AutoFit();
            librof.Worksheets.Add();
            librof.Activate();
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            PestañaF = (Excel.Worksheet)librof.ActiveSheet;
            rango = PestañaF.get_Range("A:Z");
            rango.Select();
            rango.NumberFormat = "@";
            copiarArchivos(SheetExcel, PestañaF);
            PestañaF.Name = "Documentos Retenidos";
            PestañaF.Columns.EntireColumn.AutoFit();
            rango = PestañaF.get_Range("1:1");
            rango.Select();
            rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
            rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            rango.Font.Bold = true;
            libro.Close();


            libro = miExcel.Workbooks.Open(ruta3);
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            SheetExcel = (Excel.Worksheet)libro.ActiveSheet;
            SheetExcel.Columns.EntireColumn.AutoFit();
            librof.Worksheets.Add();
            librof.Activate();
            ((Excel.Worksheet)miExcel.Sheets[1]).Select();
            PestañaF = (Excel.Worksheet)librof.ActiveSheet;
            rango = PestañaF.get_Range("A:Z");
            rango.Select();
            rango.NumberFormat = "@";
            copiarArchivos(SheetExcel, PestañaF);
            PestañaF.Name = "Movimientos Historicos";
            PestañaF.Columns.EntireColumn.AutoFit();
            rango = PestañaF.get_Range("1:1");
            rango.Select();
            rango.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent2;
            rango.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            rango.Font.Bold = true;
            libro.Close();


            if (!Directory.Exists(rutaEscritorio + @"\Archivos Generados\")) Directory.CreateDirectory(rutaEscritorio + @"\Archivos Generados\");
            string nombre = "Chedraui Estado de Cuenta " + nombreAleatorio() + ".xlsx";
            librof.SaveAs(rutaEscritorio + @"\Archivos Generados\" + nombre);
            File.Delete(ruta1);
            File.Delete(ruta2);
            File.Delete(ruta3);
        }

        public string[] TotalArchivosDownloads(string RutaDescargas)
        {
            int i = 0;
            string[] xlsAux = Directory.GetFiles(RutaDescargas, "*.xls");
            foreach (string FileName in xlsAux)
            {
                xlsAux[i] = FileName;
                i++;
            }
            return (xlsAux);
        }

        public String[] descargarSaldo(DateTime Fechaini, DateTime Fechafin)
        {
            String[] archivosAntes = TotalArchivosDownloads(rutaDescargas);
            string[] Nuevos = new string[0];
            bool bandera = true;
            driver.FindElement(By.Name("fechadesde")).Clear();
            driver.FindElement(By.Name("fechadesde")).SendKeys(Fechaini.ToString());

            driver.FindElement(By.Name("fechahasta")).Clear();
            driver.FindElement(By.Name("fechahasta")).SendKeys(Fechafin.ToString());
            Thread.Sleep(1000);
            ReadOnlyCollection<IWebElement> divs = driver.FindElements(By.TagName("div"));
            foreach (IWebElement item in divs)
            {
                string texto = item.Text;
                if (texto == "Resultados")
                {
                    item.Click();
                }
            }
            driver.FindElement(By.Id("btn-consultar")).Click();
            do
            {
                try
                {
                    IWebElement alerta = driver.FindElement(By.Id("WaitProgress"));
                    string propiedad = alerta.GetAttribute("style");
                    if (propiedad == "display: none;")
                    {
                        bandera = false;

                    }
                    else
                    {
                        bandera = true;
                        Thread.Sleep(500);
                    }
                }
                catch
                {
                    //driver.Navigate().Refresh();
                }
            } while (bandera != false);

            String[] archivosDespues;
            bool noHay = false;
            Thread.Sleep(1000);
            ReadOnlyCollection<IWebElement> alertaDatos = driver.FindElements(By.TagName("p"));
            foreach (IWebElement item in alertaDatos)
            {
                string texto = item.Text;
                if (texto.Contains("No existe información para el criterio de búsqueda solicitado."))
                {
                    noHay = true;
                    break;
                }
            }
            if (noHay != true)
            {
                ReadOnlyCollection<IWebElement> tags = driver.FindElements(By.TagName("a"));
                foreach (IWebElement item in tags)
                {
                    string texto = item.Text;
                    if (texto == "EXCEL")
                    {
                        item.Click();
                        break;
                    }
                }
                int intentos = 1;
                do
                {
                    archivosDespues = TotalArchivosDownloads(rutaDescargas);
                    if (archivosDespues.Length > 0 && archivosDespues.Length > archivosAntes.Length)
                    {
                        intentos = 21;
                        Nuevos = NombresArchivosNuevos(archivosAntes, archivosDespues);
                    }
                    else { intentos++; Thread.Sleep(1000); }

                } while (intentos <= 20);
            }
            else
            {
                ReadOnlyCollection<IWebElement> tags = driver.FindElements(By.TagName("button"));
                foreach (IWebElement item in tags)
                {
                    string texto = item.Text;
                    if (texto == "CERRAR")
                    {
                        item.Click();
                        break;
                    }
                }
                Thread.Sleep(1000);
            }

            return Nuevos;
        }

        public String nombreAleatorio()
        {
            String nombre = Convert.ToString(DateTime.Now.Year) +
                Convert.ToString(DateTime.Now.Month.ToString("00")) +
                Convert.ToString(DateTime.Now.Day.ToString("00")) +
                Convert.ToString(DateTime.Now.Hour.ToString("00")) +
                Convert.ToString(DateTime.Now.Minute.ToString("00")) +
                Convert.ToString(DateTime.Now.Second.ToString("00"));
            return nombre;
        }

        public void copiarArchivos(Excel.Worksheet copiar, Excel.Worksheet pegar)
        {
            int ultimafila = ultimoRenglon(copiar, "A");
            int columna = ultimaColumna(copiar, 1);
            for (int i = 1; i <= ultimafila; i++)
            {
                for (int b = 1; b <= columna; b++)
                {
                    pegar.Cells[i, b] = copiar.Cells[i, b].Text;
                }
            }
        }

        private int ultimoRenglon(Excel.Worksheet hoja, string columna)
        {
            Excel.Range rango;
            int a = 1;
            int encontrado = 0, actual = 0;
            do
            {
                rango = hoja.get_Range(columna + a.ToString());
                string text = rango.Text;
                if (text == "" && actual == 0)
                {
                    encontrado++;
                    actual = a;
                }
                else if (text == "")
                {
                    encontrado++;
                }
                else
                {
                    encontrado = 0;
                    actual = 0;
                }
                a++;
            } while (encontrado < 10);
            return actual - 1;
        }

        private int ultimaColumna(Excel.Worksheet hoja, int fila)
        {
            int a = 1;
            int encontrado = 0, actual = 0;
            do
            {
                string text = hoja.Cells[fila, a].Text;
                if (text == "" && actual == 0)
                {
                    encontrado++;
                    actual = a;
                }
                else if (text == "")
                {
                    encontrado++;
                }
                else
                {
                    encontrado = 0;
                    actual = 0;
                }
                a++;
            } while (encontrado < 10);
            return actual - 1;
        }
    }
}