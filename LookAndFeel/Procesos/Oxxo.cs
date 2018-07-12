namespace Pruebas_clase7.Clases
{
    using System;
    using System.Linq;
    using System.IO;
    using System.Diagnostics;
    using Excel = Microsoft.Office.Interop.Excel;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System.Collections.ObjectModel;
    using System.Threading;
    using System.Data;
    using System.Windows;

    class Oxxo
    {

        String rutaDescargas = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads";
        String rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        public void FuncionPrincipal(string usuario, string contraseña, DateTime FI, DateTime FF)
        {
            String[] rutaArchivoDescargado = new String[0];
            rutaArchivoDescargado = PortalOxxo(FI, FF, usuario, contraseña);
            if (rutaArchivoDescargado.Length > 0)
            {
                for (int c = 0; c < rutaArchivoDescargado.Length; c++)
                {
                    Excel.Application appExcel = new Excel.Application();
                    Excel.Workbook booksExcel = appExcel.Workbooks.Open(rutaArchivoDescargado[c]);
                    appExcel.Visible = true;
                    appExcel.DisplayAlerts = false;
                    if (!Directory.Exists(rutaEscritorio + @"\Archivos Generados")) Directory.CreateDirectory(rutaEscritorio + @"\Archivos Generados");
                    booksExcel.SaveAs(rutaEscritorio + @"\Archivos Generados\Estado de Cuenta Oxxo " + nombreAleatorio() + ".xlsx");
                }
            }
        }

        public String[] PortalOxxo(DateTime fechaini, DateTime fechafin, string USER, string PASS)
        {
            IWebDriver driver;
            String[] docOxxo = new String[0];

            String FechaI = fechaini.Day + "/" + fechaini.Month + "/" + fechaini.Year;
            String FechaF = fechafin.Day + "/" + fechafin.Month + "/" + fechafin.Year;

            ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());
            driver.Navigate().GoToUrl("https://proveedores.oxxo.com/");
            Wait("Id", "wm_login-username ", driver);
            driver.FindElement(By.Id("wm_login-username")).Clear();
            driver.FindElement(By.Id("wm_login-username")).SendKeys(USER);
            driver.FindElement(By.Id("wm_login-password")).Clear();
            driver.FindElement(By.Id("wm_login-password")).SendKeys(PASS);  //rrvall123%

            driver.FindElement(By.Id("wm_login-password")).SendKeys(OpenQA.Selenium.Keys.Enter);

            Wait("Id", "wm_login-username", driver);

            ReadOnlyCollection<IWebElement> login = driver.FindElements(By.Id("wm_login-username"));
            if (login.Count == 0)
            {
                string[] ANTES = TotalArchivosDownloads(rutaDescargas);
                driver.Navigate().GoToUrl("https://proveedores.oxxo.com/meta/default/folder/0000006981");

                driver.FindElement(By.Id("jsfwmp6989:defaultForm:startDate__date")).SendKeys(FechaI);
                driver.FindElement(By.Id("jsfwmp6989:defaultForm:endDate__date")).SendKeys(FechaF);

                if (true)
                {
                    new SelectElement(driver.FindElement(By.Id("jsfwmp6989:defaultForm:ddwnTipoFactura"))).SelectByText("Standard");
                    driver.FindElement(By.XPath("//option[@value='STANDARD']")).Click();
                }
                else
                {
                    new SelectElement(driver.FindElement(By.Id("jsfwmp6989:defaultForm:ddwnTipoFactura"))).SelectByText("Nota de Crédito");
                    driver.FindElement(By.XPath("//option[@value='CREDIT']")).Click();
                }
                new SelectElement(driver.FindElement(By.Id("jsfwmp6989:defaultForm:ddwnEstatus"))).SelectByText("Todas");

                driver.FindElement(By.Id("jsfwmp6989:defaultForm:btnSearch")).Click();

                ReadOnlyCollection<IWebElement> btn = driver.FindElements(By.Id("jsfwmp6989:defaultForm:htmlGraphicImage4"));
                while (btn.Count == 0)
                {
                    btn = driver.FindElements(By.Id("jsfwmp6989:defaultForm:htmlGraphicImage4"));
                }

                System.Threading.Thread.Sleep(1500);
                driver.FindElement(By.Id("jsfwmp6989:defaultForm:htmlGraphicImage4")).Click();

                driver.FindElement(By.Id("jsfwmp7100:htmlOutputText1")).Click();

                System.Threading.Thread.Sleep(5000);
                //System.Windows.Forms.SendKeys.Send("%g");
                string[] DESP = TotalArchivosDownloads(rutaDescargas);

                for (int w = 0; w < 20; w++)
                {
                    System.Threading.Thread.Sleep(3000);
                    DESP = TotalArchivosDownloads((rutaDescargas));
                    if (ANTES.Length != DESP.Length)
                        break;
                }


                string[] Nuevos = NombresArchivosNuevos(ANTES, DESP);
                System.Threading.Thread.Sleep(5000);
                //DocDespues = Archivos();

                Process[] myProcesses;
                myProcesses = Process.GetProcessesByName("chromedriver");
                foreach (Process myProcess in myProcesses)
                {
                    myProcess.CloseMainWindow();
                }
                System.Threading.Thread.Sleep(4000);
                try
                {
                    driver.Close();
                    driver.Quit();
                }
                catch { }
                string RutaCarpeta = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Desktop\NuJDV\docs\OXXO";
                if (!Directory.Exists(RutaCarpeta)) Directory.CreateDirectory(RutaCarpeta);
                ANTES = TotalArchivosDownloads(RutaCarpeta);
                for (int c = 0; c < Nuevos.Length; c++)
                {
                    string fecha = nombreAleatorio();
                    string NombreArchivo = RutaCarpeta + "\\" + System.IO.Path.GetFileNameWithoutExtension(Nuevos[c]) + fecha + ".xls";
                    if (!File.Exists(NombreArchivo)) File.Copy(Nuevos[c], NombreArchivo);
                    File.Delete(Nuevos[c]);
                }
                if (ANTES.Length == 0)
                {
                    Nuevos = Directory.GetFiles(RutaCarpeta);
                }
                else
                {
                    DESP = TotalArchivosDownloads(RutaCarpeta);

                    for (int w = 0; w < 20; w++)
                    {
                        System.Threading.Thread.Sleep(3000);
                        DESP = TotalArchivosDownloads((RutaCarpeta));
                        if (ANTES.Length != DESP.Length)
                            break;
                    }
                    Nuevos = NombresArchivosNuevos(ANTES, DESP);
                }
                docOxxo = Nuevos;
            }
            else
            {
                MessageBox.Show("Contraseña Incorrecta o no cargo bien el portal. Intentelo nuevamente.");
            }

            return docOxxo;
        }

        public String[] Archivos()
        {
            String[] Archivo = Directory.GetFiles(rutaDescargas, "*");
            return Archivo;
        }

        public String DocDescargado(String[] DocAntes, String[] DocDespues)
        {
            String doc = "";
            var idsNotInB = DocDespues.AsEnumerable().Select(r => r).Except(DocAntes.AsEnumerable().Select(r => r));
            var C = from row in DocDespues.AsEnumerable() join id in idsNotInB on row equals id select row;

            foreach (var item in C)
            {
                doc = item;
            }

            return doc;
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

        public string[] NombresArchivosNuevos(string[] FilesBefore, string[] FilesAfter)
        {
            string[] Archivos = new string[10];
            foreach (string xl in FilesBefore)
                FilesAfter = Array.FindAll(FilesAfter, s => !s.Equals(xl));
            return (FilesAfter);
        }

        private String nombreAleatorio()
        {
            String nombre = Convert.ToString(DateTime.Now.Year) +
                Convert.ToString(DateTime.Now.Month.ToString("00")) +
                Convert.ToString(DateTime.Now.Day.ToString("00")) +
                Convert.ToString(DateTime.Now.Hour.ToString("00")) +
                Convert.ToString(DateTime.Now.Minute.ToString("00")) +
                Convert.ToString(DateTime.Now.Second.ToString("00"));
            return nombre;
        }

        public bool Wait(string tipo, string IDElement, IWebDriver driver)
        {
            bool Seguir = false;
            int intentos = 0;
            if (tipo == "Id")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.Id(IDElement));
                int w = 0;
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.Id(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "LinkText")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.LinkText(IDElement));
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.LinkText(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "xPath")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.XPath(IDElement));
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.XPath(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "cSelector")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.CssSelector(IDElement));
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.CssSelector(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "Name")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.Name(IDElement));
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.Name(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "Class")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.ClassName(IDElement));
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.ClassName(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "TagName")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.TagName(IDElement));
                while (ChecarElemento.Count == 0 && intentos < 10)
                {
                    ChecarElemento = driver.FindElements(By.TagName(IDElement));
                    System.Threading.Thread.Sleep(5000);
                    intentos++;
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            return (Seguir);
        }

    }
}