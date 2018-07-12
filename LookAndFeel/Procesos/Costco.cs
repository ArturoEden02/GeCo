namespace Pruebas_clase7.Clases
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System.Threading;
    using System.Collections.ObjectModel;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.IO;

    class Costco
    {
        private int intentos;
        private IWebDriver driver;
        private string xpath1 = "//li[3]/a/span";
        Excel.Application MiExcel;
        Excel.Workbook ArchivoTrabajoExcel;
        Excel.Worksheet HojaExcel;
        Excel.Range Rango;
        string paginaPrin = "", numStrac = "";
        IWebElement tabla;
        string[] loslInks;
        int trunco = 0;

        public void FuncionPrincipal(string usuario, string contraseña)
        {
            if (!logueoPortal(0, usuario, contraseña))
            {
                paginaPrin = driver.Url;
                MiExcel = new Excel.Application();
                MiExcel.DisplayAlerts = false;
                MiExcel.Visible = true;
                ArchivoTrabajoExcel = MiExcel.Workbooks.Add();
                ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
                HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
                HojaExcel.Name = "Reporte-Age";
                string[] titulo = new string[] { "No Documento", "Status", "Fecha de Documento", "Fecha de Vencimiento", "Antigüedad", "Neto", "Descuento", "Departamento", "Sucursal", "No Pedido" };
                for (int i = 0; i < titulo.Length; i++)
                {
                    HojaExcel.Cells[1, i + 1] = titulo[i];
                }
                Rango = HojaExcel.get_Range("A1:J1");
                HojaExcel.get_Range("A1:J1").Interior.Color = System.Drawing.Color.FromArgb(0x808080);
                Rango.Font.Color = "-16711681";
                Rango = HojaExcel.get_Range("A2:A2");
                IWebElement DesgloseSaldos = driver.FindElement(By.XPath("//*[@id=\"unpaidTable\"]/tbody"));
                PintarTabla(DesgloseSaldos, HojaExcel, 1);
                HojaExcel.Columns.EntireColumn.AutoFit();
                string msg = "";
                bool respuesta = false;

                tabla = driver.FindElement(By.XPath("//*[@id=\"unpaidTable\"]/tbody"));
                string[] numeros = tabla.Text.Split('\n');
                for (int a = 0; a < numeros.Length; a++)
                {
                    string[] stract = numeros[a].Split(' ');
                    numStrac = numStrac + stract[9].Replace("\r", "") + Environment.NewLine;
                }
                loslInks = numStrac.Split('\n');
                ObtenerRelacion(loslInks, usuario, contraseña);
                driver.Close();
                driver.Quit();
                string nombre = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\";
                if (!Directory.Exists(nombre)) Directory.CreateDirectory(nombre);
                nombre += "Estado de Cuenta Costco " + nombreAleatorio() + ".xlsx";
                ArchivoTrabajoExcel.SaveAs(nombre);
            }
        }

        private bool logueoPortal(int intento, string usuario, string contraseña)
        {
            bool banderosa = false;
            do
            {
                try
                {
                    bool exploradorIE = false;

                    var driverService = ChromeDriverService.CreateDefaultService();
                    driverService.HideCommandPromptWindow = true;
                    driver = new ChromeDriver(driverService, new ChromeOptions());


                    try { driver.Navigate().GoToUrl("https://www3.costco.com.mx/wps/portal/publico/!ut/p/a1/04_Sj9CPykssy0xPLMnMz0vMAfGjzOLN_Q09PSwtDLz8fX0sDBzdQ4O9PC1cDN0tzIEKIoEKDHAARwN8-g38zKD68SggYH-4fhReKwIMoArwOLEgNzTCINNREQDZPxVX/dl5/d5/L2dBISEvZ0FBIS9nQSEh/"); }
                    catch (Exception)
                    {
                    }
                    ReadOnlyCollection<IWebElement> dat = driver.FindElements(By.XPath("/html/body/div[4]"));
                    if (dat.Count > 0)
                    {
                        comunicado(dat[0]);
                    }
                    driver.FindElement(By.Id("usuario")).Clear();
                    driver.FindElement(By.Id("usuario")).SendKeys(usuario);
                    driver.FindElement(By.Id("clave")).Clear();
                    driver.FindElement(By.Id("clave")).SendKeys(contraseña);
                    
                    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    IWebElement element = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("btnEnviar")));
                    element.Click();
                    
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath(xpath1)));
                    element.Click();
                    
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[3]/div/nav/ul/li[2]/a/span")));
                    element.Click();

                    driver.FindElement(By.XPath("//*[@id=\"logoutlink\"]"));
                    banderosa = false;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    banderosa = true;
                    intentos++;
                    driver.Close();
                    driver.Quit();
                    if (intentos > 4)
                    {
                        return true;
                    }
                }
            } while (banderosa);
            return banderosa;
        }

        private void comunicado(IWebElement comu)
        {
            if (comu.Text != "")
            {

                ReadOnlyCollection<IWebElement> divs = comu.FindElements(By.TagName("div"));
                if (divs[0].Text.Contains("Comunicado"))
                {

                    ReadOnlyCollection<IWebElement> botones = comu.FindElements(By.TagName("button"));
                    for (int a = 0; a < botones.Count; a++)
                    {
                        string texto = botones[a].Text;
                        if (texto.Contains("Aceptar"))
                        {
                            botones[a].Click();
                            a = botones.Count;
                        }
                    }
                }
                ReadOnlyCollection<IWebElement> dat = driver.FindElements(By.XPath("/html/body/div[4]"));
                if (dat.Count > 0)
                {
                    comunicado(dat[0]);
                }
            }
        }

        private void PintarTabla(IWebElement tabla, Excel.Worksheet pestaña, int fila)
        {
            ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
            for (int i = 0; i < filas.Count; i++)
            {
                ReadOnlyCollection<IWebElement> columnas = filas[i].FindElements(By.TagName("td"));
                for (int a = 0; a < columnas.Count; a++)
                {
                    MiExcel.Cells[(fila + (i + 1)), (a + 1)] = columnas[a].Text;
                }
            }
        }

        private void PintarTabla(IWebElement tabla, Excel.Worksheet pestaña, int fila, int columnaLink)
        {
            ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
            for (int i = 0; i < filas.Count; i++)
            {
                ReadOnlyCollection<IWebElement> columnas = filas[i].FindElements(By.TagName("td"));
                for (int a = 0; a < columnas.Count; a++)
                {
                    if (a == columnaLink)
                    {
                        ReadOnlyCollection<IWebElement> element = columnas[a].FindElements(By.TagName("a"));
                        if (element.Count > 0)
                        {
                            MiExcel.Cells[(fila + (i + 1)), (a + 1)] = element[0].GetAttribute("href");
                        }
                    }
                    else
                    {
                        MiExcel.Cells[(fila + (i + 1)), (a + 1)] = columnas[a].Text;
                    }
                }
            }
        }

        private void PintarTabla(IWebElement tabla, Excel.Worksheet pestaña, int fila, int columnaLink, int inicio, int fin)
        {
            ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
            for (int i = inicio; i < fin; i++)
            {
                ReadOnlyCollection<IWebElement> columnas = filas[i].FindElements(By.TagName("td"));
                for (int a = 0; a < columnas.Count; a++)
                {
                    if (a == columnaLink)
                    {
                        ReadOnlyCollection<IWebElement> element = columnas[a].FindElements(By.TagName("a"));
                        if (element.Count > 0)
                        {
                            MiExcel.Cells[(fila + (i + 1)), (a + 1)] = element[0].GetAttribute("href");
                        }
                    }
                    else
                    {
                        MiExcel.Cells[(fila + (i + 1)), (a + 1)] = columnas[a].Text;
                    }
                }
            }
        }

        private void PintarTabla(IWebElement tabla, Excel.Worksheet pestaña, int fila, int inicio, int fin)
        {
            ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
            for (int i = inicio; i < fin; i++)
            {
                ReadOnlyCollection<IWebElement> columnas = filas[i].FindElements(By.TagName("td"));
                for (int a = 0; a < columnas.Count; a++)
                {
                    MiExcel.Cells[(fila + (i + 1)), (a + 1)] = columnas[a].Text;
                }
            }
        }

        private void ObtenerRelacion(string[] links, string usuario, string contraseña)
        {
            ArchivoTrabajoExcel.Worksheets.Add();
            ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            HojaExcel.Name = "Detalle-Pedidos";

            string[] encabezado = new string[] { "Unidades Recibidas", "Item", "No Producto", "Descipcion", "Precio Unitario", "Precio Total", "Departamento", "No Pedido2", "Total del Pedido" };
            for (int i = 0; i < encabezado.Length; i++)
            {
                HojaExcel.Cells[1, i + 1] = encabezado[i];
            }
            Rango = HojaExcel.get_Range("A1:I1");
            HojaExcel.get_Range("A1:I1").Interior.Color = System.Drawing.Color.FromArgb(0x808080);
            Rango.Font.Color = "-16711681";
            for (int a = trunco; a < links.Length; a++)
            {
                if (links[a] != "" && links[a] != null && !links[a].StartsWith("15"))
                {
                    try
                    {
                        driver.FindElement(By.LinkText(links[a].Replace("\r", ""))).Click();
                        string error = driver.FindElement(By.XPath("//*[@id=\"layoutContainers\"]/div/table/tbody/tr/td/table/tbody/tr/td/div/section/div[2]/div[3]")).Text;
                        if (error != "No se encontró ningún resultado en la búsqueda." && error != "No results found.")
                        {
                            string datosTabla = "";
                            string Pedido = driver.FindElement(By.XPath("//*[@id=\"layoutContainers\"]/div/table/tbody/tr/td/table/tbody/tr/td/div/section/div[2]/div[3]/table/tbody/tr[1]/td[2]")).Text;
                            string Departamento = driver.FindElement(By.XPath("//*[@id=\"layoutContainers\"]/div/table/tbody/tr/td/table/tbody/tr/td/div/section/div[2]/div[3]/table/tbody/tr[2]/td[2]")).Text;
                            string totalPedido = driver.FindElement(By.XPath("//*[@id=\"layoutContainers\"]/div/table/tbody/tr/td/table/tbody/tr/td/div/section/div[2]/div[3]/table/tbody/tr[3]/td[2]")).Text;
                            IWebElement detalles = driver.FindElement(By.XPath("//*[@id=\"detallePO\"]/tbody"));
                            int renglon = ultimoRenglon(HojaExcel, "A");
                            HojaExcel.Cells[renglon + 1, 1].Select();
                            PintarTabla(detalles, HojaExcel, renglon);
                            int ultimo = ultimoRenglon(HojaExcel, "A");
                            for (int i = renglon + 1; i <= ultimo; i++)
                            {
                                HojaExcel.Cells[i, 7] = Departamento;
                            }
                            for (int i = renglon + 1; i <= ultimo; i++)
                            {
                                HojaExcel.Cells[i, 8] = Pedido;
                            }
                            for (int i = renglon + 1; i <= ultimo; i++)
                            {
                                HojaExcel.Cells[i, 9] = totalPedido;
                            }
                        }
                        driver.FindElement(By.XPath("//*[@id=\"layoutContainers\"]/div/table/tbody/tr/td/table/tbody/tr/td/div/section/div[2]/table[2]/tbody/tr[1]/td/table/tbody/tr/td[1]/table/tbody/tr/td/a")).Click();
                    }
                    catch (Exception e)
                    {
                        a = a - 1;

                        string titulo = driver.Title;
                        if (titulo == "Saldo en cuenta")
                        {
                            driver.FindElement(By.XPath("//*[@id=\"command\"]/input[2]")).Click();
                        }
                        else
                        {
                            driver.Close(); driver.Quit();
                            logueoPortal(1, usuario, contraseña);
                        }
                    }
                }
            }
            HojaExcel.Columns.EntireColumn.AutoFit();
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
    }
}