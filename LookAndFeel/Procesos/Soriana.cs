namespace Pruebas_clase7.Clases
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System.Threading;
    using System.Collections.ObjectModel;
    using Excel = Microsoft.Office.Interop.Excel;
    using nu4itFox;
    using System.IO;

    class Soriana
    {
        Excel.Application MiExcel;
        Excel.Workbook ArchivoTrabajoExcel;
        Excel.Worksheet HojaExcel;
        Excel.Range Rango;
        nufox objNuFox = new nufox();

        public void FuncionPrincipal(string usuario, string contraseña)
        {
            MiExcel = new Excel.Application();
            MiExcel.Visible = true;
            MiExcel.DisplayAlerts = false;
            ArchivoTrabajoExcel = MiExcel.Workbooks.Add();
            ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            HojaExcel.Name = "Facturas";
            Rango = HojaExcel.get_Range("A:Z");
            Rango.NumberFormat = "@";
            string[] titulo = new string[] { "MOVTO", "SUC", "FOLIO", "FACTURA", "FACSG", "SUSTAT", "VENC", "DEBITO", "CREDITO", "DIFER", "LINKDIF", "SUBTPAG", "PAG" };
            for (int i = 0; i < titulo.Length; i++)
            {
                HojaExcel.Cells[1, i + 1] = titulo[i];
            }
            Rango = HojaExcel.get_Range("A1:M1");
            Rango.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            Rango.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.OrangeRed);
            Rango.EntireRow.Font.Bold = true;
            IWebDriver driver;
            ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());
            driver.Navigate().GoToUrl("https://www1.soriana.com/site/default.aspx?p=8388");


            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement element = wait.Until(ExpectedConditions.ElementExists(By.Name("input_Email")));

            driver.FindElement(By.Name("input_Email")).Clear();
            driver.FindElement(By.Name("input_Email")).SendKeys(usuario);
            driver.FindElement(By.Name("input_Password")).Clear();
            driver.FindElement(By.Name("input_Password")).SendKeys(contraseña);
            try
            {
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.Name("continuar")));
                element.Click();
            }
            catch (Exception) { }

            Thread.Sleep(1500);

            String textoPortal = driver.FindElement(By.TagName("html")).Text;
            int intPortal = 0;
            while ((textoPortal.Contains("No se puede acceder a este sitio") || textoPortal.Contains("Microsoft OLE DB Provider for ODBC")) && intPortal < 10)
            {

                driver.Navigate().Refresh();
                Thread.Sleep(3000);
                textoPortal = driver.FindElement(By.TagName("html")).Text;
                System.Threading.Thread.Sleep(2000);
                intPortal++;
            }

            if (!textoPortal.Contains("No se puede acceder a este sitio") && !textoPortal.Contains("Microsoft OLE DB Provider for ODBC"))
            {
                NavegandoMprincipal(driver);
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/center/table[2]/tbody/tr[1]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/a[2]")));
                driver.FindElement(By.XPath("/html/body/center/table[2]/tbody/tr[1]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/a[2]")).Click();

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/center/table[2]/tbody/tr[1]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/a[3]")));
                driver.FindElement(By.XPath("/html/body/center/table[2]/tbody/tr[1]/td/table/tbody/tr/td[2]/table/tbody/tr[2]/td/a[3]")).Click();

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Pago Factura")));
                driver.FindElement(By.LinkText("Pago Factura")).Click();

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/center/table[4]/tbody/tr/td/table/tbody/tr[2]/td[1]/form/table/tbody/tr[2]/td/select")));
                element.Click();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/center/table[4]/tbody/tr/td/table/tbody/tr[2]/td[1]/form/table/tbody/tr[2]/td/select/option[2]")));
                element.Click();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/center/table[4]/tbody/tr/td/table/tbody/tr[2]/td[1]/form/table/tbody/tr[8]/td/input")));
                element.Click();

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("/html/body/center/table[4]/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody")));

                ReadOnlyCollection<IWebElement> lista = ObtenerHojas(driver);
                for (int hoja = 1; hoja <= lista.Count; hoja++)
                {
                    int fila = ultimoRenglon(HojaExcel, "B");
                    HojaExcel.Cells[fila + 1, 2].Select();
                    IWebElement tabla = driver.FindElement(By.XPath("/html/body/center/table[4]/tbody/tr/td/table/tbody/tr[2]/td[2]/table/tbody/tr/td/table/tbody"));
                    PintarTabla(tabla, HojaExcel, hoja.ToString());

                    if (hoja < lista.Count)
                    {
                        try
                        {
                            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                            element = wait.Until(ExpectedConditions.ElementToBeClickable(By.LinkText((hoja + 1).ToString())));
                            element.Click();
                        }
                        catch { }
                    }
                }
                List<string> link = new List<string>();
                int ultimo = ultimoRenglon(HojaExcel, "B");
                for (int i = 2; i <= ultimo; i++)
                {
                    Rango = HojaExcel.get_Range("K" + i.ToString());
                    link.Add(Rango.Text);
                }
                HojaExcel.Columns.EntireColumn.AutoFit();
                ArchivoTrabajoExcel.Worksheets.Add();
                ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
                HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
                HojaExcel.Name = "Detalle-Facturas";
                Rango = HojaExcel.get_Range("A:Z");
                Rango.NumberFormat = "@";
                string[] encabezado = new string[] { "BODEGA","FOLIO","NOFACT","TOTFACT","DIFERENCIA","TOTALPAGAD","NUM","ELSEC","CODIGO","DESCRIP","CNTREC","DCTOBASE",
                    "VIGENCIA","COSTOBRUTO","COSTONETO","PEDIDO","CAPEMP","FORMAREC","VIGENCIA2","FACTOR","IMPUESTOS"};
                for (int i = 0; i < encabezado.Length; i++)
                {
                    HojaExcel.Cells[1, i + 1] = encabezado[i];
                }
                Rango = HojaExcel.get_Range("A1:U1");
                Rango.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                Rango.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.OrangeRed);
                Rango.EntireRow.Font.Bold = true;
                Diferencias(link, driver, HojaExcel);
                HojaExcel.Columns.EntireColumn.AutoFit();
                try
                {
                    driver.Close();
                    driver.Quit();
                }
                catch { }
                string nombre = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\";
                if (!Directory.Exists(nombre)) Directory.CreateDirectory(nombre);
                nombre += "Estado de Cuenta Soriana " + nombreAleatorio() + ".xlsx";
                ArchivoTrabajoExcel.SaveAs(nombre);
            }
        }

        private void PintarTabla(IWebElement tabla, Excel.Worksheet pestaña, string pagina)
        {
            int renglon = ultimoRenglon(HojaExcel, "B");
            List<string> datos = new List<string>();
            ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
            ReadOnlyCollection<IWebElement> tot = filas[filas.Count - 7].FindElements(By.TagName("td"));
            for (int i = 0; i < filas.Count; i++)
            {
                if (filas[i].Text.Contains("Sub-Total")) tot = filas[i].FindElements(By.TagName("td"));
            }
            for (int i = 3; i < filas.Count - 7; i++)
            {
                int columna = 1;
                ReadOnlyCollection<IWebElement> columnas = filas[i].FindElements(By.TagName("td"));
                for (int a = 0; a < columnas.Count; a++)
                {
                    if (a == 3)
                    {
                        HojaExcel.Cells[renglon + 1, columna] = columnas[a].Text;
                        columna++;
                        HojaExcel.Cells[renglon + 1, columna] = columnas[a].Text.Replace("-", "");
                        columna++;
                    }
                    else if (a == 4)
                    {
                        string imaText = columnas[a].GetAttribute("innerHTML");
                        if (imaText.Contains("resources/v3_pl.gif"))
                        {
                            HojaExcel.Cells[renglon + 1, columna] = "Partida por liquidar";
                            columna++;
                        }
                        else if (imaText.Contains("resources/v3_pd.gif"))
                        {
                            HojaExcel.Cells[renglon + 1, columna] = "Partida detenida";
                            columna++;
                        }
                        else
                        {
                            HojaExcel.Cells[renglon + 1, columna] = "En Tramite de Factoraje";
                            columna++;
                        }
                    }
                    else if (a == 8)
                    {
                        string imaText = columnas[a].GetAttribute("innerHTML");
                        if (imaText.Contains("resources/v3_ok.gif"))
                        {
                            HojaExcel.Cells[renglon + 1, columna] = "OK";
                            columna++;
                            HojaExcel.Cells[renglon + 1, columna] = "";
                            columna++;
                        }
                        else if (imaText.Contains("resources/v3_dif.gif"))
                        {
                            HojaExcel.Cells[renglon + 1, columna] = "DIFERENCIA";
                            columna++;
                            ReadOnlyCollection<IWebElement> ima = columnas[a].FindElements(By.TagName("a"));
                            HojaExcel.Cells[renglon + 1, columna] = ima[0].GetAttribute("href");
                            columna++;
                        }
                        else if (imaText.Contains("resources/v3_acl.gif"))
                        {
                            HojaExcel.Cells[renglon + 1, columna] = "POR ACLARAR";
                            columna++;
                            ReadOnlyCollection<IWebElement> ima = columnas[a].FindElements(By.TagName("a"));
                            HojaExcel.Cells[renglon + 1, columna] = ima[0].GetAttribute("href");
                            columna++;
                        }
                    }
                    else
                    {
                        HojaExcel.Cells[renglon + 1, columna] = columnas[a].Text;
                        columna++;
                    }

                }

                HojaExcel.Cells[renglon + 1, columna] = tot[1].Text;
                columna++;
                HojaExcel.Cells[renglon + 1, columna] = pagina;
                columna++;
                renglon++;
                if (filas[i + 1].Text.Contains("Sub-Total")) i = filas.Count;
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

        private void NavegandoMprincipal(IWebDriver driver)
        {
            ReadOnlyCollection<IWebElement> losimg = driver.FindElements(By.TagName("img"));
            String[] losSRC = losimg.Select(x => x.GetAttribute("src")).ToArray();

            int intentos = 0;
            Boolean hayBoton = true;
            while (intentos < 10)
            {
                losimg = driver.FindElements(By.TagName("img"));
                losSRC = losimg.Select(x => x.GetAttribute("src")).ToArray();
                for (int i = 0; i < losSRC.Length; i++)
                {
                    if (losSRC[i].Contains("resources/v3_presioneaqui.gif"))
                    {
                        try
                        {
                            losimg[i].Click();
                            Thread.Sleep(2000);
                            losimg[i].Click();
                            i = losSRC.Length;
                        }
                        catch (Exception) { }
                    }
                }
                intentos++;
            }
        }

        public ReadOnlyCollection<IWebElement> ObtenerHojas(IWebDriver driver)
        {
            Wait("cSelector", "body>center>table:nth-child(5)>tbody>tr>td>table>tbody>tr:nth-child(2)>td:nth-child(2)>table>tbody>tr>td>table>tbody>tr:nth-child(52)>td>table>tbody>tr>td:nth-child(2)>table>tbody>tr>td", driver);
            ReadOnlyCollection<IWebElement> Tablalista = driver.FindElements(By.CssSelector("body>center>table:nth-child(5)>tbody>tr>td>table>tbody>tr:nth-child(2)>td:nth-child(2)>table>tbody>tr>td>table>tbody>tr:nth-child(52)>td>table>tbody>tr>td:nth-child(2)>table>tbody>tr>td"));
            ReadOnlyCollection<IWebElement> lista = Tablalista[0].FindElements(By.TagName("a"));
            return lista;
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

        public void Diferencias(List<string> links, IWebDriver web, Excel.Worksheet hoja)
        {
            int fila = ultimoRenglon(hoja, "A");
            for (int item = 0; item < links.Count; item++)
            {
                Excel.Range rg = hoja.get_Range("A" + (fila + 1).ToString());
                rg.Select();
                try
                {
                    string a;
                    if (links[item].Contains("http://proveedor2.soriana.com/sprov/pagos/"))
                    {
                        a = links[item].Replace("http://proveedor2.soriana.com/sprov/pagos/", "");
                    }
                    else
                    {
                        a = links[item].Replace("http://proveedor.soriana.com/sprov/pagos/", "");
                    }
                    if (a.StartsWith("neCdi"))
                    {
                        web.Navigate().GoToUrl(links[item]);

                        Wait("xPath", "/html/body/center/table[5]/tbody/tr[3]", web);
                        string generales = web.FindElement(By.XPath("/html/body/center/table[5]/tbody/tr[1]")).Text;
                        string datosPrec = web.FindElement(By.XPath("/html/body/center/table[5]/tbody/tr[3]/td/table[3]/tbody/tr")).Text + "\r";
                        string datosPreci = web.FindElement(By.XPath("/html/body/center/table[5]/tbody/tr[3]/td/table[1]/tbody")).Text + "\r";
                        IWebElement descripciones = web.FindElement(By.XPath("/html/body/center/table[5]/tbody/tr[3]/td/table[2]/tbody"));

                        string bodega, folio, nofact, totfact, diferencia, totalpag, fechacalculo, fechaven, url;
                        bodega = objNuFox.StrExtract(generales, "Bodega: ", "\r").Replace("\r", "");
                        folio = objNuFox.StrExtract(generales, "Folio: ", "\r").Replace("\r", "");
                        fechacalculo = objNuFox.StrExtract(generales, "Fecha Calculo: ").Replace("\r", "");
                        nofact = objNuFox.StrExtract(datosPreci, "No:", " Fecha").Replace("\r", "");
                        totfact = objNuFox.StrExtract(datosPrec, "Total Facturado ", "\n").Replace("\r", "");
                        diferencia = objNuFox.StrExtract(datosPrec, "Diferencia ", "\n").Replace("\r", "");
                        totalpag = objNuFox.StrExtract(datosPrec, "Total Calculado ", "\n").Replace("\r", "");
                        fechaven = objNuFox.StrExtract(datosPrec, "Fecha Pago\r\n", "$").Replace("\r", "");
                        url = a;
                        ReadOnlyCollection<IWebElement> filasDes = descripciones.FindElements(By.TagName("tr"));
                        for (int z = 2; z < filasDes.Count; z = z + 2)
                        {
                            int columna = 1;
                            ReadOnlyCollection<IWebElement> primeras = filasDes[z].FindElements(By.TagName("td"));
                            ReadOnlyCollection<IWebElement> segundas = filasDes[z + 1].FindElements(By.TagName("td"));
                            HojaExcel.Cells[fila + 1, columna] = bodega; columna++;
                            HojaExcel.Cells[fila + 1, columna] = folio; columna++;
                            HojaExcel.Cells[fila + 1, columna] = nofact; columna++;
                            HojaExcel.Cells[fila + 1, columna] = totfact; columna++;
                            HojaExcel.Cells[fila + 1, columna] = diferencia; columna++;
                            HojaExcel.Cells[fila + 1, columna] = totalpag; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[0].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[1].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[2].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[3].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[4].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[5].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[6].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[7].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = primeras[8].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = segundas[2].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = segundas[3].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = segundas[4].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = fechaven; columna++;
                            HojaExcel.Cells[fila + 1, columna] = segundas[5].Text; columna++;
                            HojaExcel.Cells[fila + 1, columna] = segundas[6].Text; columna++;
                            fila++;
                        }
                    }
                }
                catch (Exception ex) { }
            }
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