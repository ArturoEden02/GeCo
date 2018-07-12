namespace Pruebas_clase7.Clases
{
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using System.Windows;
    using Excel = Microsoft.Office.Interop.Excel;

    class Comex
    {

        Excel.Application MiExcel;//Instancia de Excel
        Excel.Workbook ArchivoTrabajoExcel;
        Excel.Worksheet HojaExcel;
        Excel.Range Rango;
        IWebDriver driver;
        int conta = 0, numhoja = 1, indiceTot = 0, subproceso = 0;
        int pestana = 0;
        string sucursal = "", data = "", dia = "", SFF = string.Empty, SFI = string.Empty;
        List<IWebElement> List_Click = new List<IWebElement>();
        List<string> seleccion = new List<string>();
        List<string> LinksPagos = new List<string>();
        List<string> Lista = new List<string>();
        List<string> Nfolio = new List<string>();
        WebDriverWait wait;
        string[] titulosRecibos = new string[] { "Folio de recibo", "Fecha de alta", "Fecha publicacion", "Consultado", "Status", "Remision", "Acla.", "Importe", "Sucursal" };
        string[] titulosDetalles = new string[] { "Folio Rec", "Conse", "Articulo", "Pedido", "Doc. Val.", "Sec.", "Sub-folio", "Plazo", "Cant.","Emp.","Cant. unidad"
            ,"Costo Bruto","Factor","Dto. Adic","Dto. Bol.","Costo unitario","Pct. I.E.P.S","Pct. I.V.A","Importe Total","Pallet","Sucursal"};
        string[] titulosFacturas = new string[] { "Folio Sup", "Serie", "Folio", "Folio Fiscal", "Importe", "Status" };

        public bool FuncionPrincipal(string usuario, string contraseña, DateTime FechaI, DateTime FechaF)
        {
            bool exito = false;
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());
            driver.Navigate().GoToUrl("http://www.proveccm.com/htmlProvecomer/provecomer.html");
            Thread.Sleep(2000);
            try { driver.FindElement(By.Id("boxcloseAvisoSuc")).Click(); } catch { }
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            Thread.Sleep(1000);
            driver.FindElement(By.Name("proveedor")).Clear();
            driver.FindElement(By.Name("proveedor")).SendKeys(usuario);
            driver.FindElement(By.Name("password")).Clear();
            driver.FindElement(By.Name("password")).SendKeys(contraseña);
            Thread.Sleep(1000);
            driver.FindElement(By.Name("enviar1")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            Thread.Sleep(1000);
            driver.SwitchTo().Frame("areaTrabajo");
            driver.FindElement(By.Id("boxclose")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            string instanciaPrincipal = driver.WindowHandles.Last();
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.SwitchTo().Frame("menu");
            driver.FindElement(By.LinkText("Parametros de Consulta")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            Thread.Sleep(1000);
            driver.SwitchTo().Window(driver.WindowHandles.Last());

            string[] FI = new string[0];
            dia = FechaI.ToShortDateString();
            FI = dia.Split('/');

            SFI = "";
            SFI = FI[0] + "/" + FI[1] + "/" + FI[2];
            driver.FindElement(By.Name("dia1")).Clear();
            driver.FindElement(By.Name("dia1")).SendKeys("01");
            switch (FI[1])
            {
                case "01": driver.FindElement(By.Name("mes1")).SendKeys("Ene"); break;
                case "02": driver.FindElement(By.Name("mes1")).SendKeys("Feb"); break;
                case "03": driver.FindElement(By.Name("mes1")).SendKeys("Mar"); break;
                case "04": driver.FindElement(By.Name("mes1")).SendKeys("Abr"); break;
                case "05": driver.FindElement(By.Name("mes1")).SendKeys("May"); break;
                case "06": driver.FindElement(By.Name("mes1")).SendKeys("Jun"); break;
                case "07": driver.FindElement(By.Name("mes1")).SendKeys("Jul"); break;
                case "08": driver.FindElement(By.Name("mes1")).SendKeys("Ago"); break;
                case "09": driver.FindElement(By.Name("mes1")).SendKeys("Sep"); break;
                case "10": driver.FindElement(By.Name("mes1")).SendKeys("Oct"); break;
                case "11": driver.FindElement(By.Name("mes1")).SendKeys("Nov"); break;
                case "12": driver.FindElement(By.Name("mes1")).SendKeys("Dic"); break;
            }
            driver.FindElement(By.Name("dia1")).Clear();
            driver.FindElement(By.Name("dia1")).SendKeys(FI[0]);
            driver.FindElement(By.Name("anno1")).SendKeys(FI[2]);

            //Se va a poner la fecha de fin
            string hF = "";
            hF = FechaF.ToShortDateString();
            string[] fechaMod = hF.Split('/');
            SFF = "";
            SFF = fechaMod[0] + "/" + fechaMod[1] + "/" + fechaMod[2]; //fecha

            driver.FindElement(By.Name("dia2")).Clear();
            driver.FindElement(By.Name("dia2")).SendKeys("01");
            switch (fechaMod[1])
            {
                case "01": driver.FindElement(By.Name("mes2")).SendKeys("Ene"); break;
                case "02": driver.FindElement(By.Name("mes2")).SendKeys("Feb"); break;
                case "03": driver.FindElement(By.Name("mes2")).SendKeys("Mar"); break;
                case "04": driver.FindElement(By.Name("mes2")).SendKeys("Abr"); break;
                case "05": driver.FindElement(By.Name("mes2")).SendKeys("May"); break;
                case "06": driver.FindElement(By.Name("mes2")).SendKeys("Jun"); break;
                case "07": driver.FindElement(By.Name("mes2")).SendKeys("Jul"); break;
                case "08": driver.FindElement(By.Name("mes2")).SendKeys("Ago"); break;
                case "09": driver.FindElement(By.Name("mes2")).SendKeys("Sep"); break;
                case "10": driver.FindElement(By.Name("mes2")).SendKeys("Oct"); break;
                case "11": driver.FindElement(By.Name("mes2")).SendKeys("Nov"); break;
                case "12": driver.FindElement(By.Name("mes2")).SendKeys("Dic"); break;
            }
            driver.FindElement(By.Name("dia2")).Clear();
            driver.FindElement(By.Name("dia2")).SendKeys(fechaMod[0]);
            driver.FindElement(By.Name("anno2")).SendKeys(fechaMod[2]);
            Thread.Sleep(250);

            driver.FindElement(By.CssSelector("img[alt=\"Enviar\"]")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            driver.FindElement(By.CssSelector("img[alt=\"Cerrar\"]")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.SwitchTo().Frame("menu");
            driver.FindElement(By.LinkText("Recibos")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);
            driver.FindElement(By.LinkText("Consulta por sucursal")).Click();
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(180);

            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("areaTrabajo");
            ReadOnlyCollection<IWebElement> col = null;
            pestana = 0;
            List<string> centros = new List<string>();
            string noFound = driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody/tr[3]/td")).Text;
            if (noFound != "No se encontró información para los criterios seleccionados.")
            {
                col = driver.FindElements(By.ClassName("liga"));
                ReadOnlyCollection<IWebElement> tab = driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody")).FindElements(By.TagName("tr"));
                bool suc = false;
                for (int i = 0; i < tab.Count; i++)
                {
                    if (tab[i].Text.Contains("Total por")) suc = false;
                    if (suc == true)
                    {
                        ReadOnlyCollection<IWebElement> td = tab[i].FindElements(By.TagName("td"));
                        centros.Add(td[3].Text);
                    }

                    if (tab[i].Text.Contains("Sucursal")) suc = true;
                }

                crearExcel();

                for (int i = 0, tit = 0; i < col.Count; i += 2, tit++)
                {
                    col[i].Click();
                    wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                    wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody")));
                    obtenerRecibos(centros[tit]);
                    col = driver.FindElements(By.ClassName("liga"));
                }
            }
            else
            {
                driver.Close();
                MiExcel.Quit();
                MessageBox.Show("No se encontraron resultados con los parámetros de búsqueda");
                return false;
            }
            string nombre = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\";
            if (!Directory.Exists(nombre)) Directory.CreateDirectory(nombre);
            nombre += "Estado de Cuenta Comex " + nombreAleatorio() + ".xlsx";
            ArchivoTrabajoExcel.SaveAs(nombre);
            return exito;
        }

        private void crearExcel()
        {
            MiExcel = new Excel.Application();
            MiExcel.DisplayAlerts = false;
            MiExcel.Visible = true;
            ArchivoTrabajoExcel = MiExcel.Workbooks.Add();
            ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            HojaExcel.Name = "Recibos";
            HojaExcel.Columns.EntireColumn.NumberFormat = "@";
            for (int i = 0; i < titulosRecibos.Length; i++)
            {
                HojaExcel.Cells[1, i + 1] = titulosRecibos[i];
            }
            ArchivoTrabajoExcel.Worksheets.Add();
            ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            HojaExcel.Name = "Detalle";
            HojaExcel.Columns.EntireColumn.NumberFormat = "@";
            for (int i = 0; i < titulosDetalles.Length; i++)
            {
                HojaExcel.Cells[1, i + 1] = titulosDetalles[i];
            }
            ArchivoTrabajoExcel.Worksheets.Add();
            ((Excel.Worksheet)MiExcel.Sheets[1]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            HojaExcel.Name = "Facturas";
            HojaExcel.Columns.EntireColumn.NumberFormat = "@";
            for (int i = 0; i < titulosFacturas.Length; i++)
            {
                HojaExcel.Cells[1, i + 1] = titulosFacturas[i];
            }
        }

        private void obtenerRecibos(string centro)
        {
            ((Excel.Worksheet)MiExcel.Sheets["Recibos"]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            int filI = ultimoRenglon(HojaExcel, "B") + 1;
            HojaExcel.Cells[filI, 1].Select();
            int filF = filI;
            IWebElement tabla = driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody"));
            List<string> folR = new List<string>();
            ReadOnlyCollection<IWebElement> linksDetalle = tabla.FindElements(By.ClassName("liga"));
            ReadOnlyCollection<IWebElement> detalle = tabla.FindElements(By.TagName("tr"));
            bool tit = false;
            for (int i = 0; i < detalle.Count; i++)
            {
                if (detalle[i].Text.Contains("Total por Sucursal")) tit = true;
                if (tit == true)
                {
                    int Columna = 1;
                    ReadOnlyCollection<IWebElement> colum = detalle[i].FindElements(By.TagName("td"));
                    for (int a = 2; a < colum.Count - 1; a++)
                    {
                        if (a == 2) folR.Add(colum[a].Text);
                        HojaExcel.Cells[filF, Columna] = colum[a].Text;
                        Columna++;
                    }
                    filF++;
                }
                if (detalle[i].Text.Contains("Folio de recibo")) tit = true;
            }
            filF = ultimoRenglon(HojaExcel, "B");
            for (int i = filI; i <= filF; i++)
            {
                HojaExcel.Cells[i, 9] = centro;
            }

            for (int i = 0, suc = 0; i < linksDetalle.Count; i += 2, suc++)
            {
                linksDetalle[i].Click();
                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody")));
                obtenerDetalleFacturas(folR[suc], centro);
                tabla = driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody"));
                linksDetalle = tabla.FindElements(By.ClassName("liga"));
            }
            driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]")).Click();
        }

        private void obtenerDetalleFacturas(string recibo, string sucursal)
        {
            int filaIF = 0;
            ((Excel.Worksheet)MiExcel.Sheets["Detalle"]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            int filI = ultimoRenglon(HojaExcel, "B") + 1;
            HojaExcel.Cells[filI, 1].Select();
            int filF = filI, fac = 0; ;
            IWebElement tabla = driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody"));
            ReadOnlyCollection<IWebElement> detalle = tabla.FindElements(By.TagName("tr"));
            bool tit = false, facs = false;
            for (int i = 0; i < detalle.Count; i++)
            {
                if (detalle[i].Text.Contains("Total por Re calcular")) { tit = false; fac = 2; }
                if (tit == true)
                {
                    int Columna = 2;
                    ReadOnlyCollection<IWebElement> colum = detalle[i].FindElements(By.TagName("td"));
                    for (int a = 0; a < colum.Count; a++)
                    {
                        HojaExcel.Cells[filF, Columna] = colum[a].Text.Replace("\n", " ");
                        Columna++;
                    }
                    filF++;
                }
                if (facs == true)
                {
                    ((Excel.Worksheet)MiExcel.Sheets["Facturas"]).Select();
                    HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
                    int ultima = ultimoRenglon(HojaExcel, "B") + 1;
                    HojaExcel.Cells[ultima, 1].Select();
                    int final = ultima;
                    int Columna = 2;
                    ReadOnlyCollection<IWebElement> colum = detalle[i].FindElements(By.TagName("td"));
                    for (int a = 0; a < colum.Count; a++)
                    {
                        HojaExcel.Cells[final, Columna] = colum[a].Text.Replace("\n", " ");
                        Columna++;
                    }
                    final++;
                }
                if (detalle[i].Text.Contains("Pedido")) tit = true;
                if (detalle[i].Text.Contains("Remisiones o facturas en remisión") && !detalle[i + 1].Text.Contains("Folio fiscal") && fac == 2) { facs = true; i += 3; }
            }

            ((Excel.Worksheet)MiExcel.Sheets["Detalle"]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            filF = ultimoRenglon(HojaExcel, "B");
            for (int i = filI; i <= filF; i++)
            {
                HojaExcel.Cells[i, 1] = recibo;
                HojaExcel.Cells[i, 21] = sucursal;
            }

            ((Excel.Worksheet)MiExcel.Sheets["Facturas"]).Select();
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            filF = ultimoRenglon(HojaExcel, "A") + 1;
            int fin = ultimoRenglon(HojaExcel, "B");
            for (int i = filF; i <= fin; i++)
            {
                HojaExcel.Cells[i, 1] = recibo;
            }
            driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody/tr[1]/td/table/tbody/tr/td[3]")).Click();
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
    }
}