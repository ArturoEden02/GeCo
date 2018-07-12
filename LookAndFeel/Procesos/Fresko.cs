namespace Pruebas_clase7.Clases
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using Excel = Microsoft.Office.Interop.Excel;
    using System.IO;
    using System.Threading;
    using System.Data;
    using System.Collections.ObjectModel;

    class Fresko
    {
        IWebDriver driver;
        Excel.Application MiExcel;
        Excel.Workbook ArchivoTrabajoExcel;
        String rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\";
        Excel.Worksheet HojaExcel;
        Excel.PivotTable TablaDinamica;
        Excel.Range Rango;
        string[] titulos = new string[] { "Suc.","Folio","Sec.","Alta recibo","Importe recibo","Publicación recibo","Serie","Folio","Imp. facturado","UUID","Proveedor","Compañia","Desglose",
        "Fecha de corte","Fecha de pago","Banco","Tipo","Importe total"};
        public bool FuncionPrincipal(string usuario, string contraseña, string fechaInicial, string fechafinal)
        {
            bool exito = false;

            string anioselect, AXConcat = "", fechas = "";
            int value = 0;
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());
            //driver.Navigate().GoToUrl("http://www.provecomer.com.mx/htmlProvecomer/provecomer.html");
            driver.Navigate().GoToUrl("http://www.provecomer.com.mx/htmlProvecomer/provecomer.html");
            driver.FindElement(By.Name("proveedor")).SendKeys(usuario);//colocando usuario en la pagina "904482"
            driver.FindElement(By.Name("password")).SendKeys(contraseña);//colocando pass en la pagina j223j135
            Wait("Name", "enviar1", driver);
            driver.FindElement(By.Name("enviar1")).Click();
            Thread.Sleep(5000);
            Wait("Name", "areaTrabajo", driver);
            driver.SwitchTo().Frame("areaTrabajo");
            Wait("Id", "boxclose", driver);
            int ventana = driver.FindElements(By.Id("boxclose")).Count;
            if (ventana > 0)
            {
                driver.FindElement(By.Id("boxclose")).Click();
            }
            driver.SwitchTo().DefaultContent();
            Wait("Name", "menu", driver);
            driver.SwitchTo().Frame("menu");
            driver.FindElement(By.LinkText("Parametros de Consulta")).Click();//accediendo al menu del calendario 
            System.Threading.Thread.Sleep(1000);
            IWebDriver driver2 = driver.SwitchTo().Window(driver.WindowHandles.Last());//obteniendo la instancia del segundo pop up donde sale el calendario
                                                                                       //

            string[] FI = fechaInicial.Split('/');

            driver.FindElement(By.Name("dia1")).Clear();
            driver.FindElement(By.Name("dia1")).SendKeys(FI[0]);
            //if (FI[1] == "01" || FI[1] == "1")
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
            driver.FindElement(By.Name("anno1")).SendKeys(FI[2]);
            string[] fechaMod = fechafinal.Split('/');
            driver.FindElement(By.Name("dia2")).Clear();
            driver.FindElement(By.Name("dia2")).SendKeys(fechaMod[0]);
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
            driver2.FindElement(By.XPath("/html/body/form/table/tbody/tr[6]/td/table/tbody/tr/td[1]/a")).Click();
            driver2.FindElement(By.XPath("/html/body/table/tbody/tr[8]/td/a/img")).Click();
            System.Threading.Thread.Sleep(5000);
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.SwitchTo().DefaultContent();
            Thread.Sleep(5000);
            driver.SwitchTo().Frame("menu");
            driver.FindElement(By.LinkText("Desgloses")).Click();
            Thread.Sleep(5000);
            driver.FindElement(By.LinkText("Folios por rango de fechas")).Click();
            Wait("Id", "miTabla7", driver);
            driver.SwitchTo().DefaultContent();
            driver.SwitchTo().Frame("areaTrabajo");
            System.Threading.Thread.Sleep(5000);
            int cont = 2, numehoja = 1, val = 1;
            Wait("Class", "liga", driver);
            var col = driver.FindElements(By.ClassName("liga"));
            MiExcel = new Excel.Application();
            MiExcel.DisplayAlerts = false;
            Excel.Workbooks book = MiExcel.Workbooks;
            do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
            ArchivoTrabajoExcel = book.Add();
            MiExcel.Visible = true;
            ((Excel.Worksheet)this.MiExcel.Sheets[1]).Select();
            do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
            HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
            HojaExcel.Activate();
            HojaExcel.Columns.EntireColumn.NumberFormat = "@";
            HojaExcel.Name = "Estado de Cuenta";
            for (int i = 0; i < titulos.Length; i++)
            {
                HojaExcel.Cells[1, i + 1] = titulos[i];
            }
            ICollection<IWebElement> aux = col;
            int conta = 1;

            IWebElement tabla_datos_For = driver.FindElement(By.TagName("tbody"));
            var tr_collection_For = tabla_datos_For.FindElements(By.TagName("tr"));

            int contLink = 0;
            for (int ind_tr = 0; ind_tr < tr_collection_For.Count; ind_tr++)
            {
                Wait("TagName", "tbody", driver);
                IWebElement tabla_datos = driver.FindElement(By.TagName("tbody"));
                var tr_collection = tabla_datos.FindElements(By.TagName("tr"));
                var td = tr_collection[ind_tr].FindElements(By.TagName("td"));
                if (contLink >= 9 && td.Count == 14)
                {
                    Wait("xPath", "//*[@id=\"GeneraReporteFrm\"]/table/tbody/tr[" + ind_tr + "]/td[2]/a/img", driver);
                    var id = tr_collection[ind_tr].FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody/tr[" + ind_tr + "]/td[2]/a/img"));
                    id.Click();
                    Wait("xPath", "//*[@id=\"GeneraReporteFrm\"]/table/tbody", driver);
                    IWebElement tabla1 = driver.FindElement(By.XPath("//*[@id=\"GeneraReporteFrm\"]/table/tbody"));
                    PintarTabla(tabla1, HojaExcel);
                    Wait("xPath", "//*[@id='GeneraReporteFrm']/table/tbody/tr[1]/td/table/tbody/tr/td[3]/a", driver);
                    driver.FindElement(By.XPath("//*[@id='GeneraReporteFrm']/table/tbody/tr[1]/td/table/tbody/tr/td[3]/a")).Click();
                    numehoja = numehoja + 1;
                }
                contLink++;
            }
            HojaExcel.Columns.EntireRow.AutoFit();
            if (!Directory.Exists(rutaEscritorio)) Directory.CreateDirectory(rutaEscritorio);
            ArchivoTrabajoExcel.SaveAs(rutaEscritorio + "Fresko Estado de Cuenta " + nombreAleatorio() + ".xlsx");
            driver.Close();
            driver.Quit();
            exito = true;
            return exito;
        }

        private bool Wait(string tipo, string IDElement, IWebDriver driver)
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

        private void PintarTabla(IWebElement tabla, Excel.Worksheet pestaña)
        {
            int fila = ultimoRenglon(HojaExcel, "A");
            int inicial = fila + 1;
            pestaña.Cells[inicial, 1].Select();
            string proveedor = "", compañia = "", desglose = "", fechacorte = "", fechapago = "", banco = "", tipo = "", importe = "";
            ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
            for (int i = 0; i < filas.Count; i++)
            {
                if (i > 12)
                {

                    ReadOnlyCollection<IWebElement> columnas = filas[i].FindElements(By.TagName("td"));
                    for (int a = 0; a < columnas.Count; a++)
                    {
                        pestaña.Cells[(fila + 1), (a + 1)] = columnas[a].Text;
                    }
                    fila++;
                }
                else
                {
                    string datos = filas[i].Text;
                    if (datos.StartsWith("Proveedor:"))
                    {
                        proveedor = datos.Replace("Proveedor:", "");
                    }
                    else if (datos.StartsWith("Compañía:"))
                    {
                        compañia = datos.Replace("Compañía:", "");
                    }
                    else if (datos.StartsWith("Desglose:"))
                    {
                        desglose = datos.Replace("Desglose:", "");
                    }
                    else if (datos.StartsWith("Fecha de corte:"))
                    {
                        fechacorte = datos.Replace("Fecha de corte:", "");
                    }
                    else if (datos.StartsWith("Fecha de pago:"))
                    {
                        fechapago = datos.Replace("Fecha de pago:", "");
                    }
                    else if (datos.StartsWith("Banco:"))
                    {
                        banco = datos.Replace("Banco:", "");
                    }
                    else if (datos.StartsWith("Tipo:"))
                    {
                        tipo = datos.Replace("Tipo:", "");
                    }
                    else if (datos.StartsWith("Importe total:"))
                    {
                        importe = datos.Replace("Importe total:", "");
                    }
                }
            }
            int final = ultimoRenglon(pestaña, "A");
            for (int i = inicial; i <= final; i++)
            {
                pestaña.Cells[i, 11] = proveedor;
                pestaña.Cells[i, 12] = compañia;
                pestaña.Cells[i, 13] = desglose;
                pestaña.Cells[i, 14] = fechacorte;
                pestaña.Cells[i, 15] = fechapago;
                pestaña.Cells[i, 16] = banco;
                pestaña.Cells[i, 17] = tipo;
                pestaña.Cells[i, 18] = importe;
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