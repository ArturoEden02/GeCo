namespace Pruebas_clase7.Clases
{
    using System;
    using System.IO;
    using System.Linq;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium;
    using System.Collections.ObjectModel;
    using Excel = Microsoft.Office.Interop.Excel;
    using OpenQA.Selenium.Support.UI;
    using System.Data;
    using System.Threading;
    using System.Windows;

    class CEF
    {

        IWebDriver driver;
        string[] FI = new string[0];
        string[] FN = new string[0];
        DateTime miFecha;
        String textoObtenido = "";
        String rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\";
        Excel.Application appExcel;
        Excel.Workbooks booksExcel;
        Excel.Workbook bookExcel;
        Excel.Worksheet SheetExcel;
        Excel.Range unRango;
        Excel.Range Rango; String titulos = "Fec Emision" + "\t" +
             "Num Cliente" + "\t" +
             "dCFD" + "\t" +
             "RFC Emisor" + "\t" +
             "RFC. Receptor" + "\t" +
             "Serie/Folio" + "\t" +
             "Monto" + "\t" +
             "Fec. Envio" + "\t" +
             "Fecha Revision" + "\t" +
             "Estatus" + "\t" +
             "UUID" + "\t" +
             "Estatus Receptor" + "\t" +
             "Estado de Factura" + "\t" +
             "Fec. Tent. Pago Receptor"
             ;

        public Boolean FuncionPrincipal(string usuario, string password, DateTime inicio, DateTime fin)
        {
            Boolean taMuyBien = true;

            string URL = "https://www.masfacturaweb.com.mx/servicio/EnviaCfdOtro.aspx";
            string USER = usuario;
            string PASS = password;

            String fecI = "";
            String fecF = "";
            String dia = "";
            String mes = "";
            String año = "";
            String diaf = "";
            String mesf = "";
            String añof = "";
            int mesSeleccionado = 0;

            fecI = inicio.ToShortDateString();
            fecF = fin.ToShortDateString();
            FN = fecF.Split('/');
            FI = fecI.Split('/');
            dia = FI[0]; diaf = FN[0];
            mes = FI[1]; mesf = FN[1];
            año = FI[2]; añof = FN[2];
            mesSeleccionado = Int32.Parse(mes);

            int añoSeleccionado = Int32.Parse(año);
            int mesActual = miFecha.Month;
            int añoActual = miFecha.Year;
            int resmes = mesActual - mesSeleccionado;
            int restaYers = añoActual - añoSeleccionado;
            resmes = resmes + (restaYers * 12);

            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());

            //objNu4.ReportarLog(RUTA_ARCHIVO_LOG, "Navegando a menu princial de CEF.");

            driver.Navigate().GoToUrl(URL);

            driver.FindElement(By.Id("ctl00_txtUsuario")).Clear();
            driver.FindElement(By.Id("ctl00_txtUsuario")).SendKeys(USER);
            driver.FindElement(By.Id("ctl00_txtContrasenia")).Clear();
            driver.FindElement(By.Id("ctl00_txtContrasenia")).SendKeys(PASS);
            driver.FindElement(By.Id("ctl00_btnAceptar")).Click();

            driver.Navigate().GoToUrl("https://www.masfacturaweb.com.mx/servicio/consultaCfdTercero.aspx");

            //objNu4.ReportarLog(RUTA_ARCHIVO_LOG, "Ingresando fechas la Portal: " + fecI + "----------" + fecF);

            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            js.ExecuteScript("document.getElementById('ctl00_ContentPlaceHolder_mfwCalFecIni_txtFecha').value = '" + fecI + "'");
            js.ExecuteScript("document.getElementById('ctl00_ContentPlaceHolder_mfwCalFecFin_txtFecha').value = '" + fecF + "'");

            //driver.FindElement(By.Id("ctl00_ContentPlaceHolder_btnBuscar")).Click();
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(90));
            IWebElement element = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("ctl00_ContentPlaceHolder_btnBuscar")));
            element.Click();

            //jugos.Wait("Id", "ctl00_ContentPlaceHolder_grvConsulta", driver);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(90));
            element = wait.Until(ExpectedConditions.ElementExists(By.Id("ctl00_ContentPlaceHolder_grvConsulta")));

            string error = driver.FindElement(By.XPath("//*[@id=\"ctl00_ContentPlaceHolder_lblMensajeError\"]")).Text;
            if (error.Equals("No existe información con esos criterios, verifique su búsqueda."))
            {
                //objNu4.ReportarLog(RUTA_ARCHIVO_LOG, "No existe información con esos criterios, verifique su búsqueda.");
                MessageBox.Show("No existe información con esos criterios, cambie el rango de fechas");
                driver.Close(); driver.Quit();
                return false;
            }

            IWebElement tablaConsulta = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta"));
            String tabla1 = tablaConsulta.Text;
            int numeroTR = tablaConsulta.FindElements(By.TagName("tr")).Count;
            ReadOnlyCollection<IWebElement> Paginas = tablaConsulta.FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("a"));

            if (!Directory.Exists(rutaEscritorio)) Directory.CreateDirectory(rutaEscritorio);

            string nombrearchivos = rutaEscritorio + @"CEF Estado de Cuenta " + nombreAleatorio() + ".xlsx";

            appExcel = new Excel.Application();
            booksExcel = appExcel.Workbooks;
            appExcel.Visible = true;
            bookExcel = booksExcel.Add();
            bookExcel.SaveAs(nombrearchivos);

            SheetExcel = (Excel.Worksheet)bookExcel.Worksheets.Add();
            SheetExcel.Name = "CEF Edo Cuenta";
            SheetExcel.Activate();
            Rango = SheetExcel.get_Range("A:Z");
            Rango.NumberFormat = "@";

            Rango = SheetExcel.get_Range("A:A");
            Rango.NumberFormat = "dd / mm / yyyy; @";

            Rango = SheetExcel.get_Range("H:H");
            Rango.NumberFormat = "dd / mm / yyyy; @";

            Rango = SheetExcel.get_Range("N:N");
            Rango.NumberFormat = "dd / mm / yyyy; @";

            Clipboard.Clear();
            Clipboard.SetText(titulos);
            Thread.Sleep(1000);
            Rango = SheetExcel.get_Range("A1");
            Rango.Select();
            Rango.PasteSpecial();

            int unaPagina = 0;
            int clickpagina = 2;
            int numpaginas = Paginas.Count;
            String clickpaginatxt = clickpagina.ToString();
            String mipagina = String.Empty;
            String llave = "llave";

            //Vences que tiene que dar click en la tabla
            int i = 0;
            while (unaPagina < numpaginas)
            {
                int intentos = 0;
                int contadorHojas = 0;

                try
                {
                    for (i = Paginas[0].Text.Trim() == "..." ? 1 : 0; i < Paginas.Count; i++)
                    {
                        if (Paginas[i].Text.Trim() == clickpaginatxt)
                        {
                            leerTabla();
                            Paginas[i].Click();
                            unaPagina++;
                            clickpagina++;
                            if (clickpagina.ToString().EndsWith("1"))
                            {
                                clickpaginatxt = "...";
                            }
                            else
                            {
                                clickpaginatxt = clickpagina.ToString();
                                if (clickpagina.ToString().EndsWith("2"))
                                {
                                    int nuevaLista = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("a")).Count;
                                    numpaginas = nuevaLista;
                                    unaPagina = 0;
                                }
                            }
                            break;
                        }
                    }
                    //Checar si estamos en la ultima hoja
                    if (ChecarUltimaPagina())
                    {
                        System.Threading.Thread.Sleep(2000);
                        leerTabla();
                        unaPagina = numpaginas;
                    }

                }
                catch (Exception)
                {
                    try
                    {
                        numeroTR = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr")).Count;
                        Paginas = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("a"));
                    }
                    catch (Exception)
                    {
                        System.Threading.Thread.Sleep(3000);
                        numeroTR = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr")).Count;
                        Paginas = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("a"));
                    }

                    //Paginas = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("a"));
                    intentos++;
                    System.Threading.Thread.Sleep(1000);
                }
            }

            SheetExcel.Cells[1].Rows.EntireRow.Select();
            SheetExcel.Cells[1].Rows.EntireRow.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
            SheetExcel.Cells[1].Rows.EntireRow.Interior.PatternColorIndex = Excel.XlPattern.xlPatternAutomatic;
            SheetExcel.Cells[1].Rows.EntireRow.Interior.Color = 255;
            SheetExcel.Cells[1].Rows.EntireRow.Interior.TintAndShade = 0;
            SheetExcel.Cells[1].Rows.EntireRow.Interior.PatternTintAndShade = 0;
            SheetExcel.Cells[1].Rows.EntireRow.Font.ThemeColor = Excel.XlThemeColor.xlThemeColorDark1;
            SheetExcel.Cells[1].Rows.EntireRow.Font.TintAndShade = 0;
            SheetExcel.Cells[1].Rows.EntireRow.Font.Bold = true;
            SheetExcel.Cells[1].Rows.EntireRow.Font.Size = 13;
            SheetExcel.Cells[1].Rows.EntireRow.EntireColumn.AutoFit();
            SheetExcel.Cells[1].Rows.EntireRow.EntireRow.AutoFit();

            SheetExcel.Cells.Rows.EntireRow.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            SheetExcel.Cells.Rows.EntireRow.VerticalAlignment = Excel.XlVAlign.xlVAlignBottom;

            //Subiendo el reporte al servidor
            //jugos.subirArchivo(appExcel, bookExcel, "CEF", "EstadoCuenta");

            bookExcel.Save();
            try
            {
                driver.Close();
                driver.Quit();
            }
            catch { }

            return taMuyBien;
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

        private void leerTabla()
        {
            int intentos = 0;
            while (intentos < 10)
            {
                try
                {
                    IWebElement tablaConsulta = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta"));
                    if (tablaConsulta.Text != textoObtenido)
                    {
                        textoObtenido = tablaConsulta.Text;
                        String textoTabla = tablaConsulta.Text.Replace("\r", "");
                        String[] lineasTablaTem = textoTabla.Split('\n');

                        String nuevoTexto = string.Empty;
                        for (int i = 1; i < (lineasTablaTem.Length - 1); i++)
                        {
                            nuevoTexto += lineasTablaTem[i] + " " + lineasTablaTem[i + 1] + "*";
                            i++;
                        }

                        nuevoTexto = nuevoTexto.TrimEnd('*');
                        String[] lineasTabla = nuevoTexto.Split('*');

                        separarAcomodarTexto(lineasTabla);

                        intentos = 10;
                    }
                    else
                    {
                        break;
                    }
                }
                catch (Exception)
                {
                    intentos++;
                }
            }
        }

        private void separarAcomodarTexto(String[] lineas)
        {
            String lineaNueva = string.Empty;
            for (int i = 0; i < lineas.Length; i++)
            {
                String[] xEspacios = lineas[i].Split(' ');
                if (xEspacios.Length == 21)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t";
                    lineaNueva += xEspacios[18] + "\t";
                    lineaNueva += xEspacios[19] + " " + xEspacios[20] + "\t";
                }
                else if (xEspacios.Length == 22)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t";
                    lineaNueva += xEspacios[3] + "\t";
                    lineaNueva += xEspacios[4] + "\t";
                    lineaNueva += xEspacios[5] + "\t";
                    lineaNueva += xEspacios[6] + "\t";
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";
                    lineaNueva += xEspacios[9] + "\t";
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t";
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";
                    lineaNueva += xEspacios[15] + "\t";
                    lineaNueva += xEspacios[16] + "\t";
                    lineaNueva += xEspacios[17] + "\t";
                    lineaNueva += xEspacios[18] + "\t";
                    //lineaNueva += xEspacios[19] + " " + xEspacios[20] + " " + xEspacios[21] + Environment.NewLine;
                }
                else if (xEspacios.Length == 23)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t";
                    lineaNueva += xEspacios[3] + "\t";
                    lineaNueva += xEspacios[4] + "\t";
                    lineaNueva += xEspacios[5] + "\t";
                    lineaNueva += xEspacios[6] + "\t";
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";
                    lineaNueva += xEspacios[9] + "\t";
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t";
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";
                    lineaNueva += xEspacios[15] + "\t";
                    lineaNueva += xEspacios[16] + "\t";
                    lineaNueva += xEspacios[17] + "\t";
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + "\t";
                    //lineaNueva += xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + Environment.NewLine;
                }
                else if (xEspacios.Length == 24)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + " " + xEspacios[18] + " " + xEspacios[19] + "\t";
                    lineaNueva += xEspacios[20] + "\t"; //Estado de Factura
                    //lineaNueva += xEspacios[21] + " " + xEspacios[22] + " " + xEspacios[23] + Environment.NewLine; //Fec Tent Pago RECEPTOR
                }
                else if (xEspacios.Length == 25)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t";
                    lineaNueva += xEspacios[3] + "\t";
                    lineaNueva += xEspacios[4] + "\t";
                    lineaNueva += xEspacios[5] + "\t";
                    lineaNueva += xEspacios[6] + "\t";
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";
                    lineaNueva += xEspacios[9] + "\t";
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t";
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";
                    lineaNueva += xEspacios[15] + "\t";
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t";
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + "\t";
                    //lineaNueva += xEspacios[22] + " " + xEspacios[23] + " " + xEspacios[24] + Environment.NewLine;
                }
                else if (xEspacios.Length == 26)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t"; //Receptor
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + " " + xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + "\t"; //Estado de Factura
                    //lineaNueva += xEspacios[23] + " " + xEspacios[24] + " " + xEspacios[25] + Environment.NewLine; //Fec Tent Pago RECEPTOR
                }
                else if (xEspacios.Length == 27)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + " " + xEspacios[18] + " " + xEspacios[19] + "\t"; //Receptor
                    lineaNueva += xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + " " + xEspacios[23] + " " + xEspacios[24] + "\t"; //Estado de Factura
                    //lineaNueva += xEspacios[25] + " " + xEspacios[26] + " " + xEspacios[25] + Environment.NewLine; //Fec Tent Pago RECEPTOR
                }
                else if (xEspacios.Length == 28)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t"; //Receptor
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + " " + xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + " " + xEspacios[23] + " " + xEspacios[24] + "\t"; //Estado de Factura
                }
                else if (xEspacios.Length == 29)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t"; //Receptor
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + " " + xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + " " + xEspacios[23] + " " + xEspacios[24] + " " + xEspacios[25] + "\t"; //Estado de Factura
                }
                else if (xEspacios.Length == 30)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t"; //Receptor
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + " " + xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + " " + xEspacios[23] + " " + xEspacios[24] + " " + xEspacios[25] + " " + xEspacios[26] + "\t"; //Estado de Factura
                }
                else if (xEspacios.Length >= 30)
                {
                    lineaNueva += xEspacios[0] + " " + xEspacios[1] + " " + xEspacios[2] + "\t"; // Fecha Emision
                    lineaNueva += xEspacios[3] + "\t";  //Num Cliente
                    lineaNueva += xEspacios[4] + "\t";  //idCFD
                    lineaNueva += xEspacios[5] + "\t";  //RFC Emisor
                    lineaNueva += xEspacios[6] + "\t";  //RFC Receptor
                    lineaNueva += xEspacios[7] + xEspacios[8] + "\t";   // Serie + Folio
                    lineaNueva += xEspacios[9] + "\t";  //Monto
                    lineaNueva += xEspacios[10] + " " + xEspacios[11] + " " + xEspacios[12] + "\t"; //Fec. Envio
                    lineaNueva += xEspacios[13] + " " + xEspacios[14] + "\t";   //
                    lineaNueva += xEspacios[15] + "\t"; //Status
                    lineaNueva += xEspacios[16] + "\t"; //UUID
                    lineaNueva += xEspacios[17] + "\t"; //Receptor
                    lineaNueva += xEspacios[18] + " " + xEspacios[19] + " " + xEspacios[20] + " " + xEspacios[21] + " " + xEspacios[22] + " " + xEspacios[23] + " " + xEspacios[24] + " " + xEspacios[25] + " " + xEspacios[26] + "\t"; //Estado de Factura
                }
                else
                {
                    for (int x = 0; x < xEspacios.Length; x++)
                    {
                        lineaNueva += xEspacios[x] + "\t";
                    }

                    //lineaNueva = lineaNueva.TrimEnd('\t') + Environment.NewLine;
                }

                if (xEspacios[xEspacios.Length - 3].Contains("/"))
                {
                    int valor = xEspacios.Length - 3;
                    for (int x = valor; x < xEspacios.Length; x++)
                    {
                        lineaNueva += xEspacios[x] + " " /*Environment.NewLine*/;
                    }
                }
                lineaNueva = lineaNueva.TrimEnd('\t') + Environment.NewLine;
            }
            pintarExcel(lineaNueva);
        }

        private void pintarExcel(String lineaNueva)
        {
            ((Excel.Worksheet)appExcel.Sheets["CEF Edo Cuenta"]).Select();
            SheetExcel = (Excel.Worksheet)bookExcel.ActiveSheet;
            unRango = SheetExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int row = ultimoRenglon(SheetExcel, "A");
            Clipboard.Clear();
            Clipboard.SetText(lineaNueva);
            System.Threading.Thread.Sleep(250);
            SheetExcel.Cells[(row + 1), 1].PasteSpecial();
            Clipboard.Clear();
            bookExcel.Save();
        }

        private Boolean ChecarUltimaPagina()
        {
            Boolean ultimaHoja = false;

            for (int i = 0; i < 20; i++)
            {
                try
                {
                    int numeroTR = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr")).Count;

                    if (numeroTR < 13)
                    {
                        ultimaHoja = true;
                        break;
                    }

                    int numeroTD = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("td")).Count;
                    int span = driver.FindElement(By.Id("ctl00_ContentPlaceHolder_grvConsulta")).FindElements(By.TagName("tr"))[(numeroTR - 1)].FindElements(By.TagName("td"))[numeroTD - 1].FindElements(By.TagName("span")).Count;
                    if (span != 0)
                    {
                        ultimaHoja = true;
                        break;
                    }

                }
                catch (Exception ex) { }
            }
            return ultimaHoja;
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
    }
}
