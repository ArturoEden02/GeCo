namespace Pruebas_clase7.Clases
{
    using Nu4it;
    using nu4itExcel;
    using nu4itFox;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System;
    using System.Collections.ObjectModel;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Threading;
    using Excel = Microsoft.Office.Interop.Excel;

    public class HEB
    {
        usaR objNu4 = new usaR();
        nuExcel objNuExcel = new nuExcel();
        nufox objNuFox = new nufox();
        Excel.Application MiExcel;
        Excel.Workbook LibroExcel;
        Excel.Worksheet HojaExcel;
        Excel.Range Rango;
        private IWebDriver driver;
        DataTable dtEdoCuenta = new DataTable();
        string USER = "P7_1";
        string PASS = "Delvalle.2";
        String rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        bool ExitoFP = false;
        DateTime fechaInicial, fechaFinal;
        string[] Nombre_Encabezados = new string[100];
        int cuentaArchRepetidos = 1;
        string[] AllNameArchIntegracion = new string[100]; //Arreglo que va a contener los archivos de intengración descargados del portal
        int[] NUMEROdeIntegraciones = new int[100]; //Arreglo que nos guarda en cada posición el numero que corresponde a cada DETALLE del portal
        String[] nombreEncInt = new String[100];
        int contadorDEintegraciones = 0;
        string rutaDescarga = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\HEB\";
        string RutaCarpeta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\HEB";
        string rutaGeneral = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\HEB\EdoCta";
        string nombreCarpetaRaiz;
        int indice3 = 0, ContadorIntegracion = 0, numIntegraciones = 0;

        public void PortalHEB(string USER, string PASS, DateTime FechaInicial, DateTime FechaFinal)
        {
            fechaInicial = FechaInicial;
            fechaFinal = FechaFinal;
            string nombreCarpetaIntegracion, nombreArchivoIntegracion;
            string nombreArchivo;
            string NombreArchGuardado = "";
            string Nombre_Pestania = "";

            if (!Directory.Exists(rutaDescarga))
                Directory.CreateDirectory(rutaDescarga);

            String[] ANTES = TotalArchivosDownloads(rutaDescarga);
            String[] nombreArchivosIntegracion = new String[0];
            String[] nombreArchivosDetalles = new String[0];

            MiExcel = objNuExcel.ObtenerObjetoExcel();
            //objNuExcel.InstanciaExcelVisible(MiExcel);
            objNuExcel.ActivarMensajesAlertas(MiExcel, 0);

            try
            {
                var chromeOptions = new ChromeOptions();
                chromeOptions.AddUserProfilePreference("download.default_directory", rutaDescarga);
                chromeOptions.AddUserProfilePreference("download.prompt_for_download", false);
                chromeOptions.AddUserProfilePreference("disable-popup-blocking", "true");
                var driverService = ChromeDriverService.CreateDefaultService();
                driverService.HideCommandPromptWindow = true;
                driver = new ChromeDriver(driverService, chromeOptions);

                string baseURL = "https://go.heb2b.com.mx/HEBusiness/https://go.heb2b.com.mx/HEBusiness/";
                driver.Navigate().GoToUrl(baseURL);

                Wait("Id", "Usuario");
                driver.FindElement(By.Id("Usuario")).Clear();
                driver.FindElement(By.Id("Usuario")).SendKeys(USER);
                driver.FindElement(By.Id("Password")).Clear();
                driver.FindElement(By.Id("Password")).SendKeys(PASS);

                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                IWebElement element = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("btnIngresarLogin")));
                element.Click();

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"UserMenu\"]/div/nav/div[2]/ul/li[7]/a")));
                element.Click();

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id=\"UserMenu\"]/div/nav/div[2]/ul/li[7]/ul/li[1]/a")));
                element.Click();

                Wait("Id", "FechaInicial");
                driver.FindElement(By.Id("FechaInicial")).Clear();
                driver.FindElement(By.Id("FechaInicial")).SendKeys(FechaInicial.ToShortDateString());
                driver.FindElement(By.Id("FechaFinal")).Clear();
                driver.FindElement(By.Id("FechaFinal")).SendKeys(FechaFinal.ToShortDateString());

                wait = new WebDriverWait(driver, TimeSpan.FromSeconds(120));
                element = wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("btnConsultarFacturasPorPagarConsultar")));
                element.Click();

                //---Preparando carpeta de descargas
                string[] ArchivosANTES = TotalArchivosDownloads();

                //---¿Cuantos resultados encontró?
                Thread.Sleep(4000);
                string Cantidad = driver.FindElement(By.Id("divConsultarBotones")).Text;//lblCantidadArticulos Se encontraron 2 resultados con los datos filtrados
                Cantidad = objNuFox.StrExtract(Cantidad.ToString(), "Se encontraron ", " resultados con los datos filtrados").ToString();
                if (Cantidad.ToString() != "0")
                {
                    //---Click en Guardar
                    Wait("LinkText", "Exportar a hoja de cálculo");
                    Thread.Sleep(3000);
                    string[] ArchivosDESPUES = ArchivosANTES;
                    driver.FindElements(By.LinkText("Exportar a hoja de cálculo"))[0].Click();
                    Thread.Sleep(5000);

                    //---TIEMPO DE ESPERA---
                    //string[] ArchivosDESPUES = ArchivosANTES;
                    for (int w = 0; w < 40; w++)
                    {
                        Thread.Sleep(1000);
                        ArchivosDESPUES = TotalArchivosDownloads();
                        if (ArchivosDESPUES.Length != ArchivosANTES.Length)
                            break;
                        if (w == 20)
                        {
                            driver.FindElements(By.LinkText("Exportar a hoja de cálculo"))[0].Click();
                        }
                    }
                    Thread.Sleep(10000);

                    nombreCarpetaRaiz = @"\Estado de Cuenta" + DateTime.Now.ToString("ddMMyyy hhmmss");
                    nombreArchivo = "Estado de Cuenta";

                    NombreArchGuardado = GuardaExcelDescargado(nombreCarpetaRaiz, nombreArchivo, RutaCarpeta, rutaDescarga, ANTES);
                    Nombre_Pestania = NombreArchGuardado;
                    nombreArchivosIntegracion = NumArchXlsParaDescargar(NombreArchGuardado, RutaCarpeta + nombreCarpetaRaiz, Nombre_Pestania, "Documento", "Fecha de Pago", ref dtEdoCuenta);
                    numIntegraciones = nombreArchivosDetalles.Length;


                    String[] DESPUES = TotalArchivosDownloads(rutaDescarga);
                    //num_Arch_Integracion = ArchivosIntegracionReales(RutaCarpeta + nombreCarpetaRaiz + @"\" + NombreArchGuardado + ".xls", NombreArchGuardado);
                    String[] NamesArchivosIntegracionReales = ArchivosIntegracionReales(RutaCarpeta + nombreCarpetaRaiz + @"\" + NombreArchGuardado + ".xls", NombreArchGuardado);
                    String[] NamesArchivosIntegracionRealesDuplicados = NamesArchivosIntegracionReales;
                    string[] ANTES_Integracion = TotalArchivosDownloads(rutaDescarga);

                    Thread.Sleep(5000);

                    driver.FindElement(By.CssSelector("#dtFacturas_wrapper_length_bottom > div.dataTables_length > select[name=\"dtFacturas_length\"]")).Click();
                    new SelectElement(driver.FindElement(By.CssSelector("#dtFacturas_wrapper_length_bottom > div.dataTables_length > select[name=\"dtFacturas_length\"]"))).SelectByText("Todos");
                    driver.FindElement(By.CssSelector("#dtFacturas_wrapper_length_bottom > div.dataTables_length > select[name=\"dtFacturas_length\"] > option[value=\"-1\"]")).Click();
                    //indice3 = ;
                    Thread.Sleep(5000);
                    int posArchivo = 1, totalArchivos = NamesArchivosIntegracionReales.Length;
                    for (int indice2 = 0; indice2 < NamesArchivosIntegracionReales.Length; indice2++)
                    {
                        ANTES = TotalArchivosDownloads(rutaDescarga);
                        indice3 = indice2 + 1; ;

                        for (int oportunidades = 0; oportunidades < 10; oportunidades++)
                        {
                            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(90));
                            element = wait.Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//a[contains(text(),'Integración')])[" + indice3.ToString() + "]")));
                            element.Click();

                            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(90));
                            element = wait.Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("div.btn-group.open > ul.dropdown-menu > li > a")));
                            element.Click();

                            Thread.Sleep(5000);

                            DESPUES = TotalArchivosDownloads(rutaDescarga);
                            for (int w = 0; w < 20; w++)
                            {
                                DESPUES = TotalArchivosDownloads((rutaDescarga));
                                if (ANTES.Length != DESPUES.Length)
                                {
                                    oportunidades = 10;
                                    break;
                                }
                                Thread.Sleep(1000);
                            }
                        }

                        nombreCarpetaIntegracion = @"\Archivos Integracion";
                        AllNameArchIntegracion[ContadorIntegracion] = nombreArchivosIntegracion[indice2];
                        ContadorIntegracion++;
                        nombreArchivoIntegracion = nombreArchivosIntegracion[indice2];
                        NombreArchGuardado = GuardaExcelDescargado(nombreCarpetaIntegracion, NamesArchivosIntegracionReales[indice2], RutaCarpeta + nombreCarpetaRaiz, rutaDescarga, ANTES_Integracion);
                        NamesArchivosIntegracionRealesDuplicados[indice2] = NombreArchGuardado;

                        posArchivo++;
                    }
                    ArchivosANTES = TotalArchivosDownloads();
                    Thread.Sleep(1000);
                    driver.FindElements(By.LinkText("Exportar a hoja de cálculo"))[0].Click();
                    Thread.Sleep(10000);
                    ArchivosDESPUES = TotalArchivosDownloads();
                    //---Buscando archivo descargado
                    string PathArchivo = NombreArchivoNuevo(ArchivosANTES, ArchivosDESPUES);

                    driver.Navigate().GoToUrl(baseURL + "/HEBusiness/Login/CerrarSesion");
                    //try {  }
                    //catch { driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(180)); }
                    driver.Quit();

                    string NombreArchivoNUEVO = "Consulta de compras del " + FechaInicial.ToString("dd MMM yyyy") + " al " + FechaFinal.ToString("dd MMM yyyy");
                    string ArchivoDEscargado = MoverRenombrarArchivo(PathArchivo, rutaDescarga + @"\ConsultasHEB\", NombreArchivoNUEVO, ".xls");


                    Thread.Sleep(1500);


                    nombreCarpetaIntegracion = @"\Archivos Integracion";
                    LibroExcel = objNuExcel.AbrirArchivoNuevo(MiExcel);
                    objNuExcel.ActivarArchivo(LibroExcel);
                    //objNuExcel.CrearNuevaHoja(ArchivoTrabajoExcel);
                    //int pos = objNuExcel.HojaTrabajoSolicitada(ArchivoTrabajoExcel, "HOJA");
                    HojaExcel = objNuExcel.ActivarPestaniaExcel(1, MiExcel, LibroExcel);
                    objNuExcel.PonerNombreHoja(HojaExcel, "Integraciones");
                    string NombreAcumuladoIntegraciones = "Consulta de compras del " + FechaInicial.ToString("dd MMM yyyy") + " al " + FechaFinal.ToString("dd MMM yyyy");
                    objNuExcel.ArchivoGuardarComo(LibroExcel, RutaCarpeta + nombreCarpetaRaiz + nombreCarpetaIntegracion + @"\" + NombreAcumuladoIntegraciones + ".xls");

                    ConjuntaIntegraciones(nombreCarpetaIntegracion, NamesArchivosIntegracionRealesDuplicados, RutaCarpeta + nombreCarpetaRaiz, MiExcel, LibroExcel, HojaExcel, NombreAcumuladoIntegraciones);

                }

            }
            catch (Exception ex)
            {
                try
                {
                    MiExcel.DisplayAlerts = true;
                    MiExcel.Quit();
                }
                catch (Exception) { }
            }
            finally
            {

            }
        }

        private void ConjuntaIntegraciones(string nombreCarpeta, String[] nombreArchivosIntegracion, string Ruta, Excel.Application MiExcel1, Excel.Workbook ArchTrabajo, Excel.Worksheet Hoja, string nombreExcel)
        {

            int filaTitAct = 0, FilaFolios = 0, NumStringAbuscar = 0, CuentaColumnas = 0;
            int pos = 0, num_UltimaColumna = 0, abiertos, cuentaFOLIOS = 0, INDICEfolios = 0;
            string ObtieneFolio = "", obtieneExtenFolio = "";
            string ObtieneNumOrden = "";
            string rutaArchIntegracion = "";
            string[] folios = new string[1000];
            string[] NumerosDEorden = new string[1000];
            string[] ExtencionesFolio = new string[1000];
            string nombre, cadenaPortapapeles = "";
            int[] indicesEncabezados = new int[1000];
            string[] cadenas = new string[100];

            DataTable Temporal = new DataTable();
            Temporal.Clear();
            int filasReales = 0;

            for (int indice4 = 0; indice4 < nombreArchivosIntegracion.Length; indice4++)
            {
                objNuExcel.InstanciaExcelVisible(MiExcel);
                rutaArchIntegracion = Ruta + nombreCarpeta + @"\";

                if (File.Exists(rutaArchIntegracion + nombreArchivosIntegracion[indice4] + ".xls"))
                {
                    LibroExcel = objNuExcel.AbrirArchivo(rutaArchIntegracion + nombreArchivosIntegracion[indice4] + ".xls", MiExcel);
                    objNuExcel.ActivarArchivo(LibroExcel);
                    HojaExcel = objNuExcel.ActivarPestaniaExcel(1, MiExcel, LibroExcel);
                    filaTitAct = objNuExcel.filaTitulos_2(LibroExcel, HojaExcel, "Departamento:", "Entrada:");
                }
                string[] columnasBuscar = new string[] { "Entrada:", "Total Factura V:", "Días Venc:" };
                NumStringAbuscar = columnasBuscar.Length;

                do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
                ObtieneFolio = EncuentraString("Entrada:\t\t\t\t\t\t", "\t\t\t\t\tTotal", HojaExcel, filaTitAct);
                obtieneExtenFolio = ObtieneFolio.Substring(0, 4);
                ObtieneFolio = ObtieneFolio.Substring((ObtieneFolio.Length) - 7, 7); //obtenemos los 7 últimos dígitos de la cadena Entrada: del archivo integración

                filaTitAct = objNuExcel.filaTitulos_2(LibroExcel, HojaExcel, "Número de Orden:", "Fecha Entrada:");
                columnasBuscar = new string[] { "Número de Orden:", "Fecha Entrada", "Total Faltante:" };
                NumStringAbuscar = columnasBuscar.Length;
                do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
                ObtieneNumOrden = EncuentraString("Número de Orden:\t\t\t", "\t\t\t\t\t\t\tFecha", HojaExcel, filaTitAct);

                folios[cuentaFOLIOS] = ObtieneFolio;
                NumerosDEorden[cuentaFOLIOS] = ObtieneNumOrden;
                ExtencionesFolio[cuentaFOLIOS] = obtieneExtenFolio;

                filaTitAct = objNuExcel.filaTitulos_2(LibroExcel, HojaExcel, "Artículo", "Descripción");

                string[] nombresNuevasColumnas = { "Folio", "Num Pedido", "UPC", "Busqueda", "HEB" };
                int[] posNuevasColumas = { 0, 3, 4, 5, 11 };
                do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
                num_UltimaColumna = objNuExcel.UltimaColumna(LibroExcel, filaTitAct.ToString());
                if (indice4 == 0)
                {
                    CuentaColumnas = 0;
                    for (int indice1 = 0; indice1 < num_UltimaColumna; indice1++)
                    {
                        if (objNuExcel.TextoCelda(MiExcel, filaTitAct, (indice1 + 1)) != "" && objNuExcel.TextoCelda(MiExcel, filaTitAct, (indice1 + 1)) != null)
                        {
                            Nombre_Encabezados[filasReales] = objNuExcel.TextoCelda(MiExcel, filaTitAct, (indice1 + 1));
                            indicesEncabezados[CuentaColumnas] = indice1 + 1;
                            CuentaColumnas = CuentaColumnas + 1;
                            filasReales = filasReales + 1;
                        }
                    }
                    CreaDataTable(Temporal, 0, filasReales, indicesEncabezados, nombresNuevasColumnas, posNuevasColumas);
                }
                Temporal = LlenaDataTable(Temporal, 0, filasReales, HojaExcel, filaTitAct + 1, indicesEncabezados, posNuevasColumas, FilaFolios);
                for (int indice = FilaFolios; indice < Temporal.Rows.Count; indice++)
                {
                    Temporal.Rows[indice][0] = folios[cuentaFOLIOS];
                    Temporal.Rows[indice][3] = NumerosDEorden[cuentaFOLIOS];
                }
                objNuExcel.CerrarArchivo(LibroExcel);
                FilaFolios = Temporal.Rows.Count;
                cuentaFOLIOS++;
                INDICEfolios++;
            }

            // Pegar datos en Excel Consulta de compras *********************************
            abiertos = objNuExcel.ContarArchivoAbiertosNombre(MiExcel, "CONSULTA");
            pos = objNuExcel.PosArchivoAbiertoNombre(MiExcel, "CONSULTA");
            nombre = objNuExcel.NombreArchivoAbiertoPos(MiExcel, pos);
            LibroExcel = objNuExcel.ObtenerArchivoAbierto(MiExcel, pos);
            objNuExcel.ActivarArchivo(LibroExcel);
            pos = objNuExcel.HojaTrabajoSolicitada(LibroExcel, "INTEGRACIONES");
            HojaExcel = objNuExcel.ActivarPestaniaExcel(pos, MiExcel, LibroExcel);

            for (int indice = 0; indice < contadorDEintegraciones; indice++)
                objNuExcel.EscribeTexto(nombreEncInt[indice], 1, indice + 1, HojaExcel);

            Rango = HojaExcel.get_Range("E:F");
            Rango.NumberFormat = "@";
            cadenaPortapapeles = objNu4.ConvierteDTaSTRING(Temporal);
            objNu4.clipboardAlmacenaTexto(cadenaPortapapeles);
            objNuExcel.PegarPortaPapelesRango("A2", HojaExcel);
            objNuExcel.AjustarAnchoColumnaTodaHoja(HojaExcel);
            objNuExcel.ArchivoGuardar(LibroExcel);
        }

        private void ContarDuplicados(Excel.Application Excel, Excel.Workbook Libro, Excel.Worksheet Hoja, int filaTitulos)
        {
            DataTable duplicados = new DataTable();
            string[] columnas = { "I" + filaTitulos, "M" + filaTitulos, "AI" + filaTitulos, "AJ" + filaTitulos };
            //duplicados = jugos.columnasDataTable(columnas, Excel, Hoja);
            duplicados = Utilidades.ColumnasDataTable(MiExcel, HojaExcel, columnas, filaTitulos);
            //duplicados = jugos.PrimerafilaTitulos(duplicados);
            duplicados = Utilidades.PrimeraFilaTitulos(duplicados);

            duplicados.Columns.Add("fila", typeof(string));

            for (int i = 0; i < duplicados.Rows.Count; i++)
            {
                duplicados.Rows[i][4] = i + filaTitulos + 1;
            }

            var vacios = duplicados.AsEnumerable().Where(del => String.IsNullOrEmpty(del.Field<string>("Busqueda")));
            foreach (var fila in vacios.ToList())
                fila.Delete();

            var repetidos = from consulta in duplicados.AsEnumerable()
                            group consulta by consulta.Field<string>("Busqueda") into resBus
                            select new { resBus };

            foreach (var item in repetidos)
            {
                if (item.resBus.Count() > 1)
                {
                    int cajas = 0;
                    for (int i = 0; i < item.resBus.Count(); i++)
                    {
                        cajas += Convert.ToInt32(item.resBus.ElementAt(i).Field<string>("Cajas Facturadas").Replace(",", ""));
                    }

                    if (cajas == Convert.ToInt32(item.resBus.ElementAt(0).Field<string>("Cajas Rec. Cte").Replace(",", "")))
                    {
                        for (int i = 0; i < item.resBus.Count(); i++)
                        {
                            string fila = item.resBus.ElementAt(i).Field<string>("fila");
                            string nuevacaja = item.resBus.ElementAt(i).Field<string>("Cajas Facturadas");
                            objNu4.clipboardAlmacenaTexto(nuevacaja);
                            Rango = Hoja.get_Range("AI" + fila, "AI" + fila);
                            Rango.PasteSpecial();
                        }

                    }

                }

            }
        }

        private string EncuentraString(string cadena1, string cadena2, Excel.Worksheet hoja, int NumeroFila) // función que busca un string entre 2 strings dentro de una fila de excel
        {
            do { Thread.Sleep(20); } while (!MiExcel.Application.Ready);
            Rango = hoja.get_Range("A" + NumeroFila, "BN" + NumeroFila);
            Rango.Copy();

            string columnas = objNu4.clipboardObtenerTexto();
            return objNuFox.StrExtract(columnas, cadena1, cadena2, 1);
        }

        private DataTable CreaDataTable(DataTable temporal, int columnaInicial, int numColumnas, int[] indicesEncabezados, String[] nombresNuevasColumnas, int[] posicionesNuevasColumnas)
        {
            int conta = 0, conta1 = 0, i = 0;
            contadorDEintegraciones = 0;
            for (int indice = columnaInicial; indice < (numColumnas + nombresNuevasColumnas.Length); indice++)// NumEncabezados; indice++)
            {
                if (indice == posicionesNuevasColumnas[conta])
                {
                    temporal.Columns.Add(nombresNuevasColumnas[conta], typeof(string));
                    nombreEncInt[i] = nombresNuevasColumnas[conta];
                    conta++;
                    i++;
                    contadorDEintegraciones++;
                    if (conta == nombresNuevasColumnas.Length)
                        conta = 0;
                }
                else
                {
                    temporal.Columns.Add(Nombre_Encabezados[conta1], typeof(string));
                    nombreEncInt[i] = Nombre_Encabezados[conta1];
                    contadorDEintegraciones++;
                    i++;
                    conta1++;
                }
            }
            return temporal;
        }

        private DataTable LlenaDataTable(DataTable temporal, int ColumnasIni, int ColumnasFin, Excel.Worksheet Hoja, int FILA_ENCABEZADOS, int[] indicesEncabezados, int[] posNuevasColumnas, int FilaFolios)
        {
            int posTit, totRenArr, indRen, totRenDT, aux, maximo = 0;
            int conta = 0, conta1 = 0;
            do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
            string[] DATOS = objNuExcel.LeerColumna((FILA_ENCABEZADOS), indicesEncabezados[0], MiExcel, LibroExcel, Hoja);
            totRenArr = DATOS.Length - 1;
            for (posTit = ColumnasIni; posTit < (ColumnasFin + posNuevasColumnas.Length); posTit++)
            {
                if (posTit == posNuevasColumnas[conta])
                {
                    do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
                    DATOS = objNuExcel.LeerColumna((FILA_ENCABEZADOS), indicesEncabezados[0], MiExcel, LibroExcel, Hoja);
                    //totRenArr = DATOS.Length - 1;
                    if (posNuevasColumnas[conta] == 0)
                    {
                        for (indRen = FilaFolios; indRen < (totRenArr + FilaFolios); indRen++)
                            temporal.Rows.Add("", "");
                    }
                    else
                    {
                        for (indRen = FilaFolios; indRen < (totRenArr + FilaFolios); indRen++)
                            temporal.Rows[indRen][posTit] = "";
                    }
                    conta++;
                    if (conta == posNuevasColumnas.Length)
                        conta = 0;
                }
                else
                {
                    do { Thread.Sleep(10); } while (!MiExcel.Application.Ready);
                    DATOS = objNuExcel.LeerColumna((FILA_ENCABEZADOS), indicesEncabezados[conta1], MiExcel, LibroExcel, Hoja);
                    //totRenArr = DATOS.Length - 1;
                    if (posTit == ColumnasIni)
                    {
                        //maximo = totRenArr;
                        //maximo = maximo - 1;
                        //totRenArr = maximo;
                        for (indRen = FilaFolios; indRen < (totRenArr + FilaFolios); indRen++)
                        {
                            //temporal.Rows[indRen][posTit - 1] = DATOS[indRen];
                            temporal.Rows.Add(DATOS[indRen], "");
                        }
                    }
                    else
                    {
                        totRenDT = totRenArr;// temporal.Rows.Count;// maximo;
                        for (indRen = 0; indRen < totRenDT; indRen++)
                        {
                            if (indRen < totRenArr) { temporal.Rows[indRen + FilaFolios][posTit] = DATOS[indRen]; }//[indRen][indice2-1] = DATOS[indRen]; }
                            else { temporal.Rows[indRen + FilaFolios][posTit] = ""; }//[indRen][indice2-1] = ""; }
                        }
                        if (totRenDT < totRenArr)
                        {
                            for (aux = totRenDT; aux < maximo; aux++)
                            {
                                temporal.Rows.Add();
                                temporal.Rows[aux][posTit] = DATOS[aux];//[indice2-1] = DATOS[aux];
                            }
                        }
                    }
                    conta1++;
                }
            }
            return temporal;
        }

        private void InsertaFolios(int NumTotalFolios, string Folio, int FilaFolios, Excel.Worksheet hoja)
        {
            int indice = 0;
            for (indice = 0; indice < NumTotalFolios; indice++)
                objNuExcel.EscribeTexto(Folio, (indice + FilaFolios), 1, HojaExcel);
        }

        private String[] ArchivosIntegracionReales(string rutaArchivoExcel, string nombrePestania)
        {
            String[] nombreArchivos = new String[0];
            String[] nombreArchivosIntegracionesTot = new String[0];
            String[] nombreArchivosIntegracionesReales = new String[1000];

            int[] posArchivosIntegracionesReales = new int[1000];// = new String[0];
            int totReales = 0, pos = 0, reales = 0;
            int FILA_ENCABEZADOS = 0;
            int ultimaColumnaEncabezados = 0, columnaBuscada = 0;



            if (File.Exists(rutaArchivoExcel))
            {
                LibroExcel = objNuExcel.AbrirArchivo(rutaArchivoExcel, MiExcel);
                objNuExcel.ActivarArchivo(LibroExcel);
                pos = objNuExcel.HojaTrabajoSolicitada(LibroExcel, nombrePestania, 0);
                HojaExcel = objNuExcel.ActivarPestaniaExcel(pos, MiExcel, LibroExcel);
                FILA_ENCABEZADOS = objNuExcel.filaTitulos_2(LibroExcel, HojaExcel, "Documento", "Fecha de Factura");
                //ultimaColumnaEncabezados = objNuExcel.UltimaColumna(HojaExcel, FILA_ENCABEZADOS.ToString());
                ultimaColumnaEncabezados = objNuExcel.UltimaColumna(LibroExcel, FILA_ENCABEZADOS.ToString());
                for (int indice = 1; indice <= ultimaColumnaEncabezados; indice++)
                {
                    if (objNuExcel.LeerTextoCelda(HojaExcel, FILA_ENCABEZADOS, indice) == "Diferencia")
                        columnaBuscada = indice;
                }
                nombreArchivos = objNuExcel.LeerColumna(FILA_ENCABEZADOS + 1, columnaBuscada, MiExcel, LibroExcel, HojaExcel);
                nombreArchivosIntegracionesTot = objNuExcel.LeerColumna(FILA_ENCABEZADOS + 1, 1, MiExcel, LibroExcel, HojaExcel);
                /*if (nombreArchivos.Length < nombreArchivosIntegracionesTot.Length)
                {
                    for (int indice = nombreArchivos.Length; indice <= nombreArchivosIntegracionesTot.Length; indice++)
                        nombreArchivos[indice] = "";
                }*/
                int posi = 0;
                for (int indice = 0; indice < nombreArchivos.Length; indice++)
                {
                    if (nombreArchivos[indice] != "" && nombreArchivos[indice] != null)
                        totReales = totReales + 1;
                }
                nombreArchivosIntegracionesReales = new String[totReales];
                for (int indice = 0; indice < nombreArchivos.Length; indice++)
                {
                    if (nombreArchivos[indice] != "" && nombreArchivos[indice] != null)
                    {
                        nombreArchivosIntegracionesReales[posi] = nombreArchivosIntegracionesTot[indice];
                        posi++;
                    }
                }
                objNuExcel.CerrarArchivo(LibroExcel);
            }
            return nombreArchivosIntegracionesReales;
        }

        private String[] NumArchXlsParaDescargar(string nombreArchXls, string RutaCarpeta, string pestania, string Encabezado1, string Encabezado2, ref DataTable dataTable)
        {
            string rutaArchExcel;
            String[] nombreArchivos = new String[0];
            int totalArchivos = 0, pos = 0, FILA_ENCABEZADOS = 0, pos1;
            string[] nombreTitulos = new string[] { "Documento", "Folio Recibo", "Documento Pagado", "Fecha de Pago" };

            rutaArchExcel = RutaCarpeta + @"\" + nombreArchXls + ".xls";
            if (File.Exists(rutaArchExcel))
            {
                LibroExcel = objNuExcel.AbrirArchivo(rutaArchExcel, MiExcel);
                objNuExcel.ActivarArchivo(LibroExcel);
                pos = objNuExcel.HojaTrabajoSolicitada(LibroExcel, pestania, 0);
                pos1 = objNuExcel.HojaTrabajoSolicitada(LibroExcel, pestania, 1);
                HojaExcel = objNuExcel.ActivarPestaniaExcel(pos, MiExcel, LibroExcel);
                FILA_ENCABEZADOS = objNuExcel.filaTitulos_2(LibroExcel, HojaExcel, Encabezado1, Encabezado2);
                //totalArchivos = objNuExcel.UltimoRenglon(HojaExcel, "A");
                totalArchivos = objNuExcel.UltimoRenglon(LibroExcel, "A");
                totalArchivos = totalArchivos - FILA_ENCABEZADOS;
                nombreArchivos = objNuExcel.LeerColumna(FILA_ENCABEZADOS + 1, 1, MiExcel, LibroExcel, HojaExcel);

                //nombreTitulos = jugos.ColumnaNombreColumna(nombreTitulos, LibroExcel, HojaExcel);
                int filaTitulos = 0;
                nombreTitulos = Utilidades.ColumnaCelda(MiExcel, LibroExcel, HojaExcel, nombreTitulos, ref filaTitulos);
                //dataTable = jugos.columnasDataTable(nombreTitulos, MiExcel, HojaExcel);
                dataTable = Utilidades.ColumnasDataTable(MiExcel, HojaExcel, nombreTitulos, filaTitulos);
                //dataTable = jugos.PrimerafilaTitulos(dataTable);
                dataTable = Utilidades.PrimeraFilaTitulos(dataTable);

                objNuExcel.CerrarArchivo(LibroExcel);////
            }
            return nombreArchivos;
        }

        private string GuardaExcelDescargado(string nombreCarpeta, string nombreArchivo, string RutaCarpeta, string rutaDescarga, string[] ANTES)
        {
            if (!Directory.Exists(RutaCarpeta)) Directory.CreateDirectory(RutaCarpeta);
            string[] DESP = TotalArchivosDownloads(rutaDescarga);
            string fecha = "", NombreArchivo;

            for (int w = 0; w < 20; w++)
            {
                //System.Threading.Thread.Sleep(3000);
                DESP = TotalArchivosDownloads((rutaDescarga));
                if (ANTES.Length != DESP.Length)
                    break;
            }

            string[] Nuevos = NombresArchivosNuevos(ANTES, DESP);
            System.Threading.Thread.Sleep(5000);

            RutaCarpeta = RutaCarpeta + nombreCarpeta + @"\";
            if (!Directory.Exists(RutaCarpeta)) Directory.CreateDirectory(RutaCarpeta);

            ANTES = TotalArchivosDownloads(RutaCarpeta);
            for (int c = 0; c < Nuevos.Length; c++)
            {
                fecha = DateTime.Now.ToString("ddMMyyyy hhmmss");
                NombreArchivo = RutaCarpeta + nombreArchivo + ".xls";//fecha + ".xls";
                if (!File.Exists(NombreArchivo))
                {
                    cuentaArchRepetidos = 1;
                    File.SetAttributes(Nuevos[c], FileAttributes.Normal);
                    File.Move(Nuevos[c], NombreArchivo);
                }
                else
                {
                    nombreArchivo = nombreArchivo + "-" + cuentaArchRepetidos.ToString();
                    NombreArchivo = RutaCarpeta + nombreArchivo + ".xls";//fecha + ".xls";
                    File.SetAttributes(Nuevos[c], FileAttributes.Normal);
                    File.Move(Nuevos[c], NombreArchivo);
                    cuentaArchRepetidos = cuentaArchRepetidos + 1;
                }
                //File.Delete(Nuevos[c]);
            }
            return nombreArchivo;
        }

        private string[] TotalArchivosDownloads(string RutaDescargas)
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

        private string[] NombresArchivosNuevos(string[] FilesBefore, string[] FilesAfter)
        {
            string[] Archivos = new string[10];
            foreach (string xl in FilesBefore)
                FilesAfter = Array.FindAll(FilesAfter, s => !s.Equals(xl));
            return (FilesAfter);
        }

        private void DescargarIntegracion(IWebDriver web)
        {
            int diferentes = 0;
            IWebElement paginas = web.FindElement(By.XPath("//*[@id=\"dtFacturas_wrapper_pagination_bottom\"]/div/ul"));
            ReadOnlyCollection<IWebElement> pagi = paginas.FindElements(By.TagName("li"));
            for (int a = 1; a <= (pagi.Count - 2); a++)
            {
                if (a == 1)
                {
                    IWebElement tabla = web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody"));
                    ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
                    for (int b = 1; b <= filas.Count; b++)
                    {
                        IWebElement descarg = web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody/tr[" + b.ToString() + "]/td[19]"));
                        int collec = descarg.FindElements(By.TagName("div")).Count;
                        if (collec > 0)
                        {
                            string[] ArchivosAnteriores = TotalArchivosDownloads();
                            string[] ArchivosNuevos = ArchivosAnteriores;
                            diferentes = 0;
                            do
                            {
                                string nombre = web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody/tr[" + b.ToString() + "]/td[2]")).Text;
                                web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody/tr[" + b.ToString() + "]/td[19]")).Click();
                                Thread.Sleep(300);
                                web.FindElement(By.LinkText("Excel")).Click();
                                Wait("LinkText", "Integración");
                                ArchivosNuevos = ArchivosAnteriores;
                                for (int w = 0; w < 20; w++)
                                {
                                    System.Threading.Thread.Sleep(1000);
                                    ArchivosNuevos = TotalArchivosDownloads();
                                    if (ArchivosNuevos.Length != ArchivosAnteriores.Length)
                                        break;

                                }
                                if (ArchivosNuevos.Length != ArchivosAnteriores.Length)
                                {
                                    diferentes = 1;
                                    string Archivo = NombreArchivoNuevo(ArchivosAnteriores, ArchivosNuevos);
                                    MoverRenombrarArchivo(Archivo, rutaGeneral + @"\Integraciones del " + fechaInicial.ToString("dd MMM yy") + " al " + fechaFinal.ToString("dd MMM yy"), nombre, ".xls");
                                }

                            } while (diferentes == 0);
                        }
                    }
                }
                else
                {
                    web.FindElement(By.LinkText(a.ToString())).Click();
                    Thread.Sleep(3000);
                    IWebElement tabla = web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody"));
                    ReadOnlyCollection<IWebElement> filas = tabla.FindElements(By.TagName("tr"));
                    for (int b = 1; b <= filas.Count; b++)
                    {
                        IWebElement descarg = web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody/tr[" + b.ToString() + "]/td[19]"));
                        int collec = descarg.FindElements(By.TagName("div")).Count;
                        if (collec > 0)
                        {
                            string[] ArchivosAnteriores = TotalArchivosDownloads();
                            string[] ArchivosNuevos = ArchivosAnteriores;
                            diferentes = 0;
                            do
                            {
                                string nombre = web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody/tr[" + b.ToString() + "]/td[2]")).Text;
                                web.FindElement(By.XPath("//*[@id=\"dtFacturas\"]/tbody/tr[" + b.ToString() + "]/td[19]")).Click();
                                Thread.Sleep(500);
                                web.FindElement(By.LinkText("Excel")).Click();
                                //Thread.Sleep(5000);
                                Wait("LinkText", "Integración");
                                //string[] ArchivosNuevos = ArchivosAnteriores;
                                ArchivosNuevos = ArchivosAnteriores;
                                for (int w = 0; w < 20; w++)
                                {
                                    System.Threading.Thread.Sleep(1000);
                                    ArchivosNuevos = TotalArchivosDownloads();
                                    if (ArchivosNuevos.Length != ArchivosAnteriores.Length)
                                        break;

                                }
                                if (ArchivosNuevos.Length != ArchivosAnteriores.Length)
                                {
                                    string Archivo = NombreArchivoNuevo(ArchivosAnteriores, ArchivosNuevos);
                                    diferentes = 1;
                                    MoverRenombrarArchivo(Archivo, rutaGeneral + @"\Integraciones del " + fechaInicial.ToString("dd MMM yy") + " al " + fechaFinal.ToString("dd MMM yy"), nombre, ".xls");
                                }

                            } while (diferentes == 0);
                        }
                    }
                }
            }
        }

        private bool Wait(string tipo, string IDElement)
        {
            bool Seguir = false;
            if (tipo == "Id")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.Id(IDElement));
                while (ChecarElemento.Count == 0)
                {
                    ChecarElemento = driver.FindElements(By.Id(IDElement));
                    System.Threading.Thread.Sleep(4000);
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "LinkText")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.LinkText(IDElement));
                while (ChecarElemento.Count == 0)
                {
                    ChecarElemento = driver.FindElements(By.LinkText(IDElement));
                    System.Threading.Thread.Sleep(4000);
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            if (tipo == "XPath")
            {
                ReadOnlyCollection<IWebElement> ChecarElemento = driver.FindElements(By.XPath(IDElement));
                while (ChecarElemento.Count == 0)
                {
                    ChecarElemento = driver.FindElements(By.XPath(IDElement));
                    System.Threading.Thread.Sleep(4000);
                }
                if (ChecarElemento.Count != 0) { Seguir = true; }
            }
            return (Seguir);
        }

        private string[] TotalArchivosDownloads()
        {
            int i = 0;
            string[] xlsAux = Directory.GetFiles(rutaDescarga, "*.xls");
            foreach (string FileName in xlsAux)
            {
                xlsAux[i] = FileName;
                i++;
            }
            return (xlsAux);
        }

        private string NombreArchivoNuevo(string[] FilesBefore, string[] FilesAfter)
        {
            string Archivo = "";
            foreach (string xl in FilesBefore)
                FilesAfter = Array.FindAll(FilesAfter, s => !s.Equals(xl));
            Archivo = FilesAfter[0].ToString();
            return (Archivo);
        }

        private string MoverRenombrarArchivo(string PathOrigen, string RutaDestino, string NuevoNombre, string Extencion)
        {
            string PathNuevo = "";
            if (!File.Exists(PathOrigen))
            {
                PathNuevo = RutaDestino + @"\" + NuevoNombre + Extencion;
                File.Move(PathOrigen, RutaDestino + @"\" + NuevoNombre + Extencion);
                if (!System.IO.Directory.Exists(RutaDestino))
                    System.IO.Directory.CreateDirectory(RutaDestino);
            }
            if (File.Exists(PathOrigen))
            {
                if (!System.IO.Directory.Exists(RutaDestino))
                    System.IO.Directory.CreateDirectory(RutaDestino);
                if (File.Exists(RutaDestino + @"\" + NuevoNombre + Extencion))
                {
                    PathNuevo = RutaDestino + @"\" + NuevoNombre + " " + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString() + Extencion;
                    File.Move(PathOrigen, PathNuevo);
                }
                if (!File.Exists(RutaDestino + @"\" + NuevoNombre + Extencion))
                {
                    PathNuevo = RutaDestino + @"\" + NuevoNombre + Extencion;
                    File.Move(PathOrigen, PathNuevo);
                }
            }
            return (PathNuevo);
        }

        private string DTaString(DataTable tab, int inicio)
        {
            string texto = string.Empty;
            int conta = 0;
            string[] auxRows = new string[tab.Rows.Count];
            string[] auxColumns = new string[tab.Columns.Count];

            for (int i = inicio; i < tab.Rows.Count; i++)
            {
                for (int j = 0; j < tab.Columns.Count; j++)
                    auxColumns[j] = tab.Rows[i].Field<string>(j).Replace("\r", "").Replace("\n", "");

                auxRows[conta] = String.Join("\t", auxColumns);
                conta++;
            }

            return String.Join(Environment.NewLine, auxRows);
        }

    }

}
