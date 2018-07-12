
namespace Pruebas_clase7.Clases
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Data;
    using Excel = Microsoft.Office.Interop.Excel;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System.Threading;
    using System.Collections.ObjectModel;
    using System.Windows;
    using System.IO;

    class Fragua
    {
        IWebDriver driver;
        Excel.Application MiExcel;//Instancia de Excel
        Excel.Worksheet HojaExcel;
        public static Excel.Range Rango;
        string RutaDescargas = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + @"\Downloads\";
        DataTable PagosFinal = new DataTable();
        DataTable DocumentosFinal = new DataTable();
        DataTable MoviemienotsFinal = new DataTable();
        DataTable ProdDiferenFinal = new DataTable();
        DataTable ResuDiferenFinal = new DataTable();

        public bool funcionPrincipal(String usr, String pass)
        {

            bool exito = false;
            exito = login(usr, pass);
            if (exito)
            {
                string strPagosFinal = "";
                MiExcel = new Excel.Application();
                MiExcel.DisplayAlerts = false;
                MiExcel.Visible = true;
                Excel.Workbooks books = MiExcel.Workbooks;
                Excel.Workbook ArchivoFINAL = books.Add();


                Excel.Worksheet HOJA = (Excel.Worksheet)ArchivoFINAL.Worksheets.Add();
                HOJA.Name = "Estado de Cuenta";
                Rango = HOJA.get_Range("A:Z");
                Rango.Select();
                Rango.NumberFormat = "@";
                string[] archivosdescPAGOS = ObtenerDatosPortal("Estado de cuenta");

                ProcesoExcels(archivosdescPAGOS);

                //SUSTITUIMOS PAGOS POR MOVIMIENTOS
                DataView pagos = DocumentosFinal.DefaultView;
                try { pagos.Sort = "Fecha"; } catch (Exception) { }
                DocumentosFinal = pagos.ToTable();
                strPagosFinal = ConvierteDTaSTRING(DocumentosFinal).Replace("\t\r", "");
                InsertarTitulosDT(HOJA, DocumentosFinal);
                PegarDatosExcel(HOJA, strPagosFinal);

                foreach (Excel.Worksheet item in ArchivoFINAL.Worksheets)
                {
                    item.Select();
                    item.Activate();
                    switch (item.Name.ToString())
                    {
                        case "Estado de Cuenta":
                            EliminaRegnlones(MiExcel, ArchivoFINAL, item, new string[] { "Serie", "Column1" });
                            InsertarTitulos(ArchivoFINAL, item, new string[] { "Serie", "Documento", "F.Docmto", "Saldo", "Status", "F.Pago", "Aut Nota", "Aut Pago", "Origen", "Tipo Docmto" });
                            //Factura
                            string[] Documentos = ObtenerColumna("B2", item, MiExcel);
                            string[] Series = ObtenerColumna("A2", item, MiExcel);
                            for (int i = 0; i < Documentos.Length; i++)
                            {
                                if (Documentos[i].Length == 6)
                                    Documentos[i] = Series[i] + "1" + Documentos[i];
                                if (Documentos[i].Length == 5)
                                    Documentos[i] = Series[i] + "10" + Documentos[i];
                            }
                            string Factura = string.Join(Environment.NewLine, Documentos);
                            clipboardAlmacenaTexto(Factura);
                            do { Thread.Sleep(TimeSpan.FromMilliseconds(10)); } while (!MiExcel.Application.Ready);
                            item.get_Range("C1").EntireColumn.Insert();
                            do { Thread.Sleep(TimeSpan.FromMilliseconds(10)); } while (!MiExcel.Application.Ready);
                            item.get_Range("C2").PasteSpecial();
                            item.Cells["1", "C"] = "Factura";
                            //Fechas
                            CambiarFechasFragua(item, "D2");
                            CambiarFechasFragua(item, "G2");
                            string[] saldo = LeerColumna(1, 5, MiExcel, ArchivoFINAL, item);
                            Rango = item.get_Range("E:E");
                            Rango.Cells.Clear();
                            Rango.NumberFormat = "$#,##0.00";
                            string strSaldo = string.Join(Environment.NewLine, saldo);
                            clipboardAlmacenaTexto(strSaldo);
                            PegarPortaPapelesRango("E1", item);
                            CopiarPegarFormato("D1", "D1", "E1", "E1", item);
                            MarginandoCeldas(ArchivoFINAL, item);
                            break;

                        default:
                            item.Delete();
                            break;
                    }

                }

                GuardandoArchivo(ArchivoFINAL, "Fragua", "Fragua Estado de Cuenta " + DateTime.Now.ToString("dd/mm/yyyy").Replace("/", ""));
                try
                {
                    driver.Close();
                    driver.Quit();
                }
                catch (Exception) { }
                exito = true;

                return exito;
            }
            else
            {
                exito = false;
                return exito;
            }


        }

        private bool login(String usr, String pass)
        {
            bool exito = false;

            ///variables de tipo navegador
            ///
            //•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());
            //•••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••••

            try
            {
                driver.Navigate().GoToUrl("http://interfragua.fragua.com.mx/login.jsp");
            }
            catch (Exception)
            {
                driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromMilliseconds(180));
            }
            if (Wait("Name", "userName", driver))
            {

                Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                driver.FindElement(By.Name("userName")).Clear();
                driver.FindElement(By.Name("userName")).SendKeys(usr);
                driver.FindElement(By.Name("password")).Clear();
                driver.FindElement(By.Name("password")).SendKeys(pass);

                driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromMilliseconds(180));
            }

            else
            {
                login(usr, pass);

                return true;
            }

            try
            {
                driver.FindElement(By.Name("Entrar")).Click();
                Thread.Sleep(TimeSpan.FromMilliseconds(1000));

                exito = true;
            }
            catch (Exception)
            {
                try
                {
                    driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromMilliseconds(180));
                }
                catch (Exception)
                {
                    return false;
                }
            }

            int instancias = driver.WindowHandles.Count;
            int intentos = 0;
            while (instancias < 2 && intentos < 20)
            {
                System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(400));
                instancias = driver.WindowHandles.Count;
                intentos++;
            }

            if (instancias == 2)
            {
                Thread.Sleep(TimeSpan.FromMilliseconds(400));
                driver.SwitchTo().Window(driver.WindowHandles.First()).Close();
                driver.SwitchTo().Window(driver.WindowHandles.Last());
                exito = true;

            }
            else
            {

                return false;
            }


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

        private string[] ObtenerDatosPortal(string Tipo)
        {
            string[] ArchiviosDEsc = new string[100];
            string[] ArchivosANTES = TotalArchivosDownloads(RutaDescargas);

            if (Tipo == "Estado de cuenta")
            {

                driver.SwitchTo().Window(driver.WindowHandles.Last());
                driver.SwitchTo().Frame("topf");




                Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                new SelectElement(driver.FindElement(By.Name("m_proveedor"))).SelectByText("Estado de cuenta");
                Thread.Sleep(TimeSpan.FromMilliseconds(500));
            }

            Thread.Sleep(TimeSpan.FromMilliseconds(500));
            driver.SwitchTo().Window(driver.WindowHandles.Last());
            driver.SwitchTo().Frame("center");

            try
            {
                Thread.Sleep(TimeSpan.FromMilliseconds(1000));

                driver.SwitchTo().Frame(0);
            }
            catch (Exception ex)
            {

            }


            //DSC
            //Obtener el total de documentos de la pagina
            ReadOnlyCollection<IWebElement> numP = driver.FindElements(By.TagName("p"));

            if (numP.Count > 0)
            {
                String[] total = numP.Select(x => x.Text.ToString()).ToArray();
            }


            //DSC



            //NUMERO DE PAGINAS

            try
            {
                int totpag = 0;
                IWebElement paginas = driver.FindElement(By.XPath("/html/body/div[3]/table[1]/tbody/tr/td[2]"));
                string numerosp = paginas.GetAttribute("innerHTML");
                if (numerosp.Contains("font"))
                {

                    ReadOnlyCollection<IWebElement> tot = driver.FindElements(By.LinkText(">|"));
                    if (tot.Count > 0)
                    {
                        driver.FindElement(By.LinkText(">|")).Click();
                        Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        var PAGS = driver.FindElements(By.TagName("a"));
                        Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        string[] hrefs = new string[PAGS.Count];
                        string[] text = new string[PAGS.Count];
                        for (int a = 1; a < PAGS.Count; a++)
                        {
                            text[a] = PAGS[a].Text;
                            hrefs[a] = PAGS[a].GetAttribute("href"); // %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
                            Thread.Sleep(TimeSpan.FromMilliseconds(100));
                            if (text[a] != "" && text[a] != ">|" && text[a] != ">>")
                            {
                                //Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                                string al = PAGS[a].Text;
                                int si;
                                bool seee = Int32.TryParse(al, out si);
                                if (seee == true)
                                {
                                    totpag = Convert.ToInt32(al);
                                }
                            }
                        }
                        Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        driver.FindElement(By.LinkText("|<")).Click();
                        Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                    }
                }


                for (int i = 1; i <= (totpag + 1); i++)
                {
                    if (Wait("xPath", "//*[@id='waterMark']/a/img", driver))
                    {
                        Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                        driver.FindElements(By.TagName("img"))[0].Click();

                        Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        driver.FindElements(By.TagName("img"))[4].Click();

                        Thread.Sleep(TimeSpan.FromMilliseconds(500));
                        driver.SwitchTo().Window(driver.WindowHandles.Last());


                        Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                        try
                        {
                            driver.FindElements(By.Name("tipoArchivo"))[1].Click();
                            Thread.Sleep(TimeSpan.FromMilliseconds(300));
                        }
                        catch (Exception)
                        {
                            driver.FindElements(By.Name("tipoArchivo"))[1].Click();
                            Thread.Sleep(TimeSpan.FromMilliseconds(300));
                        }

                        Thread.Sleep(TimeSpan.FromMilliseconds(400));
                        new SelectElement(driver.FindElement(By.Name("separador"))).SelectByText("Tab");


                        if (Tipo == "Estado de cuenta")
                        {
                            Thread.Sleep(TimeSpan.FromMilliseconds(400));
                            driver.FindElement(By.XPath("/html/body/form/table/tbody/tr[6]/td[2]/input")).Click();
                            Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                        }

                        string[] ArchivosDESPUES2 = ArchivosANTES;
                        for (int w = 0; w < 20; w++)
                        {
                            System.Threading.Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                            ArchivosDESPUES2 = TotalArchivosDownloads(RutaDescargas);
                            if (ArchivosDESPUES2.Length != ArchivosANTES.Length)
                                break;
                        }

                        Thread.Sleep(TimeSpan.FromMilliseconds(1000)); ;
                        driver.SwitchTo().Window(driver.WindowHandles.Last()).Close();
                        ///aqui manda error
                        Thread.Sleep(TimeSpan.FromMilliseconds(1000)); ;
                        driver.SwitchTo().Window(driver.WindowHandles.First());
                        ///x que sera
                        Thread.Sleep(TimeSpan.FromMilliseconds(1000)); ;
                        driver.SwitchTo().Frame("topf");

                        Thread.Sleep(TimeSpan.FromMilliseconds(1000)); ;
                        driver.SwitchTo().Window(driver.WindowHandles.Last());
                        driver.SwitchTo().Frame("center");

                        try
                        {
                            Thread.Sleep(TimeSpan.FromMilliseconds(200)); ;
                            driver.SwitchTo().Frame(0);
                        }
                        catch (Exception ex)
                        {

                        }


                        var siguiente = driver.FindElements(By.TagName("a"));
                        for (int sig = 0; sig < siguiente.Count; sig++)
                            if (siguiente[sig].Text.ToString().Contains(">>"))
                                try
                                {
                                    siguiente[sig].Click();
                                    Thread.Sleep(TimeSpan.FromMilliseconds(500));
                                    break;
                                }
                                catch (Exception)
                                {
                                    driver.Manage().Timeouts().ImplicitWait = (TimeSpan.FromMilliseconds(180));
                                    break;
                                }
                    }
                }
                string[] ArchivosDESPUES = TotalArchivosDownloads(RutaDescargas);
                ArchiviosDEsc = NombresArchivosNuevos(ArchivosANTES, ArchivosDESPUES);

            }
            catch (Exception ex)
            {


            }
            return ArchiviosDEsc;
        }

        private string[] TotalArchivosDownloads(string RutaDescargas)
        {
            int i = 0;
            string[] xlsAux = Directory.GetFiles(RutaDescargas, "*.csv");
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

        #region ••••• E X C E L •••••

        private void ProcesoExcels(string[] Archivos)
        {

            bool bnd = true;
            PagosFinal = new DataTable();
            DocumentosFinal = new DataTable();
            MoviemienotsFinal = new DataTable();
            ProdDiferenFinal = new DataTable();
            ResuDiferenFinal = new DataTable();
            Excel.Workbook ArchivoTrabajoExcel;
            int cont = 1;
            foreach (var item in Archivos)
            {

                cont++;
                //---Creacion del objeto para el uso de Excel
                Excel.Workbooks books = MiExcel.Workbooks;                                              //---Creacion de objeto para el uso de hoja de trabajo de excel
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); ; } while (!MiExcel.Application.Ready);
                ArchivoTrabajoExcel = books.Open(item);    //---Ruta y nombre del archivo                                                               //---Muestra el proceso de Excel
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
                ((Excel.Worksheet)MiExcel.ActiveWorkbook.Sheets[1]).Select();                   //---Nombre o numero de la hoja activa
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
                HojaExcel = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
                string[] DatosA = ObtenerColumna("A1", HojaExcel, MiExcel);

                int PagosINI = 0, DocsINI = 0, MovsINI = 0, ProdDife = 0, ResDife = 0, Devol = 0, DetalDevo = 0;
                for (int i = 0; i < DatosA.Length; i++)
                {
                    if (DatosA[i].Contains("Pagos"))
                        PagosINI = i + 1;
                    if (DatosA[i].Contains("Documentos"))
                        DocsINI = i + 1;
                    if (DatosA[i].Contains("Movimientos"))
                        MovsINI = i + 1;
                    if (DatosA[i].Contains("Productos con Diferencias"))
                        ProdDife = i + 1;
                    if (DatosA[i].Contains("Resumen de Diferencias"))
                        ResDife = i + 1;
                }

                DataTable Pagos = new DataTable();
                DataTable Documentos = new DataTable();
                DataTable Moviemienots = new DataTable();
                DataTable ProdDiferen = new DataTable();
                DataTable ResuDiferen = new DataTable();
                if (DocsINI != 0)
                    Documentos = ObtenerDatosExcel(MiExcel, ArchivoTrabajoExcel, DocsINI + 1);
                if (MovsINI != 0)
                    Moviemienots = ObtenerDatosExcel(MiExcel, ArchivoTrabajoExcel, MovsINI + 1);
                if (ProdDife != 0)
                    ProdDiferen = ObtenerDatosExcel(MiExcel, ArchivoTrabajoExcel, ProdDife + 1);
                if (ResDife != 0)
                    ResuDiferen = ObtenerDatosExcel(MiExcel, ArchivoTrabajoExcel, ResDife + 1);

                if (bnd)
                {
                    bnd = false;
                    PagosFinal = Pagos.Clone();
                    DocumentosFinal = Documentos.Clone();
                    MoviemienotsFinal = Moviemienots.Clone();
                    ProdDiferenFinal = ProdDiferen.Clone();
                    ResuDiferenFinal = ResuDiferen.Clone();
                }

                PagosFinal.Merge(Pagos);//PAGOS
                DocumentosFinal.Merge(Documentos);//PAGOS2
                MoviemienotsFinal.Merge(Moviemienots);
                ProdDiferenFinal.Merge(ProdDiferen);// DEVO
                ResuDiferenFinal.Merge(ResuDiferen);//DETALLE DEVOLUCIONES
                do { Thread.Sleep(TimeSpan.FromMilliseconds(10)); } while (!MiExcel.Application.Ready);
                ArchivoTrabajoExcel.Close();
                File.Delete(item);
            }

            //Pegando DT a Excel Final
        }

        private void AcumulaArchivosUnaPestaña(string[] Archivos, Excel.Worksheet SheetExcel)
        {
            foreach (var item in Archivos)
            {
                //Abriendo archivos
                Excel.Workbooks books = MiExcel.Workbooks;
                Excel.Workbook ArchivoTrabajoExcel = books.Open(item);
                Excel.Worksheet HOJA = (Excel.Worksheet)ArchivoTrabajoExcel.ActiveSheet;
                HOJA.Activate();
                //Copiando datos
                string datos = CopiaDatosExcel(HOJA);
                ArchivoTrabajoExcel.Close();
                File.Delete(item);
                //Pegando datos
                SheetExcel.Activate();
                SheetExcel.Cells.NumberFormat = "@";
                PegarDatosExcel(SheetExcel, datos);
            }
        }

        private string CopiaDatosExcel(Excel.Worksheet HOJA)
        {
            string datos = "";
            HOJA.Activate();
            Excel.Range Renglon = HOJA.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int row = Convert.ToInt32(Renglon.Row.ToString());
            int col = Convert.ToInt32(Renglon.Column.ToString());
            string CellIni = "A" + 1;
            string CellFin = ColumnaCorrespondiente(col) + row;
            Excel.Range Rango = HOJA.get_Range(CellIni, CellFin);
            Rango.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            Rango.Copy();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            datos = clipboardObtenerTexto();
            return datos;
        }

        private String[] ObtenerColumna(String iniColumna, Excel.Worksheet SheetExcel, Excel.Application appExcel)
        {
            // Clipboard.Clear();

            String c = "";
            String r = "";

            if (iniColumna.Length == 2)
            {
                c = Convert.ToString(iniColumna[0]);
                r = Convert.ToString(iniColumna[1]);
            }
            else
            {
                if (iniColumna.Length == 3)
                {
                    c = Convert.ToString(iniColumna[0]) + Convert.ToString(iniColumna[1]);
                    r = Convert.ToString(iniColumna[2]);
                }
            }

            Excel.Range Rango;
            String[] Col;
            int row;

            Rango = SheetExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            row = Convert.ToInt32(Rango.Row.ToString());
            Rango = SheetExcel.get_Range(iniColumna, c + row);
            Rango.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!appExcel.Application.Ready);
            Rango.Copy();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!appExcel.Application.Ready);
            //String datos = Clipboard.GetText();
            String datos = clipboardObtenerTexto();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!appExcel.Application.Ready);
            datos = datos.Replace("\r", "");
            Col = datos.Split('\n');
            return Col;
        }

        private DataTable ObtenerDatosExcel(Excel.Application APPEXC, Excel.Workbook ARCHIVO, int filaTitulos)
        {
            DataTable DT = new DataTable();
            ///---Inicio
            ((Excel.Worksheet)APPEXC.ActiveWorkbook.Sheets[1]).Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            Excel.Worksheet HojaExcelLeer = (Excel.Worksheet)ARCHIVO.ActiveSheet;
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            string ColIni = ColumnaCorrespondiente(1);
            int ColFin = UltimaColumna(ARCHIVO, filaTitulos.ToString());
            ///---Verificacion de Titulos
            int totaltitulos = ColFin;
            string CellIni = ColIni + (filaTitulos).ToString();
            string CellFin = ColumnaCorrespondiente(ColFin) + filaTitulos;
            Rango = HojaExcelLeer.get_Range(CellIni, CellFin);
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            Rango.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            Rango.Copy();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            string portapapeles = clipboardObtenerTexto();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            string[] TITULOS = portapapeles.Split('\t');
            ///---Creacion de estructura del DT
            int tit = 2;
            for (int x = 0; x < totaltitulos; x++)
            {
                try
                {
                    DT.Columns.Add(TITULOS[x].ToString().Trim(), typeof(string));
                }
                catch (Exception)
                {
                    //Dispatcher.Invoke(((Action)(() => objNu4.ReportarLog(RUTA_ARCHIVO_LOG, x.ToString())))); ///%%%%%AQUI
                    DT.Columns.Add(TITULOS[x].ToString().Trim() + (tit++).ToString(), typeof(string));
                }
            }
            ///---Gauradndo en clipboard las celdas seleccionadas
            Rango = Rango.End[Excel.XlDirection.xlDown];
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
            int UltRow = Convert.ToInt32(Rango.Row.ToString());
            if (UltRow != 1048576)
            {
                CellIni = ColIni + (filaTitulos + 1).ToString();
                CellFin = ColumnaCorrespondiente(ColFin) + UltRow;
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
                Rango = HojaExcelLeer.get_Range(CellIni, CellFin);
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
                Rango.Select();
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
                Rango.Copy();
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
                portapapeles = clipboardObtenerTexto();
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!APPEXC.Application.Ready);
                if (portapapeles != "")
                {
                    string[] DATOS_col = portapapeles.Split('\t', '\n');
                    ///---Copiando de datos de DT
                    int NumRow = APPEXC.Selection.Rows.Count;
                    int NumCol = APPEXC.Selection.Columns.Count;
                    int aux = 0, aux2 = 0;
                    for (int row = 0; row < NumRow; row++)
                    {
                        DT.Rows.Add();
                        for (int col = 0; col < NumCol; col++)
                        {
                            DT.Rows[row][col] = DATOS_col[col + aux];
                            aux2 = col;
                        }
                        aux += aux2 + 1;
                    }
                }
                else
                {
                    DT.Rows.Add();
                }
            }
            return (DT);
        }

        private void CambiarFechasFragua(Excel.Worksheet item, string Celda)
        {
            string[] APagosFECHAS = ObtenerColumna(Celda, item, MiExcel);
            APagosFECHAS = CambiarFormatoFechas(APagosFECHAS);
            string FechaPagos = string.Join(Environment.NewLine, APagosFECHAS);
            clipboardAlmacenaTexto(FechaPagos);
            Thread.Sleep(TimeSpan.FromMilliseconds(500));
            item.get_Range(Celda).EntireColumn.NumberFormat = "@";
            Thread.Sleep(TimeSpan.FromMilliseconds(500));
            item.get_Range(Celda).PasteSpecial();
            Thread.Sleep(TimeSpan.FromMilliseconds(500));
            item.get_Range(Celda).EntireColumn.NumberFormat = "m/d/yyyy";
            Thread.Sleep(TimeSpan.FromMilliseconds(500));
            Excel.Range rangoFecha = item.get_Range(Celda).EntireColumn;
            rangoFecha.Replace("-", "/", Excel.XlLookAt.xlPart, Excel.XlSearchOrder.xlByRows, false, false, false);
        }

        private string[] CambiarFormatoFechas(string[] Fechas)
        {
            for (int i = 0; i < Fechas.Length; i++)
            {
                try
                {
                    //
                    if (Fechas[i].Contains("ene"))
                        Fechas[i] = Fechas[i].Replace("ene", "01").Replace("-", "/");
                    if (Fechas[i].Contains("feb"))
                        Fechas[i] = Fechas[i].Replace("feb", "02").Replace("-", "/");
                    if (Fechas[i].Contains("mar"))
                        Fechas[i] = Fechas[i].Replace("mar", "03").Replace("-", "/");
                    if (Fechas[i].Contains("abr"))
                        Fechas[i] = Fechas[i].Replace("abr", "04").Replace("-", "/");
                    if (Fechas[i].Contains("may"))
                        Fechas[i] = Fechas[i].Replace("may", "05").Replace("-", "/");
                    if (Fechas[i].Contains("jun"))
                        Fechas[i] = Fechas[i].Replace("jun", "06").Replace("-", "/");
                    if (Fechas[i].Contains("jul"))
                        Fechas[i] = Fechas[i].Replace("jul", "07").Replace("-", "/");
                    if (Fechas[i].Contains("ago"))
                        Fechas[i] = Fechas[i].Replace("ago", "08").Replace("-", "/");
                    if (Fechas[i].Contains("sep"))
                        Fechas[i] = Fechas[i].Replace("sep", "09").Replace("-", "/");
                    if (Fechas[i].Contains("oct"))
                        Fechas[i] = Fechas[i].Replace("oct", "10").Replace("-", "/");
                    if (Fechas[i].Contains("nov"))
                        Fechas[i] = Fechas[i].Replace("nov", "11").Replace("-", "/");
                    if (Fechas[i].Contains("dic"))
                        Fechas[i] = Fechas[i].Replace("dic", "12").Replace("-", "/");
                    //
                    string[] todo = Fechas[i].Split('/');
                    switch (todo[2])
                    {
                        case "10": todo[2] = "2010"; break;
                        case "11": todo[2] = "2011"; break;
                        case "12": todo[2] = "2012"; break;
                        case "13": todo[2] = "2013"; break;
                        case "14": todo[2] = "2014"; break;
                        case "15": todo[2] = "2015"; break;
                        case "16": todo[2] = "2016"; break;
                        case "17": todo[2] = "2017"; break;
                        case "18": todo[2] = "2018"; break;
                        case "19": todo[2] = "2019"; break;
                        case "20": todo[2] = "2020"; break;
                        default:
                            break;
                    }
                    Fechas[i] = todo[0] + "/" + todo[1] + "/" + todo[2];
                    Fechas[i] = Fechas[i].Replace("/", "-");
                }
                catch (Exception)
                {

                }
            }

            return Fechas;
        }

        private DataTable ObtenerDatosExcel_2(Excel.Worksheet HOJA, string ColIni, int filaTitulos, string ColFin, int UltimoRenglon)
        {
            DataTable DT = new DataTable();
            ///---Inicio
            Rango = HOJA.get_Range(ColIni + filaTitulos, ColFin + filaTitulos);
            Rango.Select();
            Rango.Copy();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            string portapapeles = clipboardObtenerTexto();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            string[] TITULOS = portapapeles.Split('\t');
            ///---Creacion de estructura del DT
            int tit = 2;
            for (int x = 0; x < TITULOS.Length; x++)
            {
                try
                {
                    DT.Columns.Add(TITULOS[x].ToString().Trim(), typeof(string));
                }
                catch (Exception)
                {
                    DT.Columns.Add(TITULOS[x].ToString().Trim() + (tit++).ToString(), typeof(string));
                }
            }
            ///---Gauradndo en clipboard las celdas seleccionadas
            if (UltimoRenglon != 1048576)
            {
                string CellIni = ColIni + (filaTitulos + 1).ToString();
                string CellFin = ColFin + UltimoRenglon;
                Rango = HOJA.get_Range(CellIni, CellFin);
                Rango.Select();
                Rango.Copy();
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
                portapapeles = clipboardObtenerTexto();
                do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
                if (portapapeles != "")
                {
                    string[] DATOS_col = portapapeles.Split('\t', '\n');
                    ///---Copiando de datos de DT
                    int NumRow = UltimoRenglon;
                    int NumCol = TITULOS.Length;
                    int aux = 0, aux2 = 0;
                    for (int row = 0; row < NumRow; row++)
                    {
                        DT.Rows.Add();
                        try
                        {
                            for (int col = 0; col < NumCol; col++)
                            {
                                DT.Rows[row][col] = DATOS_col[col + aux];
                                aux2 = col;
                            }
                            aux += aux2 + 1;
                        }
                        catch (Exception)
                        {
                            DT.Rows[row].Delete();
                            break;
                        }
                    }
                }
                else
                {
                    DT.Rows.Add();
                }
            }
            return (DT);
        }

        private DataTable CopiaDatosExcel_DataTable(Excel.Worksheet HOJA, int FilaTitulos)
        {
            DataTable Datos = new DataTable();
            HOJA.Activate();
            Excel.Range Renglon = HOJA.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int row = Convert.ToInt32(Renglon.Row.ToString());
            int col = Convert.ToInt32(Renglon.Column.ToString());
            string CellIni = "A" + FilaTitulos;
            string CellFin = ColumnaCorrespondiente(col) + row;
            Excel.Range Rango = HOJA.get_Range(CellIni, CellFin);
            Rango.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            Rango.Copy();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            //Obtencion
            string strdatos = clipboardObtenerTexto();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            string[] Renglones = strdatos.Split('\n');
            //Titulos
            string[] Titulos = Renglones[0].Split('\t');
            for (int T = 0; T < Titulos.Length; T++)
                if (Titulos[T].ToString() != "" && Titulos[T].ToString() != " " && Titulos[T].ToString() != null)
                    Datos.Columns.Add(Titulos[T].ToString());
            //Datos
            List<string> renglones = Renglones.ToList();
            renglones.RemoveAt(0);
            Renglones = renglones.ToArray();
            for (int i = 0; i < Renglones.Length; i++)
            {
                Datos.Rows.Add();
                string[] Columnas = Renglones[i].Split('\t');
                for (int j = 0; j < Columnas.Length; j++)
                    Datos.Rows[i][j] = Columnas[j];
            }
            return Datos;
        }

        private void PegarDatosExcel(Excel.Worksheet HOJA, string Datos)
        {
            Excel.Range Renglon = HOJA.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int row = Convert.ToInt32(Renglon.Row.ToString()) + 2;
            Excel.Range CeldaPegar = HOJA.get_Range("A" + row);
            CeldaPegar.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            clipboardAlmacenaTexto(Datos);
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);

            PegarPortaPapelesRango("A" + (row + 1).ToString(), HOJA);
        }

        private void PegarDatosExcel_2(Excel.Worksheet HOJA, string Celda, string Datos)
        {
            Excel.Range CeldaPegar = HOJA.get_Range(Celda);
            CeldaPegar.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            clipboardAlmacenaTexto(Datos);
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            //CeldaPegar.PasteSpecial();
            PegarPortaPapelesRango(Celda, HOJA);
        }

        private void InsertarTitulosDT(Excel.Worksheet HOJA, DataTable DT)
        {
            HOJA.Select();
            HOJA.Activate();
            //Obtener los titulos
            string Titulos = string.Empty;
            for (int i = 0; i < DT.Columns.Count; i++)
                Titulos += DT.Columns[i].ColumnName.ToString() + "\t";
            //Pegar en el ultimo renglon de excel + 1 extra
            Excel.Range Renglon = HOJA.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            int row = Convert.ToInt32(Renglon.Row.ToString()) + 1;
            Excel.Range CeldaPegar = HOJA.get_Range("A" + (row + 1).ToString());
            CeldaPegar.Select();
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            clipboardAlmacenaTexto(Titulos);
            do { Thread.Sleep(TimeSpan.FromMilliseconds(20)); } while (!MiExcel.Application.Ready);
            //CeldaPegar.PasteSpecial();
            PegarPortaPapelesRango("A" + (row + 1).ToString(), HOJA);
        }

        private void InsertarTitulos(Excel.Workbook ARCHIVO, Excel.Worksheet HOJA, string[] Titulos)
        {
            HOJA.Select();
            HOJA.Activate();
            //---Insertando renglon
            Excel.Range Rango = HOJA.get_Range("A1");
            Rango.EntireRow.Insert();
            //---Poniendo uno nuevos titulos
            int c = 0;
            string col = "";
            foreach (var item in Titulos)
            {
                c++;
                col = ColumnaCorrespondiente(c);
                HOJA.Cells["1", col] = item;
            }
            //---Cambiando formato de los datos obtenidos
            Rango = HOJA.get_Range("A1", col + "1");
            Rango.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
            Rango.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            Rango.EntireRow.Font.Bold = true;
        }

        #endregion

        private bool GuardandoArchivo(Excel.Workbook ARCHIVO, string Cliente, string Nombre)
        {
            bool exito = false;
            string RutaCarpeta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\Archivos Generados\";
            string NombreExtra = "";
            //Guardando Archivo
            if (!System.IO.Directory.Exists(RutaCarpeta))
                System.IO.Directory.CreateDirectory(RutaCarpeta);
            Nombre = Nombre.Replace("-", "").Replace("/", "");
            string Archivo = RutaCarpeta + Nombre + ".xlsx";
            if (File.Exists(Archivo))
            {
                Archivo = Archivo.Replace(".xlsx", "") + " " + DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                Archivo += ".xlsx";
            }
            try
            {
                ArchivoGuardarComo(ARCHIVO, Archivo);

            }
            catch (Exception ex)
            {
                try
                {
                    ARCHIVO.SaveAs(Archivo);
                }
                catch (Exception ey)
                {

                }
            }
            return exito;
        }

        private void ArchivoGuardarComo(Excel.Workbook Archivo, string RutaNombreArchivo)
        {
            Archivo.SaveAs(RutaNombreArchivo);
        }

        private string clipboardObtenerTexto()
        {
            string clipboard = "Sin valor...";
            Thread staThread = new Thread(x =>
            {
                try
                {
                    clipboard = Clipboard.GetText();
                }
                catch (Exception ex)
                {
                    clipboard = ex.Message;
                }
            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
            return clipboard;
        }

        private void clipboardAlmacenaTexto(string valor)
        {
            Thread staThread = new Thread(x =>
            {
                try
                {
                    Clipboard.SetText(valor);
                }
                catch (Exception e)
                {
                }

            });
            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();
        }

        private void PegarPortaPapelesRango(string CeldaPegar, Excel.Worksheet Pestania)
        {
            Excel.Range unRango;
            unRango = Pestania.get_Range(CeldaPegar);
            unRango.Select();
            Pestania.Paste();
        }

        private String ColumnaCorrespondiente(int num)
        {
            string columna = "";
            switch (num)
            {
                case 1: columna = "A"; break;
                case 2: columna = "B"; break;
                case 3: columna = "C"; break;
                case 4: columna = "D"; break;
                case 5: columna = "E"; break;
                case 6: columna = "F"; break;
                case 7: columna = "G"; break;
                case 8: columna = "H"; break;
                case 9: columna = "I"; break;
                case 10: columna = "J"; break;
                case 11: columna = "K"; break;
                case 12: columna = "L"; break;
                case 13: columna = "M"; break;
                case 14: columna = "N"; break;
                case 15: columna = "O"; break;
                case 16: columna = "P"; break;
                case 17: columna = "Q"; break;
                case 18: columna = "R"; break;
                case 19: columna = "S"; break;
                case 20: columna = "T"; break;
                case 21: columna = "U"; break;
                case 22: columna = "V"; break; ;
                case 23: columna = "W"; break;
                case 24: columna = "X"; break;
                case 25: columna = "Y"; break;
                case 26: columna = "Z"; break;
                case 27: columna = "AA"; break;
                case 28: columna = "AB"; break;
                case 29: columna = "AC"; break;
                case 30: columna = "AD"; break;
                case 31: columna = "AE"; break;
                case 32: columna = "AF"; break;
                case 33: columna = "AG"; break;
                case 34: columna = "AH"; break;
                case 35: columna = "AI"; break;
                case 36: columna = "AJ"; break;
                case 37: columna = "AK"; break;
                case 38: columna = "AL"; break;
                case 39: columna = "AM"; break;
                case 40: columna = "AN"; break;
                case 41: columna = "AO"; break;
                case 42: columna = "AP"; break;
                case 43: columna = "AQ"; break;
                case 44: columna = "AR"; break;
                case 45: columna = "AS"; break;
                case 46: columna = "AT"; break;
                case 47: columna = "AU"; break;
                case 48: columna = "AV"; break;
                case 49: columna = "AW"; break;
                case 50: columna = "AX"; break;
                case 51: columna = "AY"; break;
                case 52: columna = "AZ"; break;
                case 53: columna = "BA"; break;
                case 54: columna = "BB"; break;
                case 55: columna = "BC"; break;
                case 56: columna = "BD"; break;
                case 57: columna = "BE"; break;
                case 58: columna = "BF"; break;
                case 59: columna = "BG"; break;
                case 60: columna = "BH"; break;
                case 61: columna = "BI"; break;
                case 62: columna = "BJ"; break;
                case 63: columna = "BK"; break;
                case 64: columna = "BL"; break;
                case 65: columna = "BM"; break;
                case 66: columna = "BN"; break;
                case 67: columna = "BO"; break;
                case 68: columna = "BP"; break;
                case 69: columna = "BQ"; break;
                case 70: columna = "BR"; break;
                case 71: columna = "BS"; break;
                case 72: columna = "BT"; break;
                case 73: columna = "BU"; break;
                case 74: columna = "BV"; break;
                case 75: columna = "BW"; break;
                case 76: columna = "BX"; break;
                case 77: columna = "BY"; break;
                case 78: columna = "BZ"; break;
                case 79: columna = "CA"; break;
                case 80: columna = "CB"; break;
                case 81: columna = "CC"; break;
                case 82: columna = "CD"; break;
                case 83: columna = "CE"; break;
                case 84: columna = "CF"; break;
                case 85: columna = "CG"; break;
                case 86: columna = "CH"; break;
                case 87: columna = "CI"; break;
                case 88: columna = "CJ"; break;
                case 89: columna = "CK"; break;
                case 90: columna = "CL"; break;
                case 91: columna = "CM"; break;
                case 92: columna = "CN"; break;
                case 93: columna = "CO"; break;
                case 94: columna = "CP"; break;
                case 95: columna = "CQ"; break;
                case 96: columna = "CR"; break;
                case 97: columna = "CS"; break;
                case 98: columna = "CT"; break;
                case 99: columna = "CU"; break;
                case 100: columna = "CV"; break;
                case 101: columna = "CW"; break;
                case 102: columna = "CX"; break;
                case 103: columna = "CY"; break;
                case 104: columna = "CZ"; break;
                case 105: columna = "DA"; break;
                default: columna = ""; break;
            }
            return (columna);
        }

        private int UltimaColumna(Excel.Worksheet HOJA)
        {
            int col = 0;
            Excel.Range Renglon = HOJA.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            col = Convert.ToInt32(Renglon.Column.ToString());
            return col;
        }

        private int UltimaColumna(Excel.Workbook booksExcel, String valRow)
        {
            int icolUltimo;
            String celda;

            Excel.Worksheet sheetExcel;
            Excel.Range rangoExcel, colUltimo;
            try
            {
                celda = "XFD" + valRow;
                booksExcel.Application.ActiveWorkbook.ActiveSheet.Range(celda).select();

            }
            catch
            {
                celda = "IV" + valRow;
                booksExcel.Application.ActiveWorkbook.ActiveSheet.Range(celda).select();
            }


            sheetExcel = (Excel.Worksheet)booksExcel.ActiveSheet;
            rangoExcel = sheetExcel.get_Range(celda);
            //rangoExcel = sheetExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //rangoExcel = rangoExcel.End[Excel.XlDirection.xlToRight];            
            colUltimo = rangoExcel.End[Excel.XlDirection.xlToLeft];
            icolUltimo = colUltimo.Column;
            booksExcel.Application.ActiveWorkbook.ActiveSheet.Range("A1").select();
            return icolUltimo;
        }

        private string ConvierteDTaSTRING(DataTable Tabla)
        {
            string contenido = "";
            int ro = Tabla.Columns.Count;
            foreach (DataRow i in Tabla.Rows)
            {
                if (i.RowState != DataRowState.Deleted)
                {
                    for (int e = 0; e < ro; e++)
                    {
                        contenido += i[e] + "\t";
                    }
                    contenido += Environment.NewLine;
                }
            }
            return (contenido);
        }

        private void EliminaRegnlones(Excel.Application APPEXC, Excel.Workbook LIBRO, Excel.Worksheet HOJA, string[] Contenido)
        {
            HOJA.Select();
            HOJA.Activate();
            Excel.Range Rango;
            string[] colAFac = LeerColumna(1, 1, APPEXC, LIBRO, HOJA);
            //Agregando opciones
            List<string> lista = new List<string>();
            lista = Contenido.ToList<string>();
            lista.Add("\r\n");
            lista.Add("\r");
            lista.Add("\n");
            int r = 1;
            foreach (var item in colAFac)
            {
                string Renglon = item.ToString();
                foreach (var cont in lista)
                {
                    if (Renglon.Contains(cont) || Renglon.StartsWith(cont) || Renglon.StartsWith(" ") || Renglon.Equals(""))
                    {
                        try
                        {
                            Rango = HOJA.get_Range("A" + r);
                            Rango.EntireRow.Delete();
                        }
                        catch
                        {
                            HOJA.Cells[r, "A"].EntireRow.Delete();
                        }
                        r--;
                        break;
                    }
                }
                r++;
            }
        }

        private string[] LeerColumna(int renglonIni, int columnaIni, Excel.Application AppExcel, Excel.Workbook Archivo, Excel.Worksheet Pestania)
        {
            string rngCompleto, elemento, primero, ult, col;
            int i, j, NumRow, ultimo;
            char[] delimiterChars = { '\r', '\n' };

            Excel.Range Ran;
            col = ColumnaCorrespondiente(columnaIni);//Obtiene la letra de la primera columna a seleccionar
            ultimo = UltimoRenglon(Archivo, col);
            primero = col + renglonIni;
            ult = col + ultimo;
            Ran = Pestania.get_Range(primero, ult);//Define Rango
            Ran.Select();//Lo selecciona
            NumRow = AppExcel.Selection.Rows.Count; //Se obtuvo el numero de filas seleccionadas.
            string[] Datos1, Datos2 = new string[NumRow];
            if (ultimo >= renglonIni)
            {
                Ran.Copy();
                //rngCompleto = Clipboard.GetText();//Se pega como un string en Clipboard
                rngCompleto = clipboardObtenerTexto();
                Datos1 = rngCompleto.Split(delimiterChars);
                i = 0;
                for (j = 0; j < Datos1.Length - 1; j++)//debe ser -1 por que por el formato de los saltos quedan los ultimos 2//elementos del arreglo Datos1 vacios 
                {
                    if (j % 2 == 0 || j == 0)//filtra los elementos en posicion de nums nones que estan vacios
                    {
                        elemento = Convert.ToString(Datos1[j]);
                        Datos2[i] = elemento;//los añade a otro arreglo que será el devuelto
                        i++;
                    }
                }
            }
            else
            {
                for (i = 0; i < NumRow; i++) { Datos2[i] = ""; }
            }
            return (Datos2);
        }

        private int UltimoRenglon(Excel.Workbook booksExcel, String valColumna)
        {
            int irowUltimo;
            String celda;
            Excel.Worksheet sheetExcel;
            Excel.Range rangoExcel, rngUltimo;
            try
            {
                celda = valColumna + "1048575";
                booksExcel.Application.ActiveWorkbook.ActiveSheet.Range(celda).select();
            }
            catch (Exception)
            {
                celda = valColumna + "65536";
                booksExcel.Application.ActiveWorkbook.ActiveSheet.Range(celda).select();
            }

            sheetExcel = (Excel.Worksheet)booksExcel.ActiveSheet;
            rangoExcel = sheetExcel.get_Range(celda);
            //rangoExcel = sheetExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //rangoExcel = rangoExcel.End[Excel.XlDirection.xlDown];
            rngUltimo = rangoExcel.End[Excel.XlDirection.xlUp];
            irowUltimo = rngUltimo.Row;
            booksExcel.Application.ActiveWorkbook.ActiveSheet.Range("A1").select();
            return irowUltimo;
        }

        private int UltimoRenglon(Excel.Worksheet HOJA)
        {
            int row = 0;
            Excel.Range Renglon = HOJA.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            row = Convert.ToInt32(Renglon.Row.ToString());
            return row;
        }

        private void CopiarPegarFormato(string CelIniCopFor, string CelFinCopFor, string CelIniPegFor, string CelFinPegFor, Excel.Worksheet Pestania)
        {
            Excel.Range unRango;
            unRango = Pestania.get_Range(CelIniCopFor, CelFinCopFor);
            unRango.Select();
            unRango.Copy();
            unRango = Pestania.get_Range(CelIniPegFor, CelFinPegFor);
            unRango.Select();
            unRango.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
        }

        private void MarginandoCeldas(Excel.Workbook ARCHIVO, Excel.Worksheet HOJA)
        {
            HOJA.Select();
            HOJA.Activate();
            //Variables
            int RenglonTitulosDatos = 1;
            string ColIniStr = ColumnaCorrespondiente(1);
            int RenFin = UltimoRenglon(HOJA);
            int ColFin = UltimaColumna(HOJA);
            string ColFinStr = ColumnaCorrespondiente(ColFin);
            string CellIni = "A1";
            string CellFin = ColFinStr + RenFin.ToString();
            //Marginando
            Excel.Range Rango = HOJA.get_Range(CellIni, CellFin);
            //Rango.ReadingOrder = (int)Excel.Constants.xlContext;
            Rango.HorizontalAlignment = 3;
            Rango.VerticalAlignment = 2;
            Rango.EntireColumn.AutoFit();
            Rango.EntireRow.AutoFit();
            Excel.Borders border = Rango.Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 1d;
        }

    }
}