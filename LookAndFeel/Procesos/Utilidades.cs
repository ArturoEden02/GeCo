namespace Pruebas_clase7.Clases
{
    using Nu4it;
    using nu4itExcel;
    using nu4itFox;
    using System;
    using System.Data;
    using System.IO;
    using System.Linq;
    using System.Net.NetworkInformation;
    using System.Threading;
    using Excel = Microsoft.Office.Interop.Excel;
    public class Utilidades
    {
        const int NO = 0;
        const int SI = 1;

        private static usaR objNu4 = new usaR();
        private static nufox objNuFox = new nufox();
        private static nuExcel objNuExcel = new nuExcel();

        public static string DTaString(DataTable dataTable, int inicio)
        {
            string texto = string.Empty;
            int conta = 0;
            string[] auxRows = new string[dataTable.Rows.Count];
            string[] auxColumns = new string[dataTable.Columns.Count];

            for (int i = inicio; i < dataTable.Rows.Count; i++)
            {
                for (int j = 0; j < dataTable.Columns.Count; j++)
                    auxColumns[j] = dataTable.Rows[i].Field<string>(j).Replace("\r", "").Replace("\n", "");

                auxRows[conta] = String.Join("\t", auxColumns);
                conta++;
            }

            return String.Join(Environment.NewLine, auxRows);
        }

        //public Excel.Workbook AbrirArchivoExcel(Excel.Application appExcel, Excel.Workbook libroExcel, string mensaje)
        //{
        //    Inicio:
        //    string rutaArchivo = FileDialog(mensaje, "Excel");
        //    if (!string.IsNullOrEmpty(rutaArchivo))
        //    {
        //        objNuExcel.InstanciaExcelVisible(appExcel);
        //        objNuExcel.ActivarMensajesAlertas(appExcel, NO);

        //        libroExcel = objNuExcel.AbrirArchivo(rutaArchivo, appExcel);
        //        objNuExcel.ActivarArchivo(libroExcel);
        //    }
        //    else
        //        goto Inicio;

        //    return libroExcel;
        //}

        //public string FileDialog(string mensaje, string TipoArchivo)
        //{
        //    string FilePath = String.Empty;
        //    string FiltroArchivo = String.Empty;
        //    MessageShowOK(mensaje);
        //    Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog
        //    {
        //        Title = mensaje
        //    };
        //    switch (TipoArchivo)
        //    {
        //        case "Excel":
        //            FiltroArchivo = "Excel Files|*.xls;*.xlsx;*.xlsb;*.xlsm;*.xlsb";
        //            break;
        //        case "txt":
        //            FiltroArchivo = "Txt Files|*.txt";
        //            break;
        //        case "pdf":
        //            FiltroArchivo = "PDF Files|*.pdf";
        //            break;
        //    }
        //    dialog.Filter = FiltroArchivo;
        //    Nullable<bool> result = dialog.ShowDialog();
        //    if (result == true)
        //        FilePath = dialog.FileName;
        //    return FilePath;
        //}

        public static string[] ObtenerColumna(Excel.Application appExcel, Excel.Worksheet hojaExcel, string celda, int fila)
        {
            Excel.Range Rango;
            string[] col;
            int row;
            do { Thread.Sleep(250); } while (!appExcel.Application.Ready);
            Rango = hojaExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            row = Convert.ToInt32(Rango.Row.ToString());
            Thread.Sleep(1000);
            Rango = hojaExcel.get_Range(celda + fila, celda + row);
            Rango.Select();
            Rango.Copy();

            string datos = objNu4.clipboardObtenerTexto();
            datos = datos.Replace("\r", "");
            col = datos.Split('\n');
            return col;
        }

        private static string Celda(Excel.Workbook libroExcel, Excel.Worksheet hojaExcel, int fila, string titulo1, string titulo2)
        {
            string celda;
            Excel.Range Rango;
            if (fila > 0)
            {
                Rango = hojaExcel.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                string row = Rango.Row.ToString();
                Thread.Sleep(500);
                Rango = hojaExcel.get_Range("A" + fila, "BN" + fila);
                Rango.Copy();
                string datos = objNu4.clipboardObtenerTexto();
                string[] titulos = datos.ToUpper().Split('\t');
                string[] datosLimpios = (from x in titulos select x.Replace("\r", "").Replace("\n", "").Trim()).ToArray();
                int columna = Array.IndexOf(datosLimpios, titulo1.ToUpper());

                celda = objNuExcel.ColumnaCorrespondiente(columna + 1);
            }
            else
                celda = "null";

            return celda;
        }

        public static string[] ColumnaCelda(Excel.Application appExcel, Excel.Workbook libroExcel, Excel.Worksheet hojaExcel, string[] titulos, ref int fila)
        {
            string[] colTitulos = new string[titulos.Length];
            do { Thread.Sleep(250); } while (!appExcel.Application.Ready);

            if (titulos.Length > 1)
                fila = objNuExcel.filaTitulos_2(libroExcel, titulos[0], titulos[1]);
            else
                fila = objNuExcel.filaTitulos_2(libroExcel, titulos[0], titulos[0]);


            for (int i = 0; i < titulos.Length; i++)
            {
                if (titulos.Length > 1)
                {
                    do { Thread.Sleep(250); } while (!appExcel.Application.Ready);
                    colTitulos[i] = Celda(libroExcel, hojaExcel, fila, titulos[i], i != 0 ? titulos[0] : titulos[i + 1]);
                }
                else
                {
                    do { Thread.Sleep(250); } while (!appExcel.Application.Ready);
                    colTitulos[i] = Celda(libroExcel, hojaExcel, fila, titulos[i], titulos[i]);
                }
            }
            return colTitulos;
        }

        public static DataTable ColumnasDataTable(Excel.Application appExcel, Excel.Worksheet hojaExcel, string[] columnasExcel, int fila)
        {
            DataTable dataTable = new DataTable();

            for (int i = 0; i < columnasExcel.Length; i++)
            {
                dataTable.Columns.Add();
                string[] datos = ObtenerColumna(appExcel, hojaExcel, columnasExcel[i], fila);
                do { Thread.Sleep(250); } while (!appExcel.Application.Ready);

                if (i == 0)
                {
                    for (int j = 0; j < datos.Length; j++)
                        dataTable.Rows.Add();
                }

                if (dataTable.Rows.Count < datos.Length)
                {
                    for (int j = 0; j < datos.Length - dataTable.Rows.Count; j++)
                        dataTable.Rows.Add();
                }

                for (int j = 0; j < datos.Length; j++)
                    dataTable.Rows[j][i] = datos[j].Trim();
            }
            return dataTable;
        }

        public static DataTable PrimeraFilaTitulos(DataTable dataTable)
        {
            int contTitulos = 0;
            bool borrarRow = false;
            for (int i = 0; i < dataTable.Columns.Count; i++)
            {
                if (String.IsNullOrEmpty(dataTable.Rows[0][i].ToString()))
                {
                    dataTable.Columns[i].ColumnName = "Titulo " + contTitulos;
                    contTitulos++;
                }
                else
                {
                    if (dataTable.Columns[i].ColumnName != dataTable.Rows[0][i].ToString())
                    {
                        string nombreColumna = dataTable.Rows[0][i].ToString().Trim();

                        if (ChecarColumnas(dataTable, dataTable.Rows[0][i].ToString().Trim()))
                            nombreColumna += " 1";

                        dataTable.Columns[i].ColumnName = nombreColumna;
                        borrarRow = true;
                    }
                }
            }

            if (borrarRow)
                dataTable.Rows.RemoveAt(0);

            return dataTable;
        }

        private static bool ChecarColumnas(DataTable dataTable, string nombreColumna)
        {
            bool existeTitulo = false;
            foreach (DataColumn item in dataTable.Columns)
            {
                if (item.ColumnName.Equals(nombreColumna))
                {
                    existeTitulo = true;
                    break;
                }
            }
            return existeTitulo;
        }

        //public int BuscarHoja(Excel.Workbook libroExcel, string hojaBuscar)
        //{
        //    int totalHojas = objNuExcel.CantidadHojas(libroExcel);
        //    string nomHoja;
        //    int hojas = 0, posHoja = 0; string nombre = "";

        //    for (int indPest = 1; indPest <= totalHojas; indPest++)
        //    {
        //        nomHoja = objNuExcel.NombreHojaEn(indPest, libroExcel);
        //        if (nomHoja == hojaBuscar)
        //        {
        //            hojas++;
        //            posHoja = indPest;
        //        }
        //    }

        //    if (hojas == 1)
        //        return posHoja;
        //    else
        //    {
        //        string[] arrayHojas = new string[totalHojas + 1];
        //        for (int indPest = 1; indPest <= totalHojas; indPest++)
        //        {
        //            nombre = objNuExcel.NombreHojaEn(indPest, libroExcel);
        //            arrayHojas[indPest] = nombre;
        //        }
        //        nombre = ComboBox("Seleccione la hoja " + hojaBuscar, arrayHojas);
        //        posHoja = objNuExcel.HojaTrabajoSolicitada(libroExcel, nombre, 0);

        //        return posHoja;
        //    }
        //}

        public string CheckDocuments(string ruta, string rutaLog)
        {
            string rutaArchivo = string.Empty;
            try
            {
                var directory = new DirectoryInfo(ruta);
                var ultimoArchivo = (from last in directory.GetFiles().Where(file => !file.Name.Contains("~")) orderby last.LastWriteTime descending select last).First();

                rutaArchivo = ruta + ultimoArchivo.ToString();
            }
            catch (Exception)
            {
                //ReportaLog(rutaLog, "No se encontró el archivo necesario para realizar las operaciones Seleccione el archivo para continuar");
                //rutaArchivo = FileDialog("No se encontró el archivo necesario para realizar las operaciones \nSeleccione el archivo de Reporte mas actual", "Excel");

                if (string.IsNullOrEmpty(rutaArchivo) || !File.Exists(rutaArchivo))
                {
                    CheckDocuments(ruta, rutaLog);
                }
            }
            return rutaArchivo;
        }

        //public string ComboBox(string pregunta, string[] opciones)
        //{
        //    Opciones opcions = new Opciones();
        //    opciones[0] = pregunta;
        //    string respuesta = "";
        //    int res = 0;
        //    for (int a = 0; a < opciones.Length; a++)
        //        opcions.cmbxOpciones.Items.Add(opciones[a]);

        //    opcions.Title = "Nu4it - Cartera";
        //    opcions.lblAvisoContent.Text = pregunta;
        //    opcions.cmbxOpciones.SelectedIndex = 0;
        //    opcions.ShowDialog();
        //    res = opcions.cmbxOpciones.SelectedIndex;

        //    if (res == 0)
        //    {
        //        MessageShowOK(pregunta);
        //        respuesta = ComboBox(pregunta, opciones);
        //    }
        //    respuesta = opcions.cmbxOpciones.SelectedItem.ToString();
        //    return respuesta;
        //}

        //public void MessageShowOK(string aviso)
        //{
        //    MsjOK msgOK = new MsjOK
        //    {
        //        Title = "Nu4it - Cartera"
        //    };
        //    msgOK.lblAvisoContent.Text = aviso;
        //    msgOK.ShowDialog();
        //}

        //public void EnviarReporte(string rutaLog)
        //{

        //    string correo = "geyder.hernandez@bestcollect.com.mx";
        //    string usuario = Environment.UserName;
        //    string mac = ObtenerMACAddress();
        //    string headerReporte = "ID: " + mac + Environment.NewLine +
        //        "Usuario: " + usuario +
        //        Environment.NewLine + "******************************************************" + Environment.NewLine;
        //    string Reporte = File.ReadAllText(rutaLog);
        //    Reporte = Reporte.Replace(".,.", Environment.NewLine);
        //    Reporte = headerReporte + Reporte;
        //    try
        //    {
        //        DateTime Today = DateTime.Now;
        //        Microsoft.Office.Interop.Outlook.Application miOutlook = new Microsoft.Office.Interop.Outlook.Application();
        //        Microsoft.Office.Interop.Outlook.MailItem mail = (Microsoft.Office.Interop.Outlook.MailItem)miOutlook.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem);
        //        mail.To = correo;
        //        mail.Subject = "Reporte Cartera: " + Today.Day + "/" + Today.Month + "/" + Today.Year + " - " + Today.Hour + ":" + Today.Minute;
        //        mail.Body = Reporte;
        //        mail.Importance = Microsoft.Office.Interop.Outlook.OlImportance.olImportanceNormal;
        //        mail.Send();
        //    }
        //    catch (Exception)
        //    {
        //        MessageShowOK("No se pudo enviar el Reporte");
        //    }
        //}

        public string ObtenerMACAddress()
        {
            NetworkInterface[] nics = NetworkInterface.GetAllNetworkInterfaces();
            string macAddress = string.Empty;
            foreach (NetworkInterface adapter in nics)
            {
                if (macAddress == string.Empty)
                {
                    IPInterfaceProperties properties = adapter.GetIPProperties();
                    macAddress = adapter.GetPhysicalAddress().ToString();
                }
            }
            return macAddress;
        }
    }

}
