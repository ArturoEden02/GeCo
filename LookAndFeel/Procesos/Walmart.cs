namespace Pruebas_clase7.Clases
{
    using System;
    using Excel = Microsoft.Office.Interop.Excel;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Chrome;
    using OpenQA.Selenium.Support.UI;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Threading;
    using System.IO;
    using System.Windows;

    class Walmart
    {
        public void FuncionPrincipal(string usuario, string contraseña)
        {
            IWebDriver driver;
            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;
            driver = new ChromeDriver(driverService, new ChromeOptions());
            driver.Navigate().GoToUrl("https://retaillink.wal-mart.com/new_home/Site");
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            IWebElement element = wait.Until(ExpectedConditions.ElementExists(By.Id("txtUser")));
            driver.FindElement(By.Id("txtUser")).Clear();
            driver.FindElement(By.Id("txtUser")).SendKeys(usuario);
            driver.FindElement(By.Id("txtPass")).Clear();
            driver.FindElement(By.Id("txtPass")).SendKeys(contraseña);
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            element = wait.Until(ExpectedConditions.ElementExists(By.Id("Login")));
            element.Click();

            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            try
            {
                element = wait.Until(ExpectedConditions.ElementExists(By.LinkText("Estado de Cuenta Proveedores")));
                element.Click();
            }
            catch
            {
                element = wait.Until(ExpectedConditions.ElementExists(By.LinkText("Vendor Balance System")));
                element.Click();
            }

            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            element = wait.Until(ExpectedConditions.ElementExists(By.Id("ddlDepVend")));
            IWebElement combo = driver.FindElement(By.Id("ddlDepVend"));
            List<string> opciones = obtenerOpciones(combo);
            ArmarEstado(opciones, driver);
            driver.Close();
            driver.Quit();
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

        private List<string> obtenerOpciones(IWebElement elemento)
        {
            List<string> opcions = new List<string>();
            ReadOnlyCollection<IWebElement> opciones = elemento.FindElements(By.TagName("option"));
            for (int i = 0; i < opciones.Count; i++)
            {
                string valor = opciones[i].GetAttribute("value");
                opcions.Add(valor);
            }
            return opcions;
        }

        private void ArmarEstado(List<string> opcions, IWebDriver web)
        {
            Excel.Application miExcel = new Excel.Application();
            miExcel.DisplayAlerts = false;
            miExcel.Visible = true;
            Excel.Workbook libro = miExcel.Workbooks.Add();
            Excel.Worksheet hojaExcel;
            foreach (string item in opcions)
            {
                libro.Worksheets.Add();
                ((Excel.Worksheet)miExcel.Sheets[1]).Select();
                hojaExcel = (Excel.Worksheet)libro.ActiveSheet;
                hojaExcel.Columns.EntireColumn.NumberFormat = "@";
                hojaExcel.Cells["1", "A"] = "CIA";
                hojaExcel.Cells["1", "B"] = "Id Mov";
                hojaExcel.Cells["1", "C"] = "Tienda";
                hojaExcel.Cells["1", "D"] = "Depto";
                hojaExcel.Cells["1", "E"] = "Folio";
                hojaExcel.Cells["1", "F"] = "Factura";
                hojaExcel.Cells["1", "G"] = "F. Recibo";
                hojaExcel.Cells["1", "H"] = "F. Vencimiento";
                hojaExcel.Cells["1", "I"] = "Importe";
                hojaExcel.Cells["1", "J"] = "Estatus";
                hojaExcel.Cells["1", "K"] = "Orden Compra";
                new SelectElement(web.FindElement(By.Id("ddlDepVend"))).SelectByValue(item);
                WebDriverWait wait = new WebDriverWait(web, TimeSpan.FromSeconds(10));
                IWebElement element = wait.Until(ExpectedConditions.ElementExists(By.Id("btnSearch")));
                element.Click();
                wait = new WebDriverWait(web, TimeSpan.FromSeconds(10));
                element = wait.Until(ExpectedConditions.ElementExists(By.Id("tb_grid")));
                string texto = "";
                char[] delimiterChars = { '\r', '\n' };
                string[] lineas = new string[100];
                texto = element.GetAttribute("innerHTML");
                lineas = texto.Split(delimiterChars);
                texto = texto.Replace("\r", "♥");
                texto = texto.Replace("\t", "♦");
                texto = texto.Replace("</tr><tr class=\"celdaCont\" align=\"right\">", "");
                texto = texto.Replace("</tr><tr align=\"right\">", "");
                texto = texto.Replace("<td>", "♠");
                texto = texto.Replace("</td>", "•");
                texto = texto.Replace("</tr><tr align=\"right\" class=\"celdatot\">", "");
                texto = texto.Replace("<td colspan=\"8\">", "");
                texto = texto.Replace("</tr>", "");
                texto = texto.Replace("</tbody>", "");
                texto = texto.Replace("<tr class=\"celdatot\" align=\"right\">", "");
                texto = texto.Replace("<tbody><tr id=\"row\" align=\"center\" valign=\"middle\" style=\"color:White;background-color:Navy;font-weight:bold;\">", "");
                texto = texto.Replace("<td id=\"CIA\" key=\"lb_cia_tc\" name=\"Tablecell01\">", "");
                texto = texto.Replace("<td id=\"MovId\" key=\"lb_mov_tc\" name=\"Tablecell02\">", "");
                texto = texto.Replace("<td id=\"Shop\" key=\"lb_tienda_tc\" name=\"Tablecell03\">", "");
                texto = texto.Replace("<td id=\"Dept\" key=\"lb_depto_tc\" name=\"Tablecell04\">", "");
                texto = texto.Replace("<td id=\"Folio\" key=\"lb_folio_tc\" name=\"Tablecell05\">", "");
                texto = texto.Replace("<td id=\"Bill\" key=\"lb_factura_tc\" name=\"Tablecell06\">", "");
                texto = texto.Replace("<td id=\"ReceiptDate\" key=\"lb_fecharec_tc\" name=\"Tablecell07\">", "");
                texto = texto.Replace("<td id=\"ExpirationDate\" key=\"lb_fechaven_tc\" name=\"Tablecell08\">", "");
                texto = texto.Replace("<td id=\"Amount\" key=\"lb_importe_tc\" name=\"Tablecell09\">", "");
                texto = texto.Replace("<td id=\"Status\" key=\"lb_estatus_tc\" name=\"Tablecell12\">", "");
                texto = texto.Replace("<td id=\"PurchaseOrder\" key=\"lb_ordencpa_tc\" name=\"Tablecell13\">", "");
                texto = texto.Replace("♦CIA•Id Mov•Tienda•Depto•Folio•Factura•Fecha Recibo•Fecha Vencimiento•Importe•Estatus•Orden Compra•", "");
                texto = texto.Replace("♥\n♦♦♥\n", "");
                texto = texto.Replace("•♦♦♦♠", Environment.NewLine);
                texto = texto.Replace("•♠", "\t");
                texto = texto.Replace("♦♦♦♦♦♠", "");
                texto = texto.Replace("•♦♦♦", Environment.NewLine);
                texto = texto.Replace("••♦", "");
                Clipboard.Clear();
                Clipboard.SetText(texto);
                Thread.Sleep(1000);
                hojaExcel.Cells[2, 1].PasteSpecial();
                hojaExcel.Cells[1].EntireRow.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                hojaExcel.Cells[1].EntireRow.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkBlue);
                hojaExcel.Cells[1].EntireRow.Font.Bold = true;
                hojaExcel.Name = item;
                hojaExcel.Columns.EntireColumn.AutoFit();
            }

            for (int i = 1; i <= libro.Worksheets.Count; i++)
            {
                if (libro.Worksheets[i].name.Contains("Hoja"))
                {
                    libro.Worksheets[i].Delete();
                    i = 0;
                }
            }

            String rutaEscritorio = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            if (!Directory.Exists(rutaEscritorio + @"\Archivos Generados\")) Directory.CreateDirectory(rutaEscritorio + @"\Archivos Generados\");
            string nombre = "Walmart Estado de Cuenta " + nombreAleatorio() + ".xlsx";
            libro.SaveAs(rutaEscritorio + @"\Archivos Generados\" + nombre);
        }
    }
}