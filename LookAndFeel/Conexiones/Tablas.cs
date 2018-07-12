using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookAndFeel.Conexiones
{
    class Tablas
    {
        #region NombreTablas
        String StrCDF = "TblCdf",
               StrChedraui = "TblChedraui",
               StrComex = "TblComex",
               StrCostco = "TblCostco",
               StrFragua = "TblFragua",
               StrFresko = "TblFresko",
               StrHeb = "TblHeb",
               StrOxxo = "TblOxxo",
               StrSoriana = "TblSoriana",
               StrBepensa = "TblBepensa",
               StrWalmart = "TblWalmart";
        #endregion
        #region Tablas
        public DataTable tblCDF = new DataTable(),
                  tblChedragui = new DataTable(),
                  tblComex = new DataTable(),
                  tblCostco = new DataTable(),
                  tblFragua = new DataTable(),
                  tblFresko = new DataTable(),
                  tblHeb = new DataTable(),
                  tblOxxo = new DataTable(),
                  tblSoriana = new DataTable(),
                  tblBepensa = new DataTable(),
                  tblWalmart = new DataTable();
        #endregion

        public void GeneraTablas()
        {
            tblBepensa.Namespace = StrBepensa;
            tblBepensa.TableName = StrBepensa;
            tblBepensa.Columns.Add(CrearColumna("Usuario"));
            tblBepensa.Columns.Add(CrearColumna("Contrasenia"));
            tblBepensa.Rows.Add("", "");

            tblChedragui.Namespace = StrChedraui;
            tblChedragui.TableName = StrChedraui;
            tblChedragui.Columns.Add(CrearColumna("Usuario"));
            tblChedragui.Columns.Add(CrearColumna("Contrasenia"));
            tblChedragui.Rows.Add("", "");

            tblCDF.Namespace = StrCDF;
            tblCDF.TableName = StrCDF;
            tblCDF.Columns.Add(CrearColumna("Usuario"));
            tblCDF.Columns.Add(CrearColumna("Contrasenia"));
            tblCDF.Rows.Add("", "");

            tblCostco.Namespace = StrCostco;
            tblCostco.TableName = StrCostco;
            tblCostco.Columns.Add(CrearColumna("Usuario"));
            tblCostco.Columns.Add(CrearColumna("Contrasenia"));
            tblCostco.Rows.Add("", "");

            tblComex.Namespace = StrComex;
            tblComex.TableName = StrComex;
            tblComex.Columns.Add(CrearColumna("Usuario"));
            tblComex.Columns.Add(CrearColumna("Contrasenia"));
            tblComex.Rows.Add("", "");

            tblFragua.Namespace = StrFragua;
            tblFragua.TableName = StrFragua;
            tblFragua.Columns.Add(CrearColumna("Usuario"));
            tblFragua.Columns.Add(CrearColumna("Contrasenia"));
            tblFragua.Rows.Add("", "");

            tblFresko.Namespace = StrFresko;
            tblFresko.TableName = StrFresko;
            tblFresko.Columns.Add(CrearColumna("Usuario"));
            tblFresko.Columns.Add(CrearColumna("Contrasenia"));
            tblFresko.Rows.Add("", "");

            tblHeb.Namespace = StrHeb;
            tblHeb.TableName = StrHeb;
            tblHeb.Columns.Add(CrearColumna("Usuario"));
            tblHeb.Columns.Add(CrearColumna("Contrasenia"));
            tblHeb.Rows.Add("", "");

            tblOxxo.Namespace = StrOxxo;
            tblOxxo.TableName = StrOxxo;
            tblOxxo.Columns.Add(CrearColumna("Usuario"));
            tblOxxo.Columns.Add(CrearColumna("Contrasenia"));
            tblOxxo.Rows.Add("", "");

            tblSoriana.Namespace = StrSoriana;
            tblSoriana.TableName = StrSoriana;
            tblSoriana.Columns.Add(CrearColumna("Usuario"));
            tblSoriana.Columns.Add(CrearColumna("Contrasenia"));
            tblSoriana.Rows.Add("", "");

            tblWalmart.Namespace = StrWalmart;
            tblWalmart.TableName = StrWalmart;
            tblWalmart.Columns.Add(CrearColumna("Usuario"));
            tblWalmart.Columns.Add(CrearColumna("Contrasenia"));
            tblWalmart.Rows.Add("", "");
        }

        private DataColumn CrearColumna(String Nombre)
        {
            DataColumn dc = new DataColumn(Nombre, typeof(string));
            return dc;
        }
    }
}
