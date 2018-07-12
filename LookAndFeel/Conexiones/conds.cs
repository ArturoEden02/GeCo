using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Threading.Tasks;

namespace LookAndFeel.Conexiones
{
    class conds
    {
        public DataSet ds = new DataSet();
        Tablas dt = new Tablas();
        String RutaBase = Directory.GetCurrentDirectory() + "\\nu4it.db";

        public conds()
        {

            GetInfo();
            if (File.Exists(RutaBase))
            {
                if (!ds.Tables.Contains(dt.tblBepensa.TableName)) { ds.Tables.Add(dt.tblBepensa); }
                if (!ds.Tables.Contains(dt.tblCDF.TableName)) { ds.Tables.Add(dt.tblCDF); }
                if (!ds.Tables.Contains(dt.tblChedragui.TableName)) { ds.Tables.Add(dt.tblChedragui); }
                if (!ds.Tables.Contains(dt.tblComex.TableName)) { ds.Tables.Add(dt.tblComex); }
                if (!ds.Tables.Contains(dt.tblCostco.TableName)) { ds.Tables.Add(dt.tblCostco); }
                if (!ds.Tables.Contains(dt.tblFragua.TableName)) { ds.Tables.Add(dt.tblFragua); }
                if (!ds.Tables.Contains(dt.tblFresko.TableName)) { ds.Tables.Add(dt.tblFresko); }
                if (!ds.Tables.Contains(dt.tblHeb.TableName)) { ds.Tables.Add(dt.tblHeb); }
                if (!ds.Tables.Contains(dt.tblOxxo.TableName)) { ds.Tables.Add(dt.tblOxxo); }
                if (!ds.Tables.Contains(dt.tblSoriana.TableName)) { ds.Tables.Add(dt.tblSoriana); }
                if (!ds.Tables.Contains(dt.tblWalmart.TableName)) { ds.Tables.Add(dt.tblWalmart); }
            }
            else
            {
                dt.GeneraTablas();
                ds.Tables.Add(dt.tblBepensa);
                ds.Tables.Add(dt.tblCDF);
                ds.Tables.Add(dt.tblChedragui);
                ds.Tables.Add(dt.tblComex);
                ds.Tables.Add(dt.tblCostco);
                ds.Tables.Add(dt.tblFragua);
                ds.Tables.Add(dt.tblFresko);
                ds.Tables.Add(dt.tblHeb);
                ds.Tables.Add(dt.tblOxxo);
                ds.Tables.Add(dt.tblSoriana);
                ds.Tables.Add(dt.tblWalmart);
            }
            SetInfo();
        }

        public void SetInfo()
        {

            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            System.Runtime.Serialization.IFormatter formatter = new System.Runtime.Serialization.Formatters.Binary.BinaryFormatter();
            formatter.Serialize(stream, ds);

            byte[] bytes = stream.GetBuffer();
            String Valor = Convert.ToBase64String(bytes);
            var text = File.CreateText(RutaBase);
            text.Write(Valor);
            text.Close();
        }

        public void GetInfo()
        {
            if (File.Exists(RutaBase))
            {
                byte[] ContenidoByte = new byte[0];
                try
                {
                    String Texto = File.ReadAllText(RutaBase);
                    ContenidoByte = Convert.FromBase64String(Texto);
                    using (MemoryStream stream = new MemoryStream(ContenidoByte))
                    { BinaryFormatter brmater = new BinaryFormatter(); ds = (DataSet)brmater.Deserialize(stream); }
                }
                catch (System.Exception ex)
                { }
            }
        }
    }
}
