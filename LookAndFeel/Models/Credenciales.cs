using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LookAndFeel.Models
{
    class Credenciales
    {
        string strUsuario, strContrasenia;
        public string Usuario { get { return strUsuario; } set { strUsuario = value; } }

        public string Contrasenia { get { return strContrasenia; } set { strContrasenia = value; } }
    }
}
