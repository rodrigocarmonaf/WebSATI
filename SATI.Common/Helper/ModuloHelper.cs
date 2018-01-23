using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SATI.Common.Helper
{
    public static class ModuloHelper
    {
        public static string[] ObtenerActionAndController(string enlace)
        {
            return enlace.Split('/');
        }
    }
}
