using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SATI.Common.Entities
{
    public class Modulos
    {
        public int Id { get; set; }
        public string Titulo { get; set; }
        public string Controller { get; set; }
        public string Action { get; set; }
        public string Activo { get; set; }
    }
}
