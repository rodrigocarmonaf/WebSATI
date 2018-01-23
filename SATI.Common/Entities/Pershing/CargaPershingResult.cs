using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SATI.Common.Entities.Pershing
{
    public class CargaPershingResult
    {
        public string Instrumento { get; set; }
        public string Mensaje { get; set; }
        public int RentaVariablePendiente { get; set; }
        public int FFMMPendiente { get; set; }
        public int RentaFijaPendiente { get; set; }

    }
}
