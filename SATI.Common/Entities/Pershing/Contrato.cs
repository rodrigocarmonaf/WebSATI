using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SATI.Common.Entities.Pershing
{
    public class Contrato
    {
        public int Id { get; set; }
        public long Folio { get; set; }
        public string NumeroCuenta { get; set; }
        public DateTime FechaProceso { get; set; }       
        public string Simbol { get; set; }
        public string Currency { get; set; }
        public string TipoInstrumento { get; set; }
        public string Isin { get; set; }
        public decimal Nominales { get; set; }
        public decimal Precio { get; set; }
        public string TipoOperacion { get; set; }
        public decimal Comision { get; set; }
        public string RutCliente { get; set; }
        public string NombreCliente { get; set; }
        
    }

    public class ContratosCliente
    {
        public string RutCliente { get; set; }
        public string NombreCliente { get; set; }
        public DateTime FechaProceso { get; set; }
        public string NumCuenta { get; set; }
        public List<Contrato> ListadoContratos { get; set; }
    }
}
