using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SATI.Common.Entities.Pershing;

namespace SATI.Common.Entities.Excel
{
    public class ExcelPershing
    {
        public string Id { get; set; }
        public string fecha_proceso { get; set; }
        public string num_cuenta { get; set; }
        public string nombre_corto { get; set; }
        public string fecha_trans { get; set; }
        public string fecha_strlmnt { get; set; }
        public string tipo_ope { get; set; }
        public string cantidad { get; set; }
        public string simbolo { get; set; }
        public string cusip { get; set; }
        public string desc_sec { get; set; }
        public string precio { get; set; }
        public string cod_trans { get; set; }
        public string interes { get; set; }
        public string monto_neto_base { get; set; }
        public string monto_neto_trans { get; set; }
        public string categoria { get; set; }
        public string comision { get; set; }
        public string impuestos { get; set; }
        public string currency { get; set; }
        public string monto_principal { get; set; }
        public string isin { get; set; }
        public string ip_number { get; set; }
        public string ip_name { get; set; }
        public string acct_mnemo { get; set; }
        public string mstr_mnemo { get; set; }
        public string corr_ref { get; set; }
        public string solicited { get; set; }
        public string org_input_src { get; set; }
        public string clase_activo { get; set; }
        public string sec_payment { get; set; }
        public string lot_selection { get; set; }
        public string tipo_cuenta { get; set; }
        public string exec_broker { get; set; }
        public string nombre_completo { get; set; }
        public string monto_neto { get; set; }
        public string act_descrip { get; set; }
        public string folio { get; set; }
        public string bruto { get; set; }
        public string calc_precio { get; set; }
        
        public ClientePershing Cliente { get; set; }    
    }
}
