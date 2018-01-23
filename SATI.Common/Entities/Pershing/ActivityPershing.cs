using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SATI.Common.Entities.Excel;
namespace SATI.Common.Entities
{
    public class ActivityPershing
    {
        public string FechaCarga { get; set; }
        public string NumeroFolio { get; set; }
        public int RegistroAnulados { get; set; }
        public int TotalRegistros { get; set; }
        public int TotalPendienteFM { get; set; }
        public bool Duplicado { get; set; }
        public double valorDolar { get; set; }
        public string Instrumento { get; set; }
        public string TotalComCLP { get; set; }
        public double TotalComUS { get; set; }
        public List<ExcelPershing> listadoExcelPershing { get; set; }
        public List<ExcelPershing> listadoExcelOmitidos { get; set; }
        public PershingError Error { get; set; }
    }  
}
