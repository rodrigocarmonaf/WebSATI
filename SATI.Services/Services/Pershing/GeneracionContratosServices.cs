using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using SATI.Common.Common;
using SATI.Common.Entities;
using SATI.Common.Entities.Excel;
using SATI.Common.Entities.Pershing;

namespace SATI.Services.Services.Pershing
{
    public class GeneracionContratosServices
    {
        private PershingCommon _pershingCommon = new PershingCommon();

        public List<DateTime> ListadoFechaContratos()
        {
            return _pershingCommon.ListadoFechasContratosAGenerar();
        }

        public List<Contrato> ListadoContratos(DateTime fecha)
        {
            return _pershingCommon.ListadoOperacionesPendientes(fecha);
        }
    }
}
