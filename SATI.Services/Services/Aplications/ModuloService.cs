using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SATI.Common.Aplications;
using SATI.Common.Entities;
namespace SATI.Services.Services
{
    public class ModuloService
    {
        private ModuloCommon _moduloCommon = new ModuloCommon();

        public List<Modulos> ListadoModulosUsuarios(string idUsuario,string moduloActivo)
        {
            return _moduloCommon.ListadoModulosUsuario(idUsuario, moduloActivo);
        }
    }
}
