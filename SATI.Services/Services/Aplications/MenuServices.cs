using System;
using System.Collections.Generic;
using SATI.Common.Aplications;
using SATI.Common.Entities;
namespace SATI.Services.Services
{
    public class MenuServices
    {
        private MenuCommon _menuCommon = new MenuCommon();

        public List<Menu> ListadoMenuAplicacionUsuarios(string idUsuario, string modulo)
        {
            return _menuCommon.ListadoMenuAplicacionsUsuario(idUsuario, modulo);
        }

        public List<Menu> ListadoItemsMenuAplicacionUsuarios(string idUsuario, string modulo)
        {
            return _menuCommon.ListadoMenuAplicacionsUsuario(idUsuario, modulo);
        }
    }
}
