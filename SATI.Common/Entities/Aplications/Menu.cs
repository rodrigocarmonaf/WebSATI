using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SATI.Common.Entities
{
    public class Menu
    {
        public string Nombre { get; set; }
        public string Enlace { get; set; }

        public List<ItemsMenu> items { get; set; }
    }

    public class ItemsMenu
    {
        public string Nombre { get; set; }
        public string Enlace { get; set; }
    }
}
