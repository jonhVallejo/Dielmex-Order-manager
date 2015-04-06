using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.com.models
{
    class ConceptoOrden
    {
        /*
         *Orden	Clave	Precio Unitario	Cantidad	Subtotal

         */
        internal int Orden { get; set; }
        internal Servicio Equipo { get; set; }
        internal double Cantidad { get; set; }
        internal double SubTotal { get; set; }
    }
}
