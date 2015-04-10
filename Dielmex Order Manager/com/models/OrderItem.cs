using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.com.models
{
    class OrderItem
    {
        /*
         *Orden	Clave	Precio Unitario	Cantidad	Subtotal

         */
        internal int OrderNumber { get; set; }
        internal Service Equipment { get; set; }
        internal double Quantity { get; set; }
        internal double SubTotal { get; set; }
    }
}
