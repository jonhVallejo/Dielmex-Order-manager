using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.com.models
{
    class Orden
    {
       

        internal int Folio { get; set; }
        internal Inventario Equipo { get; set; }
        internal string CentroTrabajo { get; set; }
        internal string Delegacion { get; set; }
        internal DateTime FechaServicio { get; set; }
        internal string Tecnico { get; set; }
        internal string Recibio { get; set; }

        internal List<ConceptoOrden> Conceptos { get; set; }


        public override string ToString()
        {
            return String.Format("{0}", Folio);
        }
    }
}
