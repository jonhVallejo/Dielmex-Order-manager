using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.com.models
{
    class Servicio
    {
        // Ref   unidad medida    descripcion    refacciones    mano obra    total
        internal string Ref
        {
            get;
            set;
        }

        internal string Descripcion
        {
            get;
            set;
        }

        internal string UnidadMedida
        {
            get;
            set;
        }

        internal double Refacciones
        {
            get;
            set;
        }

        internal double ManoObra
        {
            get;
            set;
        }

        internal double Costo
        {
            get;
            set;
        }
        public override string ToString()
        {
            return Ref;
        }


    }
}
