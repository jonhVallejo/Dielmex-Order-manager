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
        private string _ref;

        internal string Ref
        {
            get { return _ref; }
            set { _ref = value; }
        }
        private string _descripcion;

        internal string Descripcion
        {
            get { return _descripcion; }
            set { _descripcion = value; }
        }
        private string _unidadMedida;

        internal string UnidadMedida
        {
            get { return _unidadMedida; }
            set { _unidadMedida = value; }
        }
        private double _refacciones;

        internal double Refacciones
        {
            get { return _refacciones; }
            set { _refacciones = value; }
        }
        private double _manoObra;

        internal double ManoObra
        {
            get { return _manoObra; }
            set { _manoObra = value; }
        }
        private double _costo;

        internal double Costo
        {
            get { return _costo; }
            set { _costo = value; }
        }

        public override string ToString()
        {
            return _ref;
        }


    }
}
