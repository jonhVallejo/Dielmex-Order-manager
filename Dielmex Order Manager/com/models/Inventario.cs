using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dielmex_Order_Manager.com.models
{
    internal sealed class Inventario
    {
        //No.	CATEGORIA	TIPO	MARCA	MODELO	PLACA	No. ECONÓMICO	RED	CUNDROS


        internal string Categoria
        {
            get;
            set;
        }
        internal string Tipo
        {
            get;
            set;
        }
        internal string Marca
        {
            get;
            set;
        }
        internal int Modelo
        {
            get;
            set;
        }
        internal string Placa
        {
            get;
            set;
        }
        internal string NEconomico
        {
            get;
            set;
        }
        internal string Red
        {
            get;
            set;
        }
        internal int Cilindros
        {
            get;
            set;
        }
        
        public override string ToString()
        {
            return NEconomico;
        }
        

       
    }
}
