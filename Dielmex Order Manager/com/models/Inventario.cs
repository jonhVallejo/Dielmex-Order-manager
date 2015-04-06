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
        private string categoria;
        private string tipo;
        private string marca;
        private int modelo;
        private string placa;
        private string nEconomico;
        private string red;
        private int cilindros;

        internal string Categoria
        {
            get
            {
                return this.categoria;
            }
            set
            {
                this.categoria = value;
            }
        }
        internal string Tipo
        {
            get
            {
                return this.tipo;
            }
            set
            {
                this.tipo = value;
            }
        }
        internal string Marca
        {
            get
            {
                return this.marca;
            }
            set
            {
                this.marca = value;
            }
        }
        internal int Modelo
        {
            get
            {
                return this.modelo;
            }
            set
            {
                this.modelo = value;
            }
        }
        internal string Placa
        {
            get
            {
                return this.placa;
            }
            set
            {
                this.placa = value;
            }
        }
        internal string NEconomico
        {
            get
            {
                return this.nEconomico;
            }
            set
            {
                this.nEconomico = value;
            }
        }
        internal string Red
        {
            get
            {
                return this.red;
            }
            set
            {
                this.red = value;
            }
        }
        internal int Cilindros
        {
            get
            {
                return this.cilindros;
            }

            set
            {
                this.cilindros = value;
            }
        }
        
        public override string ToString()
        {
            return NEconomico;
        }
        

       
    }
}
