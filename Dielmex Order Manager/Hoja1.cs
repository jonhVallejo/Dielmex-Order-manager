using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Dielmex_Order_Manager.com.models;
using LinqToExcel;

namespace Dielmex_Order_Manager
{
    public partial class Hoja1
    {


        internal event endLoaded onLoaded;

        internal static List<Servicio> _services = new List<Servicio>();

        private void Hoja1_Startup(object sender, System.EventArgs e)
        {
            /*
            Excel.Range table = this.Tabla1.Range;

            bool header = true;
            
            int offset;

            offset = table.Column - 1;

            foreach (Excel.Range row in table.Rows)
            {
                if (!header)
                {
                    Servicio temp;
                    temp = new Servicio();

                    temp.Ref = row.Cells[1, offset + 1].value;
                    temp.UnidadMedida = row.Cells[1, offset + 2].value;
                    temp.Descripcion = row.Cells[1, offset + 3].value;
                    temp.Refacciones = (double)row.Cells[1, offset + 4].value;
                    temp.ManoObra = (double)row.Cells[1, offset + 5].value;
                    temp.Costo = (double)row.Cells[1, offset + 6].value;

                    _services.Add(temp);
                }
                else
                {
                    header = false;
                }
            }
             * */

            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);

            var res = (from row in book.Worksheet("Precios")
                       let item = new Servicio
                       {
                           Ref = row["REF"].Cast<string>(),
                           UnidadMedida = row["UNIDAD DE MEDIDA"].Cast<string>(),
                           Descripcion = row["DESCRIPCION"].Cast<string>(),
                           Refacciones = row["REFACCIONES"].Cast<double>(),
                           ManoObra = row["MANO DE OBRA"].Cast<double>(),
                           Costo = row["TOTAL"].Cast<double>()
                       }
                       select item).ToList();
            _services = res;

            onLoaded();

        }

        void Tabla1_SelectionChange(Excel.Range Target)
        {
            
        }

        private void Hoja1_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja1_Startup);
            this.Shutdown += new System.EventHandler(Hoja1_Shutdown);
        }

        #endregion


        public System.Reflection.Assembly assemblyInfo { get; set; }
    }
}
