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
        private string[] _DEFAULT_COLUMN_NAMES = { "REF", 
                                                         "UNIDAD DE MEDIDA",
                                                         "DESCRIPCION", 
                                                         "TOTAL"};

        internal event endLoaded onLoaded;

        internal static List<Service> _services = new List<Service>();

        internal static List<String> _dynamicColumNames;

        private void Hoja1_Startup(object sender, System.EventArgs e)
        {
            

            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);

            /*
             * Get the column names in the sheet 1. This is used
             * by know the dynamic columns.
             */
            var columNames = book.GetColumnNames(this.Name);

            /*
             * This Enumerable object represent the dynamic columns. The dynamic
             * columns aren't considers in the object model but can be saveds in
             * a list for search with some excel function.
             */
            _dynamicColumNames = columNames.Except(_DEFAULT_COLUMN_NAMES).ToList();

            

            /*
             * LINQ to sheet for retreive all values in a table
             */
            var res = (from row in book.Worksheet(this.Name)
                       let item = new Service
                       {
                           ServiceId = row["REF"].Cast<string>(),
                           UnitOfMeasurement = row["UNIDAD DE MEDIDA"].Cast<string>(),
                           Description = row["DESCRIPCION"].Cast<string>(),
                           Cost = row["TOTAL"].Cast<double>()
                       }
                       select item).ToList();
            _services = res;

            /*
             * Callback for syncronize the load of data.
             */
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
