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
    delegate  void endLoaded();
    public partial class Hoja2
    {
        #region < LOGGER >

      //  private static NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();

        #endregion

        private string[] _DEFAULT_COLUMN_NAMES = { "CATEGORIA", "TIPO", "MARCA", "MODELO", "NECONOMICO", "CENTRO_TRABAJO", "DELEGACION" };

        internal event endLoaded onLoaded;

        internal static List<Inventary> _inventary = new List<Inventary>();

        internal static List<String> _dynamicColumNames;

        private void Hoja2_Startup(object sender, System.EventArgs e)
        {
        }

        internal void Hoja1_onLoaded()
        {
            _inventary = new List<Inventary>();

            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);

            var columNames = book.GetColumnNames(this.Name);

            _dynamicColumNames = columNames.Except(_DEFAULT_COLUMN_NAMES).ToList();


            var res = (from row in book.Worksheet(this.Name)
                       let item = new Inventary
                       {
                           Category = row["CATEGORIA"].Cast<string>(),
                           Type = row["TIPO"].Cast<string>(),
                           Brand = row["MARCA"].Cast<string>(),
                           Model = row["MODELO"].Cast<string>(),
                           SerialNumber = row["NECONOMICO"].Cast<string>(),
                           WorkCentre = row["CENTRO_TRABAJO"].Cast<string>(),
                           Workplace = row["DELEGACION"].Cast<string>()
                       }
                       select item).ToList();
            _inventary = res;

            onLoaded();
        }


        private void Hoja2_Shutdown(object sender, System.EventArgs e)
        {
            
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja2_Startup);
            this.Shutdown += new System.EventHandler(Hoja2_Shutdown);
        }

        #endregion

    }
}
