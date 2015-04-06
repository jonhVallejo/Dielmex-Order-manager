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

        internal event endLoaded onLoaded;

        internal static List<Inventario> _inventary = new List<Inventario>();
        private void Hoja2_Startup(object sender, System.EventArgs e)
        {
           


        }

        internal void Hoja1_onLoaded()
        {
            // Inicializacion de la coleccion del inventario.
            _inventary = new List<Inventario>();

            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);

            var res = (from row in book.Worksheet("Inventario")
                       let item = new Inventario
                       {
                           Categoria = row["CATEGORIA"].Cast<string>(),
                           Tipo = row["TIPO"].Cast<string>(),
                           Marca = row["MARCA"].Cast<string>(),
                           Modelo = row["MODELO"].Cast<int>(),
                           Placa = row["PLACA"].Cast<string>(),
                           NEconomico = row["NECONOMICO"].Cast<string>(),
                           Red = row["RED"].Cast<string>(),
                           Cilindros = row["CILINDROS"] != null ? row["CILINDROS"].Cast<int>() : 0
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
