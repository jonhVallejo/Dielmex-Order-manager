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
    public partial class Hoja7
    {
        internal static List<OrderItem> _itemsOrder = new List<OrderItem>();


        internal event endLoaded onLoaded;

        private void Hoja7_Startup(object sender, System.EventArgs e)
        {
        }

        private void Hoja7_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void Hoja2_onLoaded()
        {
            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);
            // Orden	Clave	Precio Unitario	Cantidad	Subtotal

            var auxList = (from row in book.Worksheet("DBOB")
                           let item
                           = new Tuple<int, string, double>(row["Orden"].Cast<int>(),
                               row["Clave"].Cast<string>(), 
                               row["Cantidad"].Cast<double>())
                           select item).ToList();

            var res = auxList.Select(element =>
            {
                OrderItem temp;
                temp = new OrderItem();
                temp.OrderNumber = element.Item1;
                if (element.Item2 != null)
                {
                    temp.Equipment = Hoja1._services.Where(currentService => currentService.ServiceId == element.Item2).FirstOrDefault();
                    temp.Quantity = element.Item3;

                    temp.SubTotal = temp.Quantity * temp.Equipment.Cost;
                }
               
               


                return temp;
            });


            _itemsOrder = res.ToList();

            onLoaded();

        }

        internal void save()
        {
            if (this.tbOrdenBody.DataBodyRange != null)
            {
                this.tbOrdenBody.DataBodyRange.Rows.Delete();
            }
            int count = 0;

            foreach (OrderItem currentOrder in _itemsOrder)
            {
                if (currentOrder.Equipment != null)
                {
                    int offset = (this.tbOrdenBody.DataBodyRange == null) ? this.tbOrdenBody.HeaderRowRange.Row + ++count : this.tbOrdenBody.DataBodyRange.Rows.Row + count++;

                    this.Range["A" + offset].Value = currentOrder.OrderNumber;
                    this.Range["B" + offset].Value = currentOrder.Equipment.ServiceId;
                    this.Range["C" + offset].Value = currentOrder.Equipment.Cost;
                    this.Range["D" + offset].Value = currentOrder.Quantity;

                    this.tbOrdenBody.ListRows.AddEx(System.Type.Missing, true);
                }
            }
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja7_Startup);
            this.Shutdown += new System.EventHandler(Hoja7_Shutdown);
        }

        #endregion

    }
}
