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
using LinqToExcel;
using Dielmex_Order_Manager.com.models;

namespace Dielmex_Order_Manager
{
    public partial class Hoja6
    {

        internal static List<Order> _orders = new List<Order>();

        private void Hoja6_Startup(object sender, System.EventArgs e)
        {

            
            
            


        }

        internal  void Hoja7_onLoaded()
        {
            // Sacar los datos de las tablas

            var table = new ExcelQueryFactory();

            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);


            List<Tuple<int, string>> listForEquipment = new List<Tuple<int,string>>();


            listForEquipment = (from row in book.Worksheet("DBOH")
                        let item = new Tuple<int, string>(row["No Orden"].Cast<int>(), row["No Equipo"].Cast<String>())
                       
                       select item).ToList();


            var res = (from row in book.Worksheet("DBOH")
                       let item = new Order
                       {
                           Folio = row["No Orden"].Cast<int>(),
                           
                           ServiceDate = row["Fecha de Servicio"].Cast<DateTime>(),
                           Expert = row["Tecnico"].Cast<string>(),
                           ReceivedBy = row["Recibio el servicio"].Cast<string>(),
                       }
                       select item).ToList();
            
            
            res = res.Select(c =>
                {
                c.Equipment = Hoja2._inventary.Where(
                    _inv =>
                        _inv.SerialNumber == listForEquipment.Where(
                        el =>
                            el.Item1 == c.Folio).FirstOrDefault<Tuple<int, string>>().Item2
                ).FirstOrDefault();


                c.OrderItems = Hoja7._itemsOrder.Where(item => item.OrderNumber == c.Folio).ToList();

                return c;
            
            }).ToList();
            
            
            if (res.Count == 1 && res.FirstOrDefault().Folio == 0)
            {
                _orders = new List<Order>();
            }else
            {
                _orders = res;
            }

            if (_orders.Exists(el => { return el.Folio == 0; }))
            {
                _orders.RemoveAt(0);
            }
            
        }

        private void Hoja6_Shutdown(object sender, System.EventArgs e)
        {
        }

        internal void save()
        {
            if (this.tbOrdenHeader.DataBodyRange != null)
            {
                this.tbOrdenHeader.DataBodyRange.Rows.Delete();
            }

            int count = 0;
            foreach (Order currentOrden in _orders)
            {
                
                int offset = (this.tbOrdenHeader.DataBodyRange == null) ? this.tbOrdenHeader.HeaderRowRange.Row + ++count : this.tbOrdenHeader.DataBodyRange.Rows.Row + count++;

                Globals.Hoja6.Range["A" + offset].Value = currentOrden.Folio;
                Globals.Hoja6.Range["B" + offset].Value = currentOrden.Equipment.SerialNumber;
                Globals.Hoja6.Range["C" + offset].Value = currentOrden.Equipment.WorkCentre;
                Globals.Hoja6.Range["D" + offset].Value = currentOrden.Equipment.Workplace;
                Globals.Hoja6.Range["E" + offset].Value = currentOrden.ServiceDate;
                Globals.Hoja6.Range["F" + offset].Value = currentOrden.Expert;
                Globals.Hoja6.Range["G" + offset].Value = currentOrden.ReceivedBy;

                this.tbOrdenHeader.ListRows.AddEx(System.Type.Missing, true);
            }

        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja6_Startup);
            this.Shutdown += new System.EventHandler(Hoja6_Shutdown);
        }

        #endregion

    }
}
