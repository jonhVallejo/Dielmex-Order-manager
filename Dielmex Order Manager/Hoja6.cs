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

        internal static List<Orden> _ordenes = new List<Orden>();

        private void Hoja6_Startup(object sender, System.EventArgs e)
        {

            
            
            


        }

        internal  void Hoja7_onLoaded()
        {
            // Sacar los datos de las tablas

            var table = new ExcelQueryFactory();

            var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);

            List<Tuple<int, string>> auxListForEquipo = new List<Tuple<int,string>>();

            auxListForEquipo = (from row in book.Worksheet("DBOH")
                        let item = new Tuple<int, string>(row["No Orden"].Cast<int>(), row["No Equipo"].Cast<String>())
                       
                       select item).ToList();

            var res = (from row in book.Worksheet("DBOH")
                       let item = new Orden
                       {
                           Folio = row["No Orden"].Cast<int>(),
                           
                           CentroTrabajo = row["Centro de trabajo"].Cast<string>(),
                           Delegacion = row["Delegacion"].Cast<string>(),
                           FechaServicio = row["Fecha de Servicio"].Cast<DateTime>(),
                           Tecnico = row["Tecnico"].Cast<string>(),
                           Recibio = row["Recibio el servicio"].Cast<string>(),
                       }
                       select item).ToList();
            res = res.Select(c =>
                {
                c.Equipo = Hoja2._inventary.Where(
                    _inv =>
                        _inv.NEconomico == auxListForEquipo.Where(
                        el =>
                            el.Item1 == c.Folio).FirstOrDefault<Tuple<int, string>>().Item2
                ).FirstOrDefault();


                c.Conceptos = Hoja7._conceptos.Where(concepto => concepto.Orden == c.Folio).ToList();

                return c;
            
            }).ToList();
            if (res.Count == 1 && res.FirstOrDefault().Folio == 0)
            {
                _ordenes = new List<Orden>();

            }else
            {
            _ordenes = res;
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
//            int rowsCount = this.tbOrdenHeader.DataBodyRange.Rows.de;

        //    for(int i = 0; i < rowsCount)

            int count = 0;
            Hoja7._conceptos.Clear();
            foreach (Orden currentOrden in _ordenes)
            {
                Hoja7._conceptos.AddRange(currentOrden.Conceptos);
                int offset = this.tbOrdenHeader.DataBodyRange.Rows.Row + count++;

                Globals.Hoja6.Range["A" + offset].Value = currentOrden.Folio;
                Globals.Hoja6.Range["B" + offset].Value = currentOrden.Equipo.NEconomico;
                Globals.Hoja6.Range["C" + offset].Value = currentOrden.CentroTrabajo;
                Globals.Hoja6.Range["D" + offset].Value = currentOrden.Delegacion;
                Globals.Hoja6.Range["E" + offset].Value = currentOrden.FechaServicio;
                Globals.Hoja6.Range["F" + offset].Value = currentOrden.Tecnico;
                Globals.Hoja6.Range["G" + offset].Value = currentOrden.Recibio;

                this.tbOrdenHeader.ListRows.AddEx(System.Type.Missing, true);
            }

            Globals.Hoja7.save();
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
