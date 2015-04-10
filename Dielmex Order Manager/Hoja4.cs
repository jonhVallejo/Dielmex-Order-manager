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

namespace Dielmex_Order_Manager
{
    public partial class Hoja4
    {
        private void Hoja4_Startup(object sender, System.EventArgs e)
        {
            this.BeforeRightClick += Hoja4_BeforeRightClick;
            
        }

        void Hoja4_BeforeRightClick(Excel.Range Target, ref bool Cancel)
        {
            // Concatenar todos los datos de la orden seleccionada
            double selectedOrder = (double)this.Range["A7"].Value;

            var orderReference = Hoja6._orders.Find(order =>
            {
             
                return order.Folio == selectedOrder;
            });

            

            StringBuilder concatenateValues = new StringBuilder();

                
            if (orderReference != null)
            {
                orderReference.OrderItems.ForEach(element =>
                {
                    concatenateValues.Append(element.Equipment.ServiceId);
                    concatenateValues.Append("    ");
                    concatenateValues.Append(element.Equipment.Description);
                    concatenateValues.Append("    ");
                    concatenateValues.Append(element.Equipment.UnitOfMeasurement);
                    concatenateValues.Append("    ");
                    concatenateValues.Append(element.Quantity);
                    concatenateValues.Append("    ");
                    concatenateValues.Append(element.Equipment.Cost);
                    concatenateValues.Append("    ");
                    concatenateValues.Append(element.Quantity * element.Equipment.Cost);
                    concatenateValues.Append("\n");
                });
                Target.WrapText = true;
                Target.Value = concatenateValues.ToString();
            }
            
            

           

        }

        void Hoja4_ActivateEvent()
        {
        }

        private void Hoja4_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(Hoja4_Startup);
            this.Shutdown += new System.EventHandler(Hoja4_Shutdown);
        }

        #endregion

    }
}
