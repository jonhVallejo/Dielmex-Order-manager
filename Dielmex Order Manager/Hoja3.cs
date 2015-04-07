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
    public partial class Hoja3
    {

        private Dictionary<int, ComboBox> comboBoxes;
        private Dictionary<int, Tuple<Button, int>> buttons;

        enum ActionForButtonNew
        {
            NEW,
            CANCEL
        };

        private ActionForButtonNew _actionButton = ActionForButtonNew.NEW;

        private int _offsetForComboxInTable;
        private int _firstIndexForTable;

        private void Hoja3_Startup(object sender, System.EventArgs e)
        {
            cbEquipo.DisplayMember = "Dielmex_Order_Manager.com.models.Inventario.NEconomico";
            cbEquipo.ValueMember = "NEconomico";
            cbEquipo.DataSource = Hoja2._inventary;
            cbEquipo.Visible = false;

            

            cbOrdenNumber.DataSource = Hoja6._ordenes;
            cbOrdenNumber.DisplayMember = "Folio";

            cbOrdenNumber.SelectedValueChanged += cbOrdenNumber_SelectedValueChanged;

            cbEquipo.SelectedValueChanged += cbEquipo_SelectedValueChanged;

            comboBoxes = new Dictionary<int, ComboBox>();
            buttons = new Dictionary<int, Tuple<Button, int>>();




            _offsetForComboxInTable += 2;
            _firstIndexForTable = 19;
        }

        void cbOrdenNumber_SelectedValueChanged(object sender, EventArgs e)
        {
            btGuardar.Enabled = true;
            btNuevo.Enabled = false; 
            cbEquipo.Visible = true;
            cbEquipo.Enabled = true;

            _actionButton = ActionForButtonNew.CANCEL;
            btNuevo.Text = "Cancelar";
            btNuevo.Enabled = true;

            Orden tempOrden;
            tempOrden = (Orden)cbOrdenNumber.SelectedItem;

            cbEquipo.SelectedIndex = Hoja2._inventary.FindIndex(el => { return el.NEconomico == tempOrden.Equipo.NEconomico; });

            cbOrdenNumber.Enabled = false;
            cbOrdenNumber.Visible = false;

            Globals.Hoja3.Range["b12"].Value = tempOrden.CentroTrabajo;
            Globals.Hoja3.Range["b13"].Value = tempOrden.Delegacion;
            Globals.Hoja3.Range["f5"].Value = tempOrden.FechaServicio;
            Globals.Hoja3.Range["f7"].Value = tempOrden.Folio;
            Globals.Hoja3.Range["f9"].Value = tempOrden.Tecnico;
            Globals.Hoja3.Range["f11"].Value = tempOrden.Recibio;
        }

        void cbEquipo_SelectedValueChanged(object sender, EventArgs e)
        {
            Globals.Hoja3.Cells[9, 2].value = ((Inventario)cbEquipo.SelectedItem).Marca;
            Globals.Hoja3.Cells[10, 2].value = ((Inventario)cbEquipo.SelectedItem).Modelo;
            Globals.Hoja3.Cells[11, 2].value = ((Inventario)cbEquipo.SelectedItem).Tipo;
            
        }

        private void Hoja3_Shutdown(object sender, System.EventArgs e)
        {
            Console.Write("");
        }

        #region Código generado por el Diseñador de VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.btAdd.Click += new System.EventHandler(this.btAdd_Click);
            this.btNuevo.Click += new System.EventHandler(this.button1_Click);
            this.btGuardar.Click += new System.EventHandler(this.button2_Click);
            this.Startup += new System.EventHandler(this.Hoja3_Startup);
            this.Shutdown += new System.EventHandler(this.Hoja3_Shutdown);

        }

        #endregion
        int c = 0;
        private void btAdd_Click(object sender, EventArgs e)
        {
            /*
            Microsoft.Office.Tools.Excel.ControlSite dynamicControl;

            myButton.Text = "Hola";


            this.Controls.AddControl(myButton, Globals.Hoja3.Range["B1", "B1"], "Halo");
            
             this.Controls.AddComboBox(Globals.Hoja3.Range["A1", "A1"], "cbX");
             * */
            ComboBox temp;
            temp = new ComboBox();


            temp.DataSource =  new BindingSource(Hoja1._services, null);
            temp.DisplayMember = "_ref";

            temp.SelectedValueChanged += temp_SelectedValueChanged;

            this.comboBoxes.Add(temp.GetHashCode(), temp);

            Microsoft.Office.Interop.Excel.Range row = (Microsoft.Office.Interop.Excel.Range)Globals.Hoja3.Rows[_firstIndexForTable];
            row.Insert();

            this.Controls.AddControl(temp,Globals.Hoja3.Range["A" + _firstIndexForTable, "A" + _firstIndexForTable], "" + temp.GetHashCode());

            Button tempButton = new Button();
            tempButton.Text = "Eliminar";
            tempButton.Click +=tempButton_Click;
            this.buttons.Add(tempButton.GetHashCode(), new Tuple<Button, int>(tempButton, temp.GetHashCode()));
            this.Controls.AddControl(tempButton, Globals.Hoja3.Range["G" + _firstIndexForTable, "G" + _firstIndexForTable], "" + tempButton.GetHashCode());



        }

        void tempButton_Click(object sender, EventArgs e)
        {
            ComboBox t = this.comboBoxes[this.buttons[((Button)sender).GetHashCode()].Item2];
            int offset = _firstIndexForTable + (this.Controls.Count - _offsetForComboxInTable - buttons.Count - (this.Controls.IndexOf(t) / 2 ));

            this.comboBoxes.Remove(this.buttons[((Button)sender).GetHashCode()].Item2);
            this.buttons.Remove(this.buttons[((Button)sender).GetHashCode()].Item2);

            this.Controls.Remove(t);
            this.Controls.Remove((Button)sender);

            Excel.Range range = Globals.Hoja3.get_Range(String.Format("A{0}:A{0}", offset), System.Reflection.Missing.Value);
            
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

        }

        void temp_SelectedValueChanged(object sender, EventArgs e)
        {

            int offset = _firstIndexForTable + (this.Controls.Count - _offsetForComboxInTable - this.Controls.IndexOf(sender) + 1);
         //   Globals.Hoja3.Range["B" + offset, "B" + offset].Value = ((Inventario)((ComboBox)sender).SelectedItem).NEconomico;
            Globals.Hoja3.Range["B" + offset, "B" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).Descripcion;
            Globals.Hoja3.Range["C" + offset, "C" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).UnidadMedida;
            Globals.Hoja3.Range["D" + offset, "D" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).Costo;



            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            switch (_actionButton)
            {
                case ActionForButtonNew.NEW:
                    //Habilitar campos y sacar el ultimo folio
                    btNuevo.Text = "Cancelar";
                    _actionButton = ActionForButtonNew.CANCEL;
                    btGuardar.Enabled = true;
                    btAdd.Enabled = true;
                    cbEquipo.Visible = true;
                    cbEquipo.Enabled = true;
                    break;
                case ActionForButtonNew.CANCEL:
                    _actionButton = ActionForButtonNew.NEW;
                    btNuevo.Text = "Nuevo";
                    btGuardar.Enabled = false;
                    btAdd.Enabled = false;
                    cbEquipo.Visible = false;
                    cbEquipo.Enabled = false;
                    cbOrdenNumber.Enabled = true;
                    cbOrdenNumber.Visible = true;
                    clear();
                    break;
            }
            
        }

        private void clear()
        {
            Globals.Hoja3.Cells[9, 2].value = "";
            Globals.Hoja3.Cells[10, 2].value = "";
            Globals.Hoja3.Cells[11, 2].value = "";

            Globals.Hoja3.Range["b12"].Value = "";
            Globals.Hoja3.Range["b13"].Value = "";
            Globals.Hoja3.Range["f5"].Value = "";
            Globals.Hoja3.Range["f7"].Value = "";
            Globals.Hoja3.Range["f9"].Value = "";
            Globals.Hoja3.Range["f11"].Value = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Guardar la orden.
        }

       

        

    }
}
