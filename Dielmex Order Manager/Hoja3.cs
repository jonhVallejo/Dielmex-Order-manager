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
    public partial class Hoja3
    {

        private List<ComboBox> comboBoxes;
        private List<Button> buttons;

        private Orden tempOrden;

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

            comboBoxes = new List<ComboBox>();
            buttons = new List<Button>();

            _offsetForComboxInTable += this.Controls.Count;
            _firstIndexForTable = 19;
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

        void cbOrdenNumber_SelectedValueChanged(object sender, EventArgs e)
        {
            btGuardar.Enabled = true;
            btNuevo.Enabled = false; 
            cbEquipo.Visible = true;
            cbEquipo.Enabled = true;

            _actionButton = ActionForButtonNew.CANCEL;
            btNuevo.Text = "Cancelar";
            btNuevo.Enabled = true;

            
            tempOrden = (Orden)cbOrdenNumber.SelectedItem;

            renderOrden(tempOrden);
        }

        void cbEquipo_SelectedValueChanged(object sender, EventArgs e)
        {
            Globals.Hoja3.Cells[9, 2].value = ((Inventario)cbEquipo.SelectedItem).Marca;
            Globals.Hoja3.Cells[10, 2].value = ((Inventario)cbEquipo.SelectedItem).Modelo;
            Globals.Hoja3.Cells[11, 2].value = ((Inventario)cbEquipo.SelectedItem).Tipo;
            
        }

        private void renderOrden(Orden orden)
        {
            if (orden.Equipo != null)
            {
                cbEquipo.SelectedIndex = Hoja2._inventary.FindIndex(el => { return el.NEconomico == orden.Equipo.NEconomico; });
            }
            cbOrdenNumber.Enabled = false;
            cbOrdenNumber.Visible = false;

            Globals.Hoja3.Range["b12"].Value = orden.CentroTrabajo;
            Globals.Hoja3.Range["b13"].Value = orden.Delegacion;
            Globals.Hoja3.Range["f5"].Value = orden.FechaServicio;
            Globals.Hoja3.Range["f7"].Value = orden.Folio;
            Globals.Hoja3.Range["f9"].Value = orden.Tecnico;
            Globals.Hoja3.Range["f11"].Value = orden.Recibio;
            int count = 0;
            if (orden.Conceptos != null)
            {
                foreach (ConceptoOrden current in orden.Conceptos)
                {
                    int offset = this.tbBody.DataBodyRange.Rows.Row + count++;
                    Globals.Hoja3.Range["A" + offset, "A" + offset].Value = current.Equipo.Ref;
                    Globals.Hoja3.Range["B" + offset, "B" + offset].Value = current.Equipo.Descripcion;
                    Globals.Hoja3.Range["C" + offset, "C" + offset].Value = current.Equipo.UnidadMedida;
                    Globals.Hoja3.Range["D" + offset, "D" + offset].Value = current.Equipo.Costo;
                    Globals.Hoja3.Range["E" + offset, "E" + offset].Value = current.Cantidad;

                    this.tbBody.ListRows.AddEx(System.Type.Missing, true);
                }
            }
        }
        private void btAdd_Click(object sender, EventArgs e)
        {
            /*
             * Combo box para insertar 
             */
            ComboBox temp;
            temp = new ComboBox();

            /*
             * Se asocia el datasource a el control
             */
            temp.DataSource =  new BindingSource(Hoja1._services, null);
            temp.DisplayMember = "_ref";

            /*
             * Evento para indicar cuando se cambia algun dato
             */
            temp.SelectedValueChanged += temp_SelectedValueChanged;

            /*
             * La llave se forma por la dirección donde va a ser insertada.
             */
            int key;

            key = this.tbBody.DataBodyRange.Rows.Row + tbBody.DataBodyRange.Rows.Count - 1;

            this.comboBoxes.Add(temp);

            /*
             * Inserta nueva fila en la tabla
             */
            this.tbBody.ListRows.AddEx(System.Type.Missing, true);
            
            this.Controls.AddControl(temp,Globals.Hoja3.Range["A" + key], "" + temp.GetHashCode());

            Button tempButton = new Button();
            tempButton.Text = "Eliminar";
            tempButton.Click +=tempButton_Click;
            this.buttons.Add(tempButton);

            this.Controls.AddControl(tempButton, Globals.Hoja3.Range["G" + key], "" + tempButton.GetHashCode());

        }
        /// <summary>
        /// Elimina un row de la del cuerpo de la orden
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void tempButton_Click(object sender, EventArgs e)
        {
            int index = this.buttons.FindIndex(b =>
            {
                return b.GetHashCode() == ((Button)sender).GetHashCode();
            });

            ComboBox temp = this.comboBoxes[index];
            this.comboBoxes.RemoveAt(index);
            this.buttons.RemoveAt(index);



            this.Controls.Remove(temp);
            this.Controls.Remove((Button)sender);

            

            index += this.tbBody.DataBodyRange.Rows.Row;

            Excel.Range range = Globals.Hoja3.get_Range(String.Format("A{0}:A{0}", index), System.Reflection.Missing.Value);
            
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

        }
        void temp_SelectedValueChanged(object sender, EventArgs e)
        {

            int offset = this.comboBoxes.FindIndex(b =>
            {
                return b.GetHashCode() == ((ComboBox)sender).GetHashCode();
            });
            offset += this.tbBody.DataBodyRange.Rows.Row;

            Globals.Hoja3.Range["A" + offset, "A" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).Ref;
            Globals.Hoja3.Range["B" + offset, "B" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).Descripcion;
            Globals.Hoja3.Range["C" + offset, "C" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).UnidadMedida;
            Globals.Hoja3.Range["D" + offset, "D" + offset].Value = ((Servicio)((ComboBox)sender).SelectedItem).Costo;



            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            switch (_actionButton)
            {
                /*
                 * Habilitar los combos para editar la orden, el boton para añdir conceptos y cambia
                 * el estado del boton de nuevo a cancelar.
                 */
                case ActionForButtonNew.NEW:
                    btNuevo.Text = "Cancelar";
                    _actionButton = ActionForButtonNew.CANCEL;
                    btGuardar.Enabled = true;
                    btAdd.Enabled = true;
                    cbEquipo.Visible = true;
                    cbEquipo.Enabled = true;

                    /*
                     * Busca la ultima orden respecto a su folio y si no hay aun ninguna
                     * asigna el folio -1 por default.
                     */
                    var maxValue = Hoja6._ordenes.Count > 0 ? Hoja6._ordenes.Max(el => el.Folio ) : 0;

                    tempOrden = new Orden();
                    tempOrden.Folio = maxValue + 1;
                    tempOrden.FechaServicio = DateTime.Now;

                    /*
                     * Pasa la orden a la hoja de excel.
                     */
                    renderOrden(tempOrden);


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

            //caca
            int count = this.Controls.Count;

            int index = this.tbBody.DataBodyRange.Rows.Row;
            int size = this.tbBody.DataBodyRange.Rows.Count;
            for (int i = 1; i < size; i++)
            {
                Excel.Range range = Globals.Hoja3.get_Range(String.Format("A{0}:A{0}", index), System.Reflection.Missing.Value);

                range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                

            }
            for (int i = this.Controls.Count - 1; i >= 6; i--)
            {
                this.Controls.RemoveAt(i);
            }

            count = this.Controls.Count;
            buttons.Clear();
            comboBoxes.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*
             * Actualizar los otros campos que si son modificables
             */
            tempOrden.CentroTrabajo = (string)this.Range["B12"].Value;
            tempOrden.Delegacion = (string)this.Range["B13"].Value;
            tempOrden.FechaServicio = (DateTime)this.Range["F5"].Value;
            tempOrden.Tecnico = (string)this.Range["F9"].Value;
            tempOrden.Recibio = (string)this.Range["F11"].Value;

            /*
             * Se esta editando
             */
            if (Hoja6._ordenes.Exists(el => el.Folio == tempOrden.Folio))
            {
                int index = Hoja6._ordenes.FindIndex(el =>
                {
                    return el.Folio == tempOrden.Folio;
                });
                /*
                 * Actualizar el equipo is es que se edito
                 */
                if (tempOrden.Equipo.NEconomico != ((Inventario)cbEquipo.SelectedItem).NEconomico)
                {
                    tempOrden.Equipo = (Inventario)cbEquipo.SelectedItem;
                }
               


                Hoja6._ordenes[index] = tempOrden;

                Globals.Hoja6.save();
            }
            /*
             * Es nueva
             */
            else
            {
                tempOrden.Equipo = (Inventario)cbEquipo.SelectedItem;
                /*
                 * Esto no funciona en tiempo de ejecución, para que pueda ser visualizado tiene que
                 * guardarse el excel y luego instanciar el excelqueryfactory.
                 * 
                var book = new ExcelQueryFactory(Globals.ThisWorkbook.FullName);
                string startRange, endRange;
                string tr = tbBody.Range.Address;
                tr = tr.Replace("$", "");
                startRange = tr.Split(':')[0];
                endRange = tr.Split(':')[1];
                var temp = (from row in book.WorksheetRange(startRange, endRange, "Captura")
                            let item = new Tuple<string, string>(row["Cantidad"].Cast<string>(), row["Clave"].Cast<string>())
                          
                           select item).ToList();
                */


                /*
                 * Mapea la tabla del cuerpo de la orden, para generar los conceptos asociados a la orden. 
                 */
                List<Tuple<string, double>> temp = new List<Tuple<string,double>>();
                string item1;
                double item2;

                Excel.Range body = this.tbBody.DataBodyRange;

                for (int i = 1; i < body.Rows.Count; i++)
                {
                    item1 = (string)body.Cells[i, 1].value;
                    item2 = (double)body.Cells[i, 5].value;
                    temp.Add(new Tuple<string, double>(item1, item2));
                }


                /*
                 * Una vez mapeados los elementos, recorre la lista creada y los convierte en objetos
                 */
                var res = temp.Select(element =>
                {
                    ConceptoOrden tempConcepto = new ConceptoOrden();

                    /*
                     * Busca el servicio asociado en la lista de servicios para
                     * añadirlo en el equipo que se esta creando.
                     */
                    tempConcepto.Equipo = Hoja1._services.Where(el =>  el.Ref == element.Item1 ).FirstOrDefault();
                    tempConcepto.Orden = tempOrden.Folio;
                    tempConcepto.Cantidad = element.Item2;
                    tempConcepto.SubTotal = tempConcepto.Cantidad * tempConcepto.Equipo.Costo;

                    return tempConcepto;
                });

                tempOrden.Conceptos = res.ToList();

                /*
                 * Se añaden los objetos que se acaban de crear a las colecciones globales.
                 */
                Hoja7._conceptos.AddRange(tempOrden.Conceptos);
                Hoja6._ordenes.Add(tempOrden);
                
                Globals.Hoja6.save();
                Globals.Hoja7.save();

                button1_Click(null, null);
            }
        }

       

        

    }
}
