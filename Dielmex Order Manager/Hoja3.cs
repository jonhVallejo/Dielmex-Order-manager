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

        private List<ComboBox> _comboBoxes;
        private List<Button> _buttons;

        private Order _tempOrder;

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
            cbEquipo.DataSource = Hoja2._inventary;
            cbEquipo.Visible = false;

            cbOrdenNumber.DataSource = Hoja6._orders;
            cbOrdenNumber.DisplayMember = "Folio";

            cbOrdenNumber.SelectedValueChanged += cbOrdenNumber_SelectedValueChanged;

            cbEquipo.SelectedValueChanged += cbEquipo_SelectedValueChanged;

            _comboBoxes = new List<ComboBox>();
            _buttons = new List<Button>();

            _offsetForComboxInTable += this.Controls.Count;
            _firstIndexForTable = 19;


            /*
             * Take the name of the dynamic value
             */
            string dynamicValue = this.Range["A14"].Value;

            /*
             * If the dynamic value is specified, then we can search. 
             */
            if (dynamicValue != null && dynamicValue.Length > 0 && Hoja2._dynamicColumNames.Exists(culumnName => culumnName == dynamicValue))
            {
                /*
                 * dynamicValue is the tag for search in the columns of the inventary. Before of search the data, we go to modify the value
                 * of the specified cell using the function buscar (Excel Spanish version).
                 */
                string formula = String.Format("=BUSCAR(A7, Tabla2[NECONOMICO], Tabla2[{0}])", dynamicValue);

                this.Range["B14"].Formula = formula;
            }
            else
            {
                this.Range["B14"].Value = "";
            }
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

            
            _tempOrder = (Order)cbOrdenNumber.SelectedItem;

            renderOrder(_tempOrder);
        }

        void cbEquipo_SelectedValueChanged(object sender, EventArgs e)
        {
            Globals.Hoja3.Cells[7, 1].value = ((Inventary)cbEquipo.SelectedItem).SerialNumber;
            Globals.Hoja3.Cells[9, 2].value = ((Inventary)cbEquipo.SelectedItem).Brand;
            Globals.Hoja3.Cells[10, 2].value = ((Inventary)cbEquipo.SelectedItem).Model;
            Globals.Hoja3.Cells[11, 2].value = ((Inventary)cbEquipo.SelectedItem).Type;
            this.Range["B12"].Value = ((Inventary)cbEquipo.SelectedItem).WorkCentre;
            this.Range["B13"].Value = ((Inventary)cbEquipo.SelectedItem).Workplace;
            
        }

        private void renderOrder(Order order)
        {
            if (order.Equipment != null)
            {
                cbEquipo.SelectedIndex = Hoja2._inventary.FindIndex(el => { return el.SerialNumber == order.Equipment.SerialNumber; });
            }
            cbOrdenNumber.Enabled = false;
            cbOrdenNumber.Visible = false;

            if (order.Equipment != null)
            {

                Globals.Hoja3.Range["b12"].Value = order.Equipment.WorkCentre;
                Globals.Hoja3.Range["b13"].Value = order.Equipment.Workplace;
            }
            Globals.Hoja3.Range["f5"].Value = order.ServiceDate;
            Globals.Hoja3.Range["f7"].Value = order.Folio;
            Globals.Hoja3.Range["f9"].Value = order.Expert;
            Globals.Hoja3.Range["f11"].Value = order.ReceivedBy;
            int count = 0;
            if (order.OrderItems != null)
            {
                foreach (OrderItem current in order.OrderItems)
                {
                    int offset = this.tbBody.DataBodyRange.Rows.Row + count++;
                    Globals.Hoja3.Range["A" + offset, "A" + offset].Value = current.Equipment.ServiceId;
                    Globals.Hoja3.Range["B" + offset, "B" + offset].Value = current.Equipment.Description;
                    Globals.Hoja3.Range["C" + offset, "C" + offset].Value = current.Equipment.UnitOfMeasurement;
                    Globals.Hoja3.Range["D" + offset, "D" + offset].Value = current.Equipment.Cost;
                    Globals.Hoja3.Range["E" + offset, "E" + offset].Value = current.Quantity;

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

            this._comboBoxes.Add(temp);

            /*
             * Inserta nueva fila en la tabla
             */
            this.tbBody.ListRows.AddEx(System.Type.Missing, true);
            
            this.Controls.AddControl(temp,Globals.Hoja3.Range["A" + key], "" + temp.GetHashCode());

            Button tempButton = new Button();
            tempButton.Text = "Eliminar";
            tempButton.Click +=tempButton_Click;
            this._buttons.Add(tempButton);

            this.Controls.AddControl(tempButton, Globals.Hoja3.Range["G" + key], "" + tempButton.GetHashCode());

        }
        /// <summary>
        /// Elimina un row de la del cuerpo de la order
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void tempButton_Click(object sender, EventArgs e)
        {
            int index = this._buttons.FindIndex(b =>
            {
                return b.GetHashCode() == ((Button)sender).GetHashCode();
            });

            ComboBox temp = this._comboBoxes[index];
            this._comboBoxes.RemoveAt(index);
            this._buttons.RemoveAt(index);



            this.Controls.Remove(temp);
            this.Controls.Remove((Button)sender);

            

            index += this.tbBody.DataBodyRange.Rows.Row;

            Excel.Range range = Globals.Hoja3.get_Range(String.Format("A{0}:A{0}", index), System.Reflection.Missing.Value);
            
            range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);

        }
        void temp_SelectedValueChanged(object sender, EventArgs e)
        {

            int offset = this._comboBoxes.FindIndex(b =>
            {
                return b.GetHashCode() == ((ComboBox)sender).GetHashCode();
            });
            offset += this.tbBody.DataBodyRange.Rows.Row;

            Globals.Hoja3.Range["A" + offset, "A" + offset].Value = ((Service)((ComboBox)sender).SelectedItem).ServiceId;
            Globals.Hoja3.Range["B" + offset, "B" + offset].Value = ((Service)((ComboBox)sender).SelectedItem).Description;
            Globals.Hoja3.Range["C" + offset, "C" + offset].Value = ((Service)((ComboBox)sender).SelectedItem).UnitOfMeasurement;
            Globals.Hoja3.Range["D" + offset, "D" + offset].Value = ((Service)((ComboBox)sender).SelectedItem).Cost;



            
        }
        private void button1_Click(object sender, EventArgs e)
        {
            switch (_actionButton)
            {
                /*
                 * Habilitar los combos para editar la order, el boton para añdir conceptos y cambia
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
                     * Busca la ultima order respecto a su folio y si no hay aun ninguna
                     * asigna el folio -1 por default.
                     */
                    var maxValue = Hoja6._orders.Count > 0 ? Hoja6._orders.Max(el => el.Folio ) : 0;

                    _tempOrder = new Order();
                    _tempOrder.Folio = maxValue + 1;
                    _tempOrder.ServiceDate = DateTime.Now;

                    /*
                     * Pasa la order a la hoja de excel.
                     */
                    renderOrder(_tempOrder);


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
            _buttons.Clear();
            _comboBoxes.Clear();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            /*
             * Actualizar los otros campos que si son modificables
             */
            _tempOrder.ServiceDate = (DateTime)this.Range["F5"].Value;
            _tempOrder.Expert = (string)this.Range["F9"].Value;
            _tempOrder.ReceivedBy = (string)this.Range["F11"].Value;

            /*
             * Se esta editando
             */
            if (Hoja6._orders.Exists(el => el.Folio == _tempOrder.Folio))
            {
                int index = Hoja6._orders.FindIndex(el =>
                {
                    return el.Folio == _tempOrder.Folio;
                });
                /*
                 * Actualizar el equipo is es que se edito
                 */
                if (_tempOrder.Equipment.SerialNumber != ((Inventary)cbEquipo.SelectedItem).SerialNumber)
                {
                    _tempOrder.Equipment = (Inventary)cbEquipo.SelectedItem;
                }
               


                Hoja6._orders[index] = _tempOrder;

                Globals.Hoja6.save();
            }
            /*
             * Es nueva
             */
            else
            {
                _tempOrder.Equipment = (Inventary)cbEquipo.SelectedItem;
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
                 * Mapea la tabla del cuerpo de la order, para generar los conceptos asociados a la order. 
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
                    OrderItem tempConcepto = new OrderItem();

                    /*
                     * Busca el Service asociado en la lista de Services para
                     * añadirlo en el equipo que se esta creando.
                     */
                    tempConcepto.Equipment = Hoja1._services.Where(el =>  el.ServiceId == element.Item1 ).FirstOrDefault();
                    tempConcepto.OrderNumber = _tempOrder.Folio;
                    tempConcepto.Quantity = element.Item2;
                    tempConcepto.SubTotal = tempConcepto.Quantity * tempConcepto.Equipment.Cost;

                    return tempConcepto;
                });

                _tempOrder.OrderItems = res.ToList();

                /*
                 * Se añaden los objetos que se acaban de crear a las colecciones globales.
                 */
                Hoja7._itemsOrder.AddRange(_tempOrder.OrderItems);
                Hoja6._orders.Add(_tempOrder);
                
                Globals.Hoja6.save();
                Globals.Hoja7.save();

                button1_Click(null, null);
            }
        }

       

        

    }
}
