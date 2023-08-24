using MySql.Data.MySqlClient;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LeerExcel
{
    public partial class frmEtiquetas : Form
    {
        public frmEtiquetas()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
        }

        private void btnCargar_Click(object sender, EventArgs e)
        {
            string ruta = lblRuta.Text;
            if (ruta == "Selecciona un archivo" || ruta == "")
            {
                MessageBox.Show("Error debes seleccionar un archivo");
            }
            else
            {
                SLDocument sl = new SLDocument(@ruta);
                SLWorksheetStatistics propiedades = sl.GetWorksheetStatistics();

                int ultimaFila = propiedades.EndRowIndex;

                MySqlConnection conexionDB = Conexion.conexion();
                try
                {
                    conexionDB.Open();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al conectar con la base de datos");
                }
                finally
                {
                    conexionDB.Close();
                }

                string error = "";
                string cliente = sl.GetCellValueAsString("A" + 1);
                string no_orden = sl.GetCellValueAsString("B" + 1);
                string no_ventana = sl.GetCellValueAsString("C" + 1);
                string descripcion_articulo = sl.GetCellValueAsString("D" + 1);
                string no_corte = sl.GetCellValueAsString("E" + 1);
                string medida = sl.GetCellValueAsString("F" + 1);
                string posicion = sl.GetCellValueAsString("G" + 1);
                string id_ventana = sl.GetCellValueAsString("H" + 1);

                if (existeOrden(no_orden))
                {
                    MessageBox.Show("Ya existe la orden: " + no_orden + "\n");
                }
                else if (cliente == "" || no_corte == "" || no_ventana == "" || descripcion_articulo == "" || no_corte == "" || medida == "" || posicion == "" || id_ventana == "")
                {
                    MessageBox.Show("Algunas de las celdas estan vacías, verifique el formato");
                }
                else
                {
                    for (int x = 1; x < ultimaFila; x++)
                    {
                        string sql = "INSERT INTO ordenes(CLIENTE,NO_ORDEN, NO_VENTANA,DESCRIPCION_ARTICULO, NO_CORTE, MEDIDA," +
                        "POSICION, ID_VENTANA )VALUES(@CLIENTE,@NO_ORDEN, @NO_VENTANA, @DESCRIPCION_ARTICULO, @NO_CORTE, @MEDIDA," +
                        "@POSICION, @ID_VENTANA)";
                        try
                        {
                            MySqlCommand comando = new MySqlCommand(sql, conexionDB);
                            comando.Parameters.AddWithValue("@CLIENTE", sl.GetCellValueAsString("A" + x));
                            comando.Parameters.AddWithValue("@NO_ORDEN", sl.GetCellValueAsString("B" + x));
                            comando.Parameters.AddWithValue("@NO_VENTANA", sl.GetCellValueAsString("C" + x));
                            comando.Parameters.AddWithValue("@DESCRIPCION_ARTICULO", sl.GetCellValueAsString("D" + x));
                            comando.Parameters.AddWithValue("@NO_CORTE", sl.GetCellValueAsString("E" + x));
                            comando.Parameters.AddWithValue("@MEDIDA", sl.GetCellValueAsString("F" + x));
                            comando.Parameters.AddWithValue("@POSICION", sl.GetCellValueAsString("G" + x));
                            comando.Parameters.AddWithValue("@ID_VENTANA", sl.GetCellValueAsString("H" + x));
                            comando.ExecuteNonQuery();
                        }
                        catch (MySqlException ex)
                        {

                            MessageBox.Show(ex.Message);
                        }
                    }
                    MessageBox.Show("Datos Cargados");
                }
            }
        }

        private bool existeOrden(string no_orden) {
            MySqlConnection conexionDB = Conexion.conexion();
            conexionDB.Open();
            string sql = "SELECT COUNT(NO_ORDEN) FROM ordenes WHERE NO_ORDEN = '" + no_orden + "'";
            MySqlCommand comando = new MySqlCommand(sql, conexionDB);

            int cantidad = Convert.ToInt32(comando.ExecuteScalar());

            if (cantidad > 0)
            {
                return true;
            }
            else { 
                return false;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {


        }

        

        private void bntBuscar_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                lblRuta.Text = openFileDialog1.FileName;
            }
            else
            {
                lblRuta.Text = "Selecciona un archivo";
            }
        }

        private void btnEjecutarEM_Click(object sender, EventArgs e)
        {
            String numeroOrden = txtNumOrden.Text.ToUpper();
            if (numeroOrden == "")
            {
                MessageBox.Show("Ingresa un número de orden");
            }
            else {
                MessageBox.Show("Este es el numero de orden: " + numeroOrden);
                this.dataSet11.Clear();
                using (MemoryStream ms = new MemoryStream()) { 
                
                
                }
            }
            
        }
    }
}
