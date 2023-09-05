using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LeerExcel
{
    public partial class ReporteHoja : Form
    {
        public ReporteHoja()
        {
            InitializeComponent();
        }

        private void ReporteHoja_Load(object sender, EventArgs e)
        {
            
            // TODO: esta línea de código carga datos en la tabla 'dataSet1.DataTable' Puede moverla o quitarla según sea necesario.
            this.dataTableTableAdapter.getHojasVentana(this.dataSet1.DataTable);

            this.reportViewer1.RefreshReport();
            this.reportViewer1.RefreshReport();
            this.reportViewer1.RefreshReport();
        }
    }
}
