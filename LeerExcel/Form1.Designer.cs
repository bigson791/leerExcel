namespace LeerExcel
{
    partial class frmEtiquetas
    {
        /// <summary>
        /// Variable del diseñador necesaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Limpiar los recursos que se estén usando.
        /// </summary>
        /// <param name="disposing">true si los recursos administrados se deben desechar; false en caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Código generado por el Diseñador de Windows Forms

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido de este método con el editor de código.
        /// </summary>
        private void InitializeComponent()
        {
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.bntBuscar = new System.Windows.Forms.Button();
            this.lblRuta = new System.Windows.Forms.Label();
            this.btnCargar = new System.Windows.Forms.Button();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnEjecutarEM = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtNumOrden = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btnEjecutarEH = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.btnEjecutarEO = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.dataSet11 = new LeerExcel.DataSet1();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet11)).BeginInit();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Archivo Excel (*xlsx)|*.xlsx";
            // 
            // tabControl1
            // 
            this.tabControl1.AccessibleName = "";
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Controls.Add(this.tabPage4);
            this.tabControl1.Location = new System.Drawing.Point(34, 43);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(704, 366);
            this.tabControl1.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.progressBar1);
            this.tabPage1.Controls.Add(this.bntBuscar);
            this.tabPage1.Controls.Add(this.lblRuta);
            this.tabPage1.Controls.Add(this.btnCargar);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(696, 340);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Ingreso Datos";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(94, 156);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(500, 23);
            this.progressBar1.TabIndex = 7;
            // 
            // bntBuscar
            // 
            this.bntBuscar.Location = new System.Drawing.Point(362, 108);
            this.bntBuscar.Name = "bntBuscar";
            this.bntBuscar.Size = new System.Drawing.Size(48, 30);
            this.bntBuscar.TabIndex = 6;
            this.bntBuscar.Text = "...";
            this.bntBuscar.UseVisualStyleBackColor = true;
            this.bntBuscar.Click += new System.EventHandler(this.bntBuscar_Click);
            // 
            // lblRuta
            // 
            this.lblRuta.AutoSize = true;
            this.lblRuta.Location = new System.Drawing.Point(91, 117);
            this.lblRuta.Name = "lblRuta";
            this.lblRuta.Size = new System.Drawing.Size(113, 13);
            this.lblRuta.TabIndex = 5;
            this.lblRuta.Text = "Selecciona un archivo";
            // 
            // btnCargar
            // 
            this.btnCargar.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.btnCargar.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCargar.Location = new System.Drawing.Point(416, 108);
            this.btnCargar.Name = "btnCargar";
            this.btnCargar.Size = new System.Drawing.Size(178, 30);
            this.btnCargar.TabIndex = 4;
            this.btnCargar.Text = "Cargar Archivo";
            this.btnCargar.UseVisualStyleBackColor = true;
            this.btnCargar.Click += new System.EventHandler(this.btnCargar_Click);
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.btnEjecutarEM);
            this.tabPage2.Controls.Add(this.label1);
            this.tabPage2.Controls.Add(this.txtNumOrden);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(696, 340);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Etiqueta Marco";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // btnEjecutarEM
            // 
            this.btnEjecutarEM.Location = new System.Drawing.Point(216, 87);
            this.btnEjecutarEM.Name = "btnEjecutarEM";
            this.btnEjecutarEM.Size = new System.Drawing.Size(95, 23);
            this.btnEjecutarEM.TabIndex = 2;
            this.btnEjecutarEM.Text = "Ejecutar";
            this.btnEjecutarEM.UseVisualStyleBackColor = true;
            this.btnEjecutarEM.Click += new System.EventHandler(this.btnEjecutarEM_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(24, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "No. Orden";
            // 
            // txtNumOrden
            // 
            this.txtNumOrden.Location = new System.Drawing.Point(86, 90);
            this.txtNumOrden.Name = "txtNumOrden";
            this.txtNumOrden.Size = new System.Drawing.Size(113, 20);
            this.txtNumOrden.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.btnEjecutarEH);
            this.tabPage3.Controls.Add(this.label2);
            this.tabPage3.Controls.Add(this.textBox1);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(696, 340);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Etiqueta Hoja";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btnEjecutarEH
            // 
            this.btnEjecutarEH.Location = new System.Drawing.Point(224, 81);
            this.btnEjecutarEH.Name = "btnEjecutarEH";
            this.btnEjecutarEH.Size = new System.Drawing.Size(95, 23);
            this.btnEjecutarEH.TabIndex = 5;
            this.btnEjecutarEH.Text = "Ejecutar";
            this.btnEjecutarEH.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(32, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "No. Orden";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(94, 84);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(113, 20);
            this.textBox1.TabIndex = 3;
            // 
            // tabPage4
            // 
            this.tabPage4.Controls.Add(this.btnEjecutarEO);
            this.tabPage4.Controls.Add(this.label3);
            this.tabPage4.Controls.Add(this.textBox2);
            this.tabPage4.Location = new System.Drawing.Point(4, 22);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(696, 340);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "Etiqueta Otros";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // btnEjecutarEO
            // 
            this.btnEjecutarEO.Location = new System.Drawing.Point(221, 83);
            this.btnEjecutarEO.Name = "btnEjecutarEO";
            this.btnEjecutarEO.Size = new System.Drawing.Size(95, 23);
            this.btnEjecutarEO.TabIndex = 8;
            this.btnEjecutarEO.Text = "Ejecutar";
            this.btnEjecutarEO.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(29, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(56, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "No. Orden";
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(91, 86);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(113, 20);
            this.textBox2.TabIndex = 6;
            // 
            // dataSet11
            // 
            this.dataSet11.DataSetName = "DataSet1";
            this.dataSet11.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // frmEtiquetas
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tabControl1);
            this.Name = "frmEtiquetas";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Etiquetas";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataSet11)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button bntBuscar;
        private System.Windows.Forms.Label lblRuta;
        private System.Windows.Forms.Button btnCargar;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.Button btnEjecutarEH;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Button btnEjecutarEO;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnEjecutarEM;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtNumOrden;
        private DataSet1 dataSet11;
    }
}

