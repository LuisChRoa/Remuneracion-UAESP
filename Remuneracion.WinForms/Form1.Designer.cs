namespace Remuneracion.WinForms
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private System.Windows.Forms.ComboBox cmbPeriodo;
        private System.Windows.Forms.Label lblPeriodo;
        private System.Windows.Forms.Label lblCarpeta;
        private System.Windows.Forms.TextBox txtCarpetaFuentes;
        private System.Windows.Forms.Button btnSeleccionarCarpeta;
        private System.Windows.Forms.Label lblPlantilla;
        private System.Windows.Forms.TextBox txtPlantilla;
        private System.Windows.Forms.Button btnSeleccionarPlantilla;
        private System.Windows.Forms.Label lblAse;
        private System.Windows.Forms.ComboBox cmbAse;
        private System.Windows.Forms.Button btnEjecutar;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.Label lblEstado;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialogPlantilla;

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            cmbPeriodo = new System.Windows.Forms.ComboBox();
            lblPeriodo = new System.Windows.Forms.Label();
            lblCarpeta = new System.Windows.Forms.Label();
            txtCarpetaFuentes = new System.Windows.Forms.TextBox();
            btnSeleccionarCarpeta = new System.Windows.Forms.Button();
            lblPlantilla = new System.Windows.Forms.Label();
            txtPlantilla = new System.Windows.Forms.TextBox();
            btnSeleccionarPlantilla = new System.Windows.Forms.Button();
            lblAse = new System.Windows.Forms.Label();
            cmbAse = new System.Windows.Forms.ComboBox();
            btnEjecutar = new System.Windows.Forms.Button();
            progressBar = new System.Windows.Forms.ProgressBar();
            txtLog = new System.Windows.Forms.TextBox();
            lblEstado = new System.Windows.Forms.Label();
            folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            openFileDialogPlantilla = new System.Windows.Forms.OpenFileDialog();
            SuspendLayout();
            // 
            // lblPeriodo
            // 
            lblPeriodo.AutoSize = true;
            lblPeriodo.Location = new System.Drawing.Point(12, 15);
            lblPeriodo.Name = "lblPeriodo";
            lblPeriodo.Size = new System.Drawing.Size(47, 15);
            lblPeriodo.Text = "Periodo:";
            // 
            // cmbPeriodo
            // 
            cmbPeriodo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbPeriodo.FormattingEnabled = true;
            cmbPeriodo.Items.AddRange(new object[]
            {
                "202605-1",
                "202605-2",
                "202606-1",
                "202606-2"
            });
            cmbPeriodo.Location = new System.Drawing.Point(108, 12);
            cmbPeriodo.Name = "cmbPeriodo";
            cmbPeriodo.Size = new System.Drawing.Size(180, 23);
            cmbPeriodo.SelectedIndex = 0;
            // 
            // lblCarpeta
            // 
            lblCarpeta.AutoSize = true;
            lblCarpeta.Location = new System.Drawing.Point(12, 46);
            lblCarpeta.Name = "lblCarpeta";
            lblCarpeta.Size = new System.Drawing.Size(82, 15);
            lblCarpeta.Text = "Carpeta fuentes:";
            // 
            // txtCarpetaFuentes
            // 
            txtCarpetaFuentes.Location = new System.Drawing.Point(108, 43);
            txtCarpetaFuentes.Name = "txtCarpetaFuentes";
            txtCarpetaFuentes.Size = new System.Drawing.Size(420, 23);
            // 
            // btnSeleccionarCarpeta
            // 
            btnSeleccionarCarpeta.Location = new System.Drawing.Point(534, 42);
            btnSeleccionarCarpeta.Name = "btnSeleccionarCarpeta";
            btnSeleccionarCarpeta.Size = new System.Drawing.Size(110, 23);
            btnSeleccionarCarpeta.Text = "Seleccionar...";
            btnSeleccionarCarpeta.UseVisualStyleBackColor = true;
            btnSeleccionarCarpeta.Click += new System.EventHandler(btnSeleccionarCarpeta_Click);
            // 
            // lblPlantilla
            // 
            lblPlantilla.AutoSize = true;
            lblPlantilla.Location = new System.Drawing.Point(12, 77);
            lblPlantilla.Name = "lblPlantilla";
            lblPlantilla.Size = new System.Drawing.Size(57, 15);
            lblPlantilla.Text = "Plantilla:";
            // 
            // txtPlantilla
            // 
            txtPlantilla.Location = new System.Drawing.Point(108, 74);
            txtPlantilla.Name = "txtPlantilla";
            txtPlantilla.Size = new System.Drawing.Size(420, 23);
            // 
            // btnSeleccionarPlantilla
            // 
            btnSeleccionarPlantilla.Location = new System.Drawing.Point(534, 73);
            btnSeleccionarPlantilla.Name = "btnSeleccionarPlantilla";
            btnSeleccionarPlantilla.Size = new System.Drawing.Size(110, 23);
            btnSeleccionarPlantilla.Text = "Seleccionar...";
            btnSeleccionarPlantilla.UseVisualStyleBackColor = true;
            btnSeleccionarPlantilla.Click += new System.EventHandler(btnSeleccionarPlantilla_Click);
            // 
            // lblAse
            // 
            lblAse.AutoSize = true;
            lblAse.Location = new System.Drawing.Point(12, 108);
            lblAse.Name = "lblAse";
            lblAse.Size = new System.Drawing.Size(28, 15);
            lblAse.Text = "ASE:";
            // 
            // cmbAse
            // 
            cmbAse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbAse.FormattingEnabled = true;
            cmbAse.Items.AddRange(new object[]
            {
                "Promoambiental",
                "LIME",
                "Ciudad Limpia",
                "Bogotá Limpia",
                "Área Limpia"
            });
            cmbAse.Location = new System.Drawing.Point(108, 105);
            cmbAse.Name = "cmbAse";
            cmbAse.Size = new System.Drawing.Size(180, 23);
            cmbAse.SelectedIndex = 0;
            // 
            // btnEjecutar
            // 
            btnEjecutar.Location = new System.Drawing.Point(12, 140);
            btnEjecutar.Name = "btnEjecutar";
            btnEjecutar.Size = new System.Drawing.Size(120, 30);
            btnEjecutar.Text = "Ejecutar";
            btnEjecutar.UseVisualStyleBackColor = true;
            btnEjecutar.Click += new System.EventHandler(btnEjecutar_Click);
            // 
            // progressBar
            // 
            progressBar.Location = new System.Drawing.Point(150, 144);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(200, 23);
            progressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            progressBar.Visible = false;
            // 
            // txtLog
            // 
            txtLog.Location = new System.Drawing.Point(12, 180);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            txtLog.Size = new System.Drawing.Size(632, 320);
            // 
            // lblEstado
            // 
            lblEstado.AutoSize = true;
            lblEstado.Location = new System.Drawing.Point(12, 515);
            lblEstado.Name = "lblEstado";
            lblEstado.Size = new System.Drawing.Size(34, 15);
            lblEstado.Text = "Listo";
            // 
            // openFileDialogPlantilla
            // 
            openFileDialogPlantilla.Filter = "Excel|*.xlsx";
            // 
            // Form1
            // 
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(656, 540);
            Controls.Add(lblPeriodo);
            Controls.Add(cmbPeriodo);
            Controls.Add(lblCarpeta);
            Controls.Add(txtCarpetaFuentes);
            Controls.Add(btnSeleccionarCarpeta);
            Controls.Add(lblPlantilla);
            Controls.Add(txtPlantilla);
            Controls.Add(btnSeleccionarPlantilla);
            Controls.Add(lblAse);
            Controls.Add(cmbAse);
            Controls.Add(btnEjecutar);
            Controls.Add(progressBar);
            Controls.Add(txtLog);
            Controls.Add(lblEstado);
            Name = "Form1";
            Text = "Remuneración Quincenal UAESP";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
