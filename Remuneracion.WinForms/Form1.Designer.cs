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

        // --- GroupBoxes ---
        private System.Windows.Forms.GroupBox grpPeriodo;
        private System.Windows.Forms.GroupBox grpArchivos;
        private System.Windows.Forms.GroupBox grpEjecucion;

        // --- Período ---
        private System.Windows.Forms.Label lblAnio;
        private System.Windows.Forms.ComboBox cmbAnio;
        private System.Windows.Forms.Label lblMes;
        private System.Windows.Forms.ComboBox cmbMes;
        private System.Windows.Forms.Label lblQuincena;
        private System.Windows.Forms.ComboBox cmbQuincena;

        // --- Archivos ---
        private System.Windows.Forms.Label lblCarpeta;
        private System.Windows.Forms.TextBox txtCarpetaFuentes;
        private System.Windows.Forms.Button btnSeleccionarCarpeta;
        private System.Windows.Forms.Label lblPlantilla;
        private System.Windows.Forms.TextBox txtPlantilla;
        private System.Windows.Forms.Button btnSeleccionarPlantilla;
        private System.Windows.Forms.Label lblCarpetaSalida = new System.Windows.Forms.Label();
        private System.Windows.Forms.TextBox txtCarpetaSalida = new System.Windows.Forms.TextBox();
        private System.Windows.Forms.Button btnSeleccionarSalida = new System.Windows.Forms.Button();

        // --- Ejecución ---
        private System.Windows.Forms.Label lblAse;
        private System.Windows.Forms.ComboBox cmbAse;
        private System.Windows.Forms.CheckBox chkCincoAse;
        private System.Windows.Forms.Button btnEjecutar;
        private System.Windows.Forms.ProgressBar progressBar;

        // --- Log ---
        private System.Windows.Forms.TextBox txtLog;

        // --- StatusStrip ---
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ToolStripStatusLabel toolStripVersion;

        // --- Diálogos ---
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.OpenFileDialog openFileDialogPlantilla;

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();

            // === GroupBoxes ===
            grpPeriodo = new System.Windows.Forms.GroupBox();
            grpArchivos = new System.Windows.Forms.GroupBox();
            grpEjecucion = new System.Windows.Forms.GroupBox();

            // === Período ===
            lblAnio = new System.Windows.Forms.Label();
            cmbAnio = new System.Windows.Forms.ComboBox();
            lblMes = new System.Windows.Forms.Label();
            cmbMes = new System.Windows.Forms.ComboBox();
            lblQuincena = new System.Windows.Forms.Label();
            cmbQuincena = new System.Windows.Forms.ComboBox();

            // === Archivos ===
            lblCarpeta = new System.Windows.Forms.Label();
            txtCarpetaFuentes = new System.Windows.Forms.TextBox();
            btnSeleccionarCarpeta = new System.Windows.Forms.Button();
            lblPlantilla = new System.Windows.Forms.Label();
            txtPlantilla = new System.Windows.Forms.TextBox();
            btnSeleccionarPlantilla = new System.Windows.Forms.Button();

            // === Ejecución ===
            lblAse = new System.Windows.Forms.Label();
            cmbAse = new System.Windows.Forms.ComboBox();
            chkCincoAse = new System.Windows.Forms.CheckBox();
            btnEjecutar = new System.Windows.Forms.Button();
            progressBar = new System.Windows.Forms.ProgressBar();

            // === Log ===
            txtLog = new System.Windows.Forms.TextBox();

            // === StatusStrip ===
            statusStrip = new System.Windows.Forms.StatusStrip();
            toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripVersion = new System.Windows.Forms.ToolStripStatusLabel();

            // === Diálogos ===
            folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            openFileDialogPlantilla = new System.Windows.Forms.OpenFileDialog();
            openFileDialogPlantilla.Filter = "Archivos Excel (*.xlsx)|*.xlsx";

            statusStrip.SuspendLayout();
            grpPeriodo.SuspendLayout();
            grpArchivos.SuspendLayout();
            grpEjecucion.SuspendLayout();
            SuspendLayout();

            // ============================
            // grpPeriodo
            // ============================
            grpPeriodo.Controls.Add(lblAnio);
            grpPeriodo.Controls.Add(cmbAnio);
            grpPeriodo.Controls.Add(lblMes);
            grpPeriodo.Controls.Add(cmbMes);
            grpPeriodo.Controls.Add(lblQuincena);
            grpPeriodo.Controls.Add(cmbQuincena);
            grpPeriodo.Location = new System.Drawing.Point(12, 12);
            grpPeriodo.Name = "grpPeriodo";
            grpPeriodo.Size = new System.Drawing.Size(676, 54);
            grpPeriodo.TabIndex = 0;
            grpPeriodo.TabStop = false;
            grpPeriodo.Text = "Período";
            //
            // lblAnio
            //
            lblAnio.AutoSize = true;
            lblAnio.Location = new System.Drawing.Point(15, 22);
            lblAnio.Name = "lblAnio";
            lblAnio.Size = new System.Drawing.Size(33, 15);
            lblAnio.Text = "Año:";
            //
            // cmbAnio
            //
            cmbAnio.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbAnio.FormattingEnabled = true;
            cmbAnio.Location = new System.Drawing.Point(55, 19);
            cmbAnio.Name = "cmbAnio";
            cmbAnio.Size = new System.Drawing.Size(80, 23);
            cmbAnio.TabIndex = 1;
            //
            // lblMes
            //
            lblMes.AutoSize = true;
            lblMes.Location = new System.Drawing.Point(150, 22);
            lblMes.Name = "lblMes";
            lblMes.Size = new System.Drawing.Size(35, 15);
            lblMes.Text = "Mes:";
            //
            // cmbMes
            //
            cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbMes.FormattingEnabled = true;
            cmbMes.Location = new System.Drawing.Point(190, 19);
            cmbMes.Name = "cmbMes";
            cmbMes.Size = new System.Drawing.Size(120, 23);
            cmbMes.TabIndex = 2;
            //
            // lblQuincena
            //
            lblQuincena.AutoSize = true;
            lblQuincena.Location = new System.Drawing.Point(325, 22);
            lblQuincena.Name = "lblQuincena";
            lblQuincena.Size = new System.Drawing.Size(64, 15);
            lblQuincena.Text = "Quincena:";
            //
            // cmbQuincena
            //
            cmbQuincena.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbQuincena.FormattingEnabled = true;
            cmbQuincena.Location = new System.Drawing.Point(395, 19);
            cmbQuincena.Name = "cmbQuincena";
            cmbQuincena.Size = new System.Drawing.Size(110, 23);
            cmbQuincena.TabIndex = 3;

            // ============================
            // grpArchivos
            // ============================
            grpArchivos.Controls.Add(lblCarpeta);
            grpArchivos.Controls.Add(txtCarpetaFuentes);
            grpArchivos.Controls.Add(btnSeleccionarCarpeta);
            grpArchivos.Controls.Add(lblPlantilla);
            grpArchivos.Controls.Add(txtPlantilla);
            grpArchivos.Controls.Add(btnSeleccionarPlantilla);
            grpArchivos.Controls.Add(lblCarpetaSalida);
            grpArchivos.Controls.Add(txtCarpetaSalida);
            grpArchivos.Controls.Add(btnSeleccionarSalida);
            grpArchivos.Location = new System.Drawing.Point(12, 74);
            grpArchivos.Name = "grpArchivos";
            grpArchivos.Size = new System.Drawing.Size(676, 102);
            grpArchivos.TabIndex = 1;
            grpArchivos.TabStop = false;
            grpArchivos.Text = "Archivos";
            //
            // lblCarpeta
            //
            lblCarpeta.AutoSize = true;
            lblCarpeta.Location = new System.Drawing.Point(15, 22);
            lblCarpeta.Name = "lblCarpeta";
            lblCarpeta.Size = new System.Drawing.Size(88, 15);
            lblCarpeta.Text = "Carpeta fuentes:";
            //
            // txtCarpetaFuentes
            //
            txtCarpetaFuentes.Location = new System.Drawing.Point(105, 19);
            txtCarpetaFuentes.Name = "txtCarpetaFuentes";
            txtCarpetaFuentes.ReadOnly = true;
            txtCarpetaFuentes.Size = new System.Drawing.Size(420, 23);
            txtCarpetaFuentes.TabIndex = 0;
            //
            // btnSeleccionarCarpeta
            //
            btnSeleccionarCarpeta.Location = new System.Drawing.Point(535, 18);
            btnSeleccionarCarpeta.Name = "btnSeleccionarCarpeta";
            btnSeleccionarCarpeta.Size = new System.Drawing.Size(110, 28);
            btnSeleccionarCarpeta.TabIndex = 1;
            btnSeleccionarCarpeta.Text = "Seleccionar...";
            btnSeleccionarCarpeta.UseVisualStyleBackColor = true;
            btnSeleccionarCarpeta.Click += new System.EventHandler(btnSeleccionarCarpeta_Click);
            //
            // lblPlantilla
            //
            lblPlantilla.AutoSize = true;
            lblPlantilla.Location = new System.Drawing.Point(15, 52);
            lblPlantilla.Name = "lblPlantilla";
            lblPlantilla.Size = new System.Drawing.Size(57, 15);
            lblPlantilla.Text = "Plantilla:";
            //
            // txtPlantilla
            //
            txtPlantilla.Location = new System.Drawing.Point(105, 49);
            txtPlantilla.Name = "txtPlantilla";
            txtPlantilla.ReadOnly = true;
            txtPlantilla.Size = new System.Drawing.Size(420, 23);
            txtPlantilla.TabIndex = 2;
            //
            // btnSeleccionarPlantilla
            //
            btnSeleccionarPlantilla.Location = new System.Drawing.Point(535, 48);
            btnSeleccionarPlantilla.Name = "btnSeleccionarPlantilla";
            btnSeleccionarPlantilla.Size = new System.Drawing.Size(110, 28);
            btnSeleccionarPlantilla.TabIndex = 3;
            btnSeleccionarPlantilla.Text = "Seleccionar...";
            btnSeleccionarPlantilla.UseVisualStyleBackColor = true;
            btnSeleccionarPlantilla.Click += new System.EventHandler(btnSeleccionarPlantilla_Click);
            //
            // lblCarpetaSalida
            //
            lblCarpetaSalida.AutoSize = true;
            lblCarpetaSalida.Location = new System.Drawing.Point(15, 82);
            lblCarpetaSalida.Name = "lblCarpetaSalida";
            lblCarpetaSalida.Size = new System.Drawing.Size(94, 15);
            lblCarpetaSalida.Text = "Carpeta salida:";
            //
            // txtCarpetaSalida
            //
            txtCarpetaSalida.Location = new System.Drawing.Point(105, 79);
            txtCarpetaSalida.Name = "txtCarpetaSalida";
            txtCarpetaSalida.ReadOnly = true;
            txtCarpetaSalida.Size = new System.Drawing.Size(420, 23);
            txtCarpetaSalida.TabIndex = 4;
            //
            // btnSeleccionarSalida
            //
            btnSeleccionarSalida.Location = new System.Drawing.Point(535, 78);
            btnSeleccionarSalida.Name = "btnSeleccionarSalida";
            btnSeleccionarSalida.Size = new System.Drawing.Size(110, 28);
            btnSeleccionarSalida.TabIndex = 5;
            btnSeleccionarSalida.Text = "Seleccionar...";
            btnSeleccionarSalida.UseVisualStyleBackColor = true;
            btnSeleccionarSalida.Click += new System.EventHandler(btnSeleccionarSalida_Click);

            // ============================
            // grpEjecucion
            // ============================
            grpEjecucion.Controls.Add(lblAse);
            grpEjecucion.Controls.Add(cmbAse);
            grpEjecucion.Controls.Add(chkCincoAse);
            grpEjecucion.Controls.Add(btnEjecutar);
            grpEjecucion.Controls.Add(progressBar);
            grpEjecucion.Location = new System.Drawing.Point(12, 184);
            grpEjecucion.Name = "grpEjecucion";
            grpEjecucion.Size = new System.Drawing.Size(676, 60);
            grpEjecucion.TabIndex = 2;
            grpEjecucion.TabStop = false;
            grpEjecucion.Text = "Ejecución";
            //
            // lblAse
            //
            lblAse.AutoSize = true;
            lblAse.Location = new System.Drawing.Point(15, 22);
            lblAse.Name = "lblAse";
            lblAse.Size = new System.Drawing.Size(30, 15);
            lblAse.Text = "ASE:";
            //
            // cmbAse
            //
            cmbAse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbAse.FormattingEnabled = true;
            cmbAse.Location = new System.Drawing.Point(50, 19);
            cmbAse.Name = "cmbAse";
            cmbAse.Size = new System.Drawing.Size(150, 23);
            cmbAse.TabIndex = 0;
            //
            // chkCincoAse
            //
            chkCincoAse.AutoSize = true;
            chkCincoAse.Location = new System.Drawing.Point(210, 22);
            chkCincoAse.Name = "chkCincoAse";
            chkCincoAse.Size = new System.Drawing.Size(127, 19);
            chkCincoAse.TabIndex = 3;
            chkCincoAse.Text = "Procesar los 5 ASE";
            chkCincoAse.UseVisualStyleBackColor = true;
            chkCincoAse.CheckedChanged += new System.EventHandler(chkCincoAse_CheckedChanged);
            //
            // btnEjecutar
            //
            btnEjecutar.Location = new System.Drawing.Point(350, 18);
            btnEjecutar.Name = "btnEjecutar";
            btnEjecutar.Size = new System.Drawing.Size(130, 28);
            btnEjecutar.TabIndex = 1;
            btnEjecutar.Text = "\u25B6 Ejecutar";
            btnEjecutar.UseVisualStyleBackColor = true;
            btnEjecutar.Click += new System.EventHandler(btnEjecutar_Click);
            //
            // progressBar
            //
            progressBar.Location = new System.Drawing.Point(15, 48);
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(646, 8);
            progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            progressBar.TabIndex = 2;
            progressBar.Visible = false;

            // ============================
            // txtLog
            // ============================
            txtLog.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            txtLog.Location = new System.Drawing.Point(12, 252);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            txtLog.Size = new System.Drawing.Size(676, 302);
            txtLog.TabIndex = 3;

            // ============================
            // statusStrip
            // ============================
            statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                toolStripStatusLabel,
                toolStripVersion
            });
            statusStrip.Location = new System.Drawing.Point(0, 558);
            statusStrip.Name = "statusStrip";
            statusStrip.Size = new System.Drawing.Size(700, 22);
            statusStrip.TabIndex = 4;
            statusStrip.Text = "statusStrip";
            //
            // toolStripStatusLabel
            //
            toolStripStatusLabel.Name = "toolStripStatusLabel";
            toolStripStatusLabel.Size = new System.Drawing.Size(661, 17);
            toolStripStatusLabel.Spring = true;
            toolStripStatusLabel.Text = "Listo";
            toolStripStatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            // toolStripVersion
            //
            toolStripVersion.Name = "toolStripVersion";
            toolStripVersion.Size = new System.Drawing.Size(46, 17);
            toolStripVersion.Text = "v1.0.0";

            // ============================
            // Form1
            // ============================
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            ClientSize = new System.Drawing.Size(700, 580);
            Controls.Add(grpPeriodo);
            Controls.Add(grpArchivos);
            Controls.Add(grpEjecucion);
            Controls.Add(txtLog);
            Controls.Add(statusStrip);
            MinimumSize = new System.Drawing.Size(650, 500);
            Name = "Form1";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Remuneración Quincenal UAESP \u2014 Fase 2";

            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            grpPeriodo.ResumeLayout(false);
            grpPeriodo.PerformLayout();
            grpArchivos.ResumeLayout(false);
            grpArchivos.PerformLayout();
            grpEjecucion.ResumeLayout(false);
            grpEjecucion.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
