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

        // --- Header ---
        private System.Windows.Forms.Panel pnlHeader;
        private System.Windows.Forms.Label lblTitulo;
        private System.Windows.Forms.Label lblSubtitulo;
        private System.Windows.Forms.Label lblVersionBadge;

        // --- Contenedor principal ---
        private System.Windows.Forms.TableLayoutPanel tlpMain;

        // --- GroupBoxes ---
        private System.Windows.Forms.GroupBox grpPeriodo;
        private System.Windows.Forms.GroupBox grpArchivos;
        private System.Windows.Forms.GroupBox grpEjecucion;
        private System.Windows.Forms.GroupBox grpLog;

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
        private System.Windows.Forms.Button btnLimpiar;
        private System.Windows.Forms.Button btnAbrirSalida;
        private System.Windows.Forms.Button btnCopiarLog;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblProgresoPct;
        private System.Windows.Forms.Label lblAseActual;

        // --- Log ---
        private System.Windows.Forms.Label lblLineasLog;
        private System.Windows.Forms.TextBox txtLog;

        // --- StatusStrip ---
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ToolStripStatusLabel toolStripVersion;

        // --- ToolTip ---
        private System.Windows.Forms.ToolTip toolTipRutas;

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

            // === Header ===
            pnlHeader = new System.Windows.Forms.Panel();
            lblTitulo = new System.Windows.Forms.Label();
            lblSubtitulo = new System.Windows.Forms.Label();
            lblVersionBadge = new System.Windows.Forms.Label();

            // === Contenedor principal ===
            tlpMain = new System.Windows.Forms.TableLayoutPanel();

            // === GroupBoxes ===
            grpPeriodo = new System.Windows.Forms.GroupBox();
            grpArchivos = new System.Windows.Forms.GroupBox();
            grpEjecucion = new System.Windows.Forms.GroupBox();
            grpLog = new System.Windows.Forms.GroupBox();

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
            btnLimpiar = new System.Windows.Forms.Button();
            btnAbrirSalida = new System.Windows.Forms.Button();
            btnCopiarLog = new System.Windows.Forms.Button();
            progressBar = new System.Windows.Forms.ProgressBar();
            lblProgresoPct = new System.Windows.Forms.Label();
            lblAseActual = new System.Windows.Forms.Label();

            // === Log ===
            lblLineasLog = new System.Windows.Forms.Label();
            txtLog = new System.Windows.Forms.TextBox();

            // === StatusStrip ===
            statusStrip = new System.Windows.Forms.StatusStrip();
            toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            toolStripVersion = new System.Windows.Forms.ToolStripStatusLabel();

            // === ToolTip ===
            toolTipRutas = new System.Windows.Forms.ToolTip(components);

            // === Diálogos ===
            folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            openFileDialogPlantilla = new System.Windows.Forms.OpenFileDialog();
            openFileDialogPlantilla.Filter = "Archivos Excel (*.xlsx)|*.xlsx";

            pnlHeader.SuspendLayout();
            tlpMain.SuspendLayout();
            grpPeriodo.SuspendLayout();
            grpArchivos.SuspendLayout();
            grpEjecucion.SuspendLayout();
            grpLog.SuspendLayout();
            statusStrip.SuspendLayout();
            SuspendLayout();

            // ============================
            // pnlHeader (Dock=Top, H=64, Highlight)
            // ============================
            pnlHeader.BackColor = System.Drawing.SystemColors.Highlight;
            pnlHeader.Controls.Add(lblTitulo);
            pnlHeader.Controls.Add(lblSubtitulo);
            pnlHeader.Controls.Add(lblVersionBadge);
            pnlHeader.Dock = System.Windows.Forms.DockStyle.Top;
            pnlHeader.Location = new System.Drawing.Point(0, 0);
            pnlHeader.Name = "pnlHeader";
            pnlHeader.Size = new System.Drawing.Size(880, 64);
            pnlHeader.TabIndex = 0;
            //
            // lblTitulo
            //
            lblTitulo.AutoSize = true;
            lblTitulo.BackColor = System.Drawing.SystemColors.Highlight;
            lblTitulo.Font = new System.Drawing.Font("Segoe UI", 14F, System.Drawing.FontStyle.Bold);
            lblTitulo.ForeColor = System.Drawing.SystemColors.HighlightText;
            lblTitulo.Location = new System.Drawing.Point(16, 9);
            lblTitulo.Name = "lblTitulo";
            lblTitulo.Size = new System.Drawing.Size(0, 25);
            lblTitulo.TabIndex = 0;
            lblTitulo.Text = "Remuneración Quincenal UAESP";
            //
            // lblSubtitulo
            //
            lblSubtitulo.AutoSize = true;
            lblSubtitulo.BackColor = System.Drawing.SystemColors.Highlight;
            lblSubtitulo.ForeColor = System.Drawing.SystemColors.HighlightText;
            lblSubtitulo.Location = new System.Drawing.Point(16, 38);
            lblSubtitulo.Name = "lblSubtitulo";
            lblSubtitulo.Size = new System.Drawing.Size(0, 15);
            lblSubtitulo.TabIndex = 1;
            lblSubtitulo.Text = "Período …";
            //
            // lblVersionBadge
            //
            lblVersionBadge.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            lblVersionBadge.AutoSize = true;
            lblVersionBadge.BackColor = System.Drawing.SystemColors.Highlight;
            lblVersionBadge.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            lblVersionBadge.ForeColor = System.Drawing.SystemColors.HighlightText;
            lblVersionBadge.Location = new System.Drawing.Point(802, 20);
            lblVersionBadge.Name = "lblVersionBadge";
            lblVersionBadge.Padding = new System.Windows.Forms.Padding(4);
            lblVersionBadge.Size = new System.Drawing.Size(0, 19);
            lblVersionBadge.TabIndex = 2;
            lblVersionBadge.Text = "v1.0.0";

            // ============================
            // tlpMain (Dock=Fill, 1 col × 4 filas, Padding=8)
            // ============================
            tlpMain.ColumnCount = 1;
            tlpMain.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            tlpMain.Controls.Add(grpPeriodo, 0, 0);
            tlpMain.Controls.Add(grpArchivos, 0, 1);
            tlpMain.Controls.Add(grpEjecucion, 0, 2);
            tlpMain.Controls.Add(grpLog, 0, 3);
            tlpMain.Dock = System.Windows.Forms.DockStyle.Fill;
            tlpMain.Location = new System.Drawing.Point(0, 64);
            tlpMain.Name = "tlpMain";
            tlpMain.Padding = new System.Windows.Forms.Padding(8);
            tlpMain.RowCount = 4;
            tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize));
            tlpMain.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            tlpMain.Size = new System.Drawing.Size(880, 554);
            tlpMain.TabIndex = 1;

            // ============================
            // grpPeriodo (Dock=Fill, Margin=4, Padding=8)
            // ============================
            grpPeriodo.Controls.Add(lblAnio);
            grpPeriodo.Controls.Add(cmbAnio);
            grpPeriodo.Controls.Add(lblMes);
            grpPeriodo.Controls.Add(cmbMes);
            grpPeriodo.Controls.Add(lblQuincena);
            grpPeriodo.Controls.Add(cmbQuincena);
            grpPeriodo.Dock = System.Windows.Forms.DockStyle.Fill;
            grpPeriodo.Margin = new System.Windows.Forms.Padding(4);
            grpPeriodo.MinimumSize = new System.Drawing.Size(0, 64);
            grpPeriodo.Name = "grpPeriodo";
            grpPeriodo.Padding = new System.Windows.Forms.Padding(8);
            grpPeriodo.Size = new System.Drawing.Size(856, 64);
            grpPeriodo.TabIndex = 0;
            grpPeriodo.TabStop = false;
            grpPeriodo.Text = "Período";
            //
            // lblAnio
            //
            lblAnio.AutoSize = true;
            lblAnio.Location = new System.Drawing.Point(16, 27);
            lblAnio.Name = "lblAnio";
            lblAnio.Size = new System.Drawing.Size(33, 15);
            lblAnio.Text = "Año:";
            //
            // cmbAnio
            //
            cmbAnio.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbAnio.FormattingEnabled = true;
            cmbAnio.Location = new System.Drawing.Point(57, 24);
            cmbAnio.Name = "cmbAnio";
            cmbAnio.Size = new System.Drawing.Size(80, 23);
            cmbAnio.TabIndex = 1;
            cmbAnio.SelectedIndexChanged += new System.EventHandler(Periodo_SelectedIndexChanged);
            //
            // lblMes
            //
            lblMes.AutoSize = true;
            lblMes.Location = new System.Drawing.Point(152, 27);
            lblMes.Name = "lblMes";
            lblMes.Size = new System.Drawing.Size(35, 15);
            lblMes.Text = "Mes:";
            //
            // cmbMes
            //
            cmbMes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbMes.FormattingEnabled = true;
            cmbMes.Location = new System.Drawing.Point(192, 24);
            cmbMes.Name = "cmbMes";
            cmbMes.Size = new System.Drawing.Size(120, 23);
            cmbMes.TabIndex = 2;
            cmbMes.SelectedIndexChanged += new System.EventHandler(Periodo_SelectedIndexChanged);
            //
            // lblQuincena
            //
            lblQuincena.AutoSize = true;
            lblQuincena.Location = new System.Drawing.Point(327, 27);
            lblQuincena.Name = "lblQuincena";
            lblQuincena.Size = new System.Drawing.Size(64, 15);
            lblQuincena.Text = "Quincena:";
            //
            // cmbQuincena
            //
            cmbQuincena.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbQuincena.FormattingEnabled = true;
            cmbQuincena.Location = new System.Drawing.Point(397, 24);
            cmbQuincena.Name = "cmbQuincena";
            cmbQuincena.Size = new System.Drawing.Size(110, 23);
            cmbQuincena.TabIndex = 3;
            cmbQuincena.SelectedIndexChanged += new System.EventHandler(Periodo_SelectedIndexChanged);

            // ============================
            // grpArchivos (Dock=Fill, Text="Rutas", Margin=4, Padding=8)
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
            grpArchivos.Dock = System.Windows.Forms.DockStyle.Fill;
            grpArchivos.Margin = new System.Windows.Forms.Padding(4);
            grpArchivos.MinimumSize = new System.Drawing.Size(0, 132);
            grpArchivos.Name = "grpArchivos";
            grpArchivos.Padding = new System.Windows.Forms.Padding(8);
            grpArchivos.Size = new System.Drawing.Size(856, 132);
            grpArchivos.TabIndex = 1;
            grpArchivos.TabStop = false;
            grpArchivos.Text = "Rutas";
            //
            // lblCarpeta
            //
            lblCarpeta.AutoSize = true;
            lblCarpeta.Location = new System.Drawing.Point(16, 27);
            lblCarpeta.Name = "lblCarpeta";
            lblCarpeta.Size = new System.Drawing.Size(88, 15);
            lblCarpeta.Text = "Carpeta fuentes:";
            //
            // txtCarpetaFuentes
            //
            txtCarpetaFuentes.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            txtCarpetaFuentes.Location = new System.Drawing.Point(105, 24);
            txtCarpetaFuentes.Name = "txtCarpetaFuentes";
            txtCarpetaFuentes.ReadOnly = true;
            txtCarpetaFuentes.Size = new System.Drawing.Size(620, 23);
            txtCarpetaFuentes.TabIndex = 0;
            txtCarpetaFuentes.TextChanged += new System.EventHandler(Ruta_TextChanged);
            txtCarpetaFuentes.DoubleClick += new System.EventHandler(Ruta_DoubleClick);
            //
            // btnSeleccionarCarpeta
            //
            btnSeleccionarCarpeta.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            btnSeleccionarCarpeta.Location = new System.Drawing.Point(736, 22);
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
            lblPlantilla.Location = new System.Drawing.Point(16, 59);
            lblPlantilla.Name = "lblPlantilla";
            lblPlantilla.Size = new System.Drawing.Size(57, 15);
            lblPlantilla.Text = "Plantilla:";
            //
            // txtPlantilla
            //
            txtPlantilla.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            txtPlantilla.Location = new System.Drawing.Point(105, 56);
            txtPlantilla.Name = "txtPlantilla";
            txtPlantilla.ReadOnly = true;
            txtPlantilla.Size = new System.Drawing.Size(620, 23);
            txtPlantilla.TabIndex = 2;
            txtPlantilla.TextChanged += new System.EventHandler(Ruta_TextChanged);
            txtPlantilla.DoubleClick += new System.EventHandler(Ruta_DoubleClick);
            //
            // btnSeleccionarPlantilla
            //
            btnSeleccionarPlantilla.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            btnSeleccionarPlantilla.Location = new System.Drawing.Point(736, 54);
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
            lblCarpetaSalida.Location = new System.Drawing.Point(16, 91);
            lblCarpetaSalida.Name = "lblCarpetaSalida";
            lblCarpetaSalida.Size = new System.Drawing.Size(94, 15);
            lblCarpetaSalida.Text = "Carpeta salida:";
            //
            // txtCarpetaSalida
            //
            txtCarpetaSalida.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            txtCarpetaSalida.Location = new System.Drawing.Point(105, 88);
            txtCarpetaSalida.Name = "txtCarpetaSalida";
            txtCarpetaSalida.ReadOnly = true;
            txtCarpetaSalida.Size = new System.Drawing.Size(620, 23);
            txtCarpetaSalida.TabIndex = 4;
            txtCarpetaSalida.TextChanged += new System.EventHandler(Ruta_TextChanged);
            txtCarpetaSalida.DoubleClick += new System.EventHandler(Ruta_DoubleClick);
            //
            // btnSeleccionarSalida
            //
            btnSeleccionarSalida.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            btnSeleccionarSalida.Location = new System.Drawing.Point(736, 86);
            btnSeleccionarSalida.Name = "btnSeleccionarSalida";
            btnSeleccionarSalida.Size = new System.Drawing.Size(110, 28);
            btnSeleccionarSalida.TabIndex = 5;
            btnSeleccionarSalida.Text = "Seleccionar...";
            btnSeleccionarSalida.UseVisualStyleBackColor = true;
            btnSeleccionarSalida.Click += new System.EventHandler(btnSeleccionarSalida_Click);

            // ============================
            // grpEjecucion (Dock=Fill, Margin=4, Padding=8)
            // ============================
            grpEjecucion.Controls.Add(lblAse);
            grpEjecucion.Controls.Add(cmbAse);
            grpEjecucion.Controls.Add(chkCincoAse);
            grpEjecucion.Controls.Add(btnEjecutar);
            grpEjecucion.Controls.Add(btnLimpiar);
            grpEjecucion.Controls.Add(btnAbrirSalida);
            grpEjecucion.Controls.Add(btnCopiarLog);
            grpEjecucion.Controls.Add(progressBar);
            grpEjecucion.Controls.Add(lblProgresoPct);
            grpEjecucion.Controls.Add(lblAseActual);
            grpEjecucion.Dock = System.Windows.Forms.DockStyle.Fill;
            grpEjecucion.Margin = new System.Windows.Forms.Padding(4);
            grpEjecucion.MinimumSize = new System.Drawing.Size(0, 126);
            grpEjecucion.Name = "grpEjecucion";
            grpEjecucion.Padding = new System.Windows.Forms.Padding(8);
            grpEjecucion.Size = new System.Drawing.Size(856, 126);
            grpEjecucion.TabIndex = 2;
            grpEjecucion.TabStop = false;
            grpEjecucion.Text = "Ejecución";
            //
            // lblAse
            //
            lblAse.AutoSize = true;
            lblAse.Location = new System.Drawing.Point(16, 27);
            lblAse.Name = "lblAse";
            lblAse.Size = new System.Drawing.Size(30, 15);
            lblAse.Text = "ASE:";
            //
            // cmbAse
            //
            cmbAse.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            cmbAse.FormattingEnabled = true;
            cmbAse.Location = new System.Drawing.Point(51, 24);
            cmbAse.Name = "cmbAse";
            cmbAse.Size = new System.Drawing.Size(150, 23);
            cmbAse.TabIndex = 0;
            //
            // chkCincoAse
            //
            chkCincoAse.AutoSize = true;
            chkCincoAse.Location = new System.Drawing.Point(211, 27);
            chkCincoAse.Name = "chkCincoAse";
            chkCincoAse.Size = new System.Drawing.Size(127, 19);
            chkCincoAse.TabIndex = 5;
            chkCincoAse.Text = "Procesar los 5 ASE";
            chkCincoAse.UseVisualStyleBackColor = true;
            chkCincoAse.CheckedChanged += new System.EventHandler(chkCincoAse_CheckedChanged);
            //
            // btnEjecutar
            //
            btnEjecutar.Location = new System.Drawing.Point(16, 56);
            btnEjecutar.Name = "btnEjecutar";
            btnEjecutar.Size = new System.Drawing.Size(130, 28);
            btnEjecutar.TabIndex = 1;
            btnEjecutar.Text = "\u25B6 Ejecutar";
            btnEjecutar.UseVisualStyleBackColor = true;
            btnEjecutar.Click += new System.EventHandler(btnEjecutar_Click);
            //
            // btnLimpiar
            //
            btnLimpiar.Location = new System.Drawing.Point(154, 56);
            btnLimpiar.Name = "btnLimpiar";
            btnLimpiar.Size = new System.Drawing.Size(110, 28);
            btnLimpiar.TabIndex = 2;
            btnLimpiar.Text = "Limpiar";
            btnLimpiar.UseVisualStyleBackColor = true;
            btnLimpiar.Click += new System.EventHandler(btnLimpiar_Click);
            //
            // btnAbrirSalida
            //
            btnAbrirSalida.Location = new System.Drawing.Point(272, 56);
            btnAbrirSalida.Name = "btnAbrirSalida";
            btnAbrirSalida.Size = new System.Drawing.Size(120, 28);
            btnAbrirSalida.TabIndex = 3;
            btnAbrirSalida.Text = "Abrir salida";
            btnAbrirSalida.UseVisualStyleBackColor = true;
            btnAbrirSalida.Click += new System.EventHandler(btnAbrirSalida_Click);
            //
            // btnCopiarLog
            //
            btnCopiarLog.Enabled = false;
            btnCopiarLog.Location = new System.Drawing.Point(400, 56);
            btnCopiarLog.Name = "btnCopiarLog";
            btnCopiarLog.Size = new System.Drawing.Size(120, 28);
            btnCopiarLog.TabIndex = 4;
            btnCopiarLog.Text = "Copiar log";
            btnCopiarLog.UseVisualStyleBackColor = true;
            btnCopiarLog.Click += new System.EventHandler(btnCopiarLog_Click);
            //
            // progressBar
            //
            progressBar.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            progressBar.Location = new System.Drawing.Point(60, 94);
            progressBar.Maximum = 100;
            progressBar.Minimum = 0;
            progressBar.Name = "progressBar";
            progressBar.Size = new System.Drawing.Size(520, 18);
            progressBar.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            progressBar.TabIndex = 6;
            progressBar.Value = 0;
            progressBar.Visible = true;
            //
            // lblProgresoPct
            //
            lblProgresoPct.AutoSize = true;
            lblProgresoPct.Location = new System.Drawing.Point(16, 95);
            lblProgresoPct.Name = "lblProgresoPct";
            lblProgresoPct.Size = new System.Drawing.Size(31, 15);
            lblProgresoPct.TabIndex = 7;
            lblProgresoPct.Text = "0 %";
            //
            // lblAseActual
            //
            lblAseActual.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            lblAseActual.AutoEllipsis = true;
            lblAseActual.Location = new System.Drawing.Point(592, 95);
            lblAseActual.Name = "lblAseActual";
            lblAseActual.Size = new System.Drawing.Size(256, 15);
            lblAseActual.TabIndex = 8;
            lblAseActual.Text = "—";
            lblAseActual.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;

            // ============================
            // grpLog (Dock=Fill, Margin=4, header interno + txtLog Fill)
            // ============================
            grpLog.Controls.Add(txtLog);
            grpLog.Controls.Add(lblLineasLog);
            grpLog.Dock = System.Windows.Forms.DockStyle.Fill;
            grpLog.Margin = new System.Windows.Forms.Padding(4);
            grpLog.Name = "grpLog";
            grpLog.Padding = new System.Windows.Forms.Padding(8, 24, 8, 8);
            grpLog.Size = new System.Drawing.Size(856, 160);
            grpLog.TabIndex = 3;
            grpLog.TabStop = false;
            grpLog.Text = "Resultado · Log";
            //
            // lblLineasLog
            //
            lblLineasLog.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            lblLineasLog.AutoSize = true;
            lblLineasLog.Location = new System.Drawing.Point(789, 6);
            lblLineasLog.Name = "lblLineasLog";
            lblLineasLog.Size = new System.Drawing.Size(51, 15);
            lblLineasLog.TabIndex = 1;
            lblLineasLog.Text = "0 líneas";
            //
            // txtLog
            //
            txtLog.Dock = System.Windows.Forms.DockStyle.Fill;
            txtLog.Font = new System.Drawing.Font("Consolas", 9F);
            txtLog.Location = new System.Drawing.Point(8, 24);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            txtLog.Size = new System.Drawing.Size(840, 128);
            txtLog.TabIndex = 0;
            txtLog.TabStop = false;
            txtLog.WordWrap = false;

            // ============================
            // toolTipRutas
            // ============================
            toolTipRutas.AutoPopDelay = 10000;
            toolTipRutas.InitialDelay = 400;
            toolTipRutas.ReshowDelay = 100;
            toolTipRutas.ShowAlways = true;
            toolTipRutas.SetToolTip(txtCarpetaFuentes, "Sin seleccionar");
            toolTipRutas.SetToolTip(txtPlantilla, "Sin seleccionar");
            toolTipRutas.SetToolTip(txtCarpetaSalida, "Sin seleccionar");

            // ============================
            // statusStrip
            // ============================
            statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
                toolStripStatusLabel,
                toolStripVersion
            });
            statusStrip.Location = new System.Drawing.Point(0, 618);
            statusStrip.Name = "statusStrip";
            statusStrip.Size = new System.Drawing.Size(880, 22);
            statusStrip.TabIndex = 2;
            statusStrip.Text = "statusStrip";
            //
            // toolStripStatusLabel
            //
            toolStripStatusLabel.Name = "toolStripStatusLabel";
            toolStripStatusLabel.Size = new System.Drawing.Size(841, 17);
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
            ClientSize = new System.Drawing.Size(880, 640);
            Controls.Add(tlpMain);
            Controls.Add(pnlHeader);
            Controls.Add(statusStrip);
            MinimumSize = new System.Drawing.Size(860, 600);
            Name = "Form1";
            StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            Text = "Remuneración Quincenal UAESP \u2014 Fase 2";

            pnlHeader.ResumeLayout(false);
            pnlHeader.PerformLayout();
            tlpMain.ResumeLayout(false);
            tlpMain.PerformLayout();
            grpPeriodo.ResumeLayout(false);
            grpPeriodo.PerformLayout();
            grpArchivos.ResumeLayout(false);
            grpArchivos.PerformLayout();
            grpEjecucion.ResumeLayout(false);
            grpEjecucion.PerformLayout();
            grpLog.ResumeLayout(false);
            grpLog.PerformLayout();
            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
