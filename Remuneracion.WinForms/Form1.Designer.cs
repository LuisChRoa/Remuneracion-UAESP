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
        private System.Windows.Forms.PictureBox picLogo;
        private System.Windows.Forms.Panel pnlAccent;

        // --- Contenedor principal ---
        private System.Windows.Forms.TableLayoutPanel tlpMain;

        // --- Cards (Panels; G3: renombre grp* → pnlCard*) ---
        private System.Windows.Forms.Panel pnlCardPeriodo;
        private System.Windows.Forms.Panel pnlCardRutas;
        private System.Windows.Forms.Panel pnlCardEjecucion;
        private System.Windows.Forms.Panel pnlAccentBarPeriodo;
        private System.Windows.Forms.Panel pnlAccentBarRutas;
        private System.Windows.Forms.Panel pnlAccentBarEjecucion;
        private System.Windows.Forms.Label lblCardTituloPeriodo;
        private System.Windows.Forms.Label lblCardTituloRutas;
        private System.Windows.Forms.Label lblCardTituloEjecucion;

        // --- Período ---
        private System.Windows.Forms.TableLayoutPanel tlpPeriodo;
        private System.Windows.Forms.Label lblAnio;
        private System.Windows.Forms.ComboBox cmbAnio;
        private System.Windows.Forms.Label lblMes;
        private System.Windows.Forms.ComboBox cmbMes;
        private System.Windows.Forms.Label lblQuincena;
        private System.Windows.Forms.ComboBox cmbQuincena;

        // --- Archivos (Rutas) ---
        private System.Windows.Forms.Label lblCarpeta;
        private System.Windows.Forms.TextBox txtCarpetaFuentes;
        private System.Windows.Forms.Button btnSeleccionarCarpeta;
        private System.Windows.Forms.Label lblPlantilla;
        private System.Windows.Forms.TextBox txtPlantilla;
        private System.Windows.Forms.Button btnSeleccionarPlantilla;
        private System.Windows.Forms.Label lblCarpetaSalida;
        private System.Windows.Forms.TextBox txtCarpetaSalida;
        private System.Windows.Forms.Button btnSeleccionarSalida;

        // --- Ejecución ---
        private System.Windows.Forms.Label lblAse;
        private System.Windows.Forms.ComboBox cmbAse;
        private System.Windows.Forms.CheckBox chkCincoAse;
        private System.Windows.Forms.Button btnEjecutar;
        private System.Windows.Forms.Button btnLimpiar;
        private System.Windows.Forms.Button btnAbrirSalida;
        private System.Windows.Forms.Button btnVerLogs;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblProgresoPct;
        private System.Windows.Forms.Label lblAseActual;

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
            pnlHeader = new Panel();
            lblTitulo = new Label();
            lblSubtitulo = new Label();
            picLogo = new PictureBox();
            pnlAccent = new Panel();
            tlpMain = new TableLayoutPanel();
            pnlCardPeriodo = new Panel();
            pnlAccentBarPeriodo = new Panel();
            lblCardTituloPeriodo = new Label();
            tlpPeriodo = new TableLayoutPanel();
            lblAnio = new Label();
            cmbAnio = new ComboBox();
            lblMes = new Label();
            cmbMes = new ComboBox();
            lblQuincena = new Label();
            cmbQuincena = new ComboBox();
            pnlCardRutas = new Panel();
            pnlAccentBarRutas = new Panel();
            lblCardTituloRutas = new Label();
            lblCarpeta = new Label();
            txtCarpetaFuentes = new TextBox();
            btnSeleccionarCarpeta = new Button();
            lblPlantilla = new Label();
            txtPlantilla = new TextBox();
            btnSeleccionarPlantilla = new Button();
            lblCarpetaSalida = new Label();
            txtCarpetaSalida = new TextBox();
            btnSeleccionarSalida = new Button();
            pnlCardEjecucion = new Panel();
            pnlAccentBarEjecucion = new Panel();
            lblCardTituloEjecucion = new Label();
            lblAse = new Label();
            cmbAse = new ComboBox();
            chkCincoAse = new CheckBox();
            btnEjecutar = new Button();
            progressBar = new ProgressBar();
            btnLimpiar = new Button();
            lblProgresoPct = new Label();
            btnAbrirSalida = new Button();
            lblAseActual = new Label();
            btnVerLogs = new Button();
            statusStrip = new StatusStrip();
            toolStripStatusLabel = new ToolStripStatusLabel();
            toolStripVersion = new ToolStripStatusLabel();
            toolTipRutas = new ToolTip(components);
            folderBrowserDialog = new FolderBrowserDialog();
            openFileDialogPlantilla = new OpenFileDialog();
            label1 = new Label();
            pnlHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)picLogo).BeginInit();
            tlpMain.SuspendLayout();
            pnlCardPeriodo.SuspendLayout();
            tlpPeriodo.SuspendLayout();
            pnlCardRutas.SuspendLayout();
            pnlCardEjecucion.SuspendLayout();
            statusStrip.SuspendLayout();
            SuspendLayout();
            // 
            // pnlHeader
            // 
            pnlHeader.BackColor = SystemColors.Window;
            pnlHeader.Controls.Add(lblTitulo);
            pnlHeader.Controls.Add(lblSubtitulo);
            pnlHeader.Controls.Add(picLogo);
            pnlHeader.Controls.Add(pnlAccent);
            pnlHeader.Dock = DockStyle.Top;
            pnlHeader.Location = new Point(0, 0);
            pnlHeader.Name = "pnlHeader";
            pnlHeader.Size = new Size(960, 56);
            pnlHeader.TabIndex = 0;
            // 
            // lblTitulo
            // 
            lblTitulo.AutoSize = true;
            lblTitulo.Font = new Font("Segoe UI Semibold", 11F);
            lblTitulo.ForeColor = SystemColors.WindowText;
            lblTitulo.Location = new Point(16, 7);
            lblTitulo.Name = "lblTitulo";
            lblTitulo.Size = new Size(288, 25);
            lblTitulo.TabIndex = 0;
            lblTitulo.Text = "Remuneración Quincenal UAESP";
            // 
            // lblSubtitulo
            // 
            lblSubtitulo.AutoSize = true;
            lblSubtitulo.ForeColor = SystemColors.GrayText;
            lblSubtitulo.Location = new Point(16, 31);
            lblSubtitulo.Name = "lblSubtitulo";
            lblSubtitulo.Size = new Size(75, 20);
            lblSubtitulo.TabIndex = 1;
            lblSubtitulo.Text = "Período …";
            // 
            // picLogo
            // 
            picLogo.AccessibleName = "Slogan FNTecnologia";
            picLogo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            picLogo.BackColor = SystemColors.Window;
            picLogo.Location = new Point(770, 4);
            picLogo.Name = "picLogo";
            picLogo.Size = new Size(168, 46);
            picLogo.SizeMode = PictureBoxSizeMode.Zoom;
            picLogo.TabIndex = 4;
            picLogo.TabStop = false;
            // 
            // pnlAccent
            // 
            pnlAccent.BackColor = SystemColors.Highlight;
            pnlAccent.Dock = DockStyle.Bottom;
            pnlAccent.Location = new Point(0, 54);
            pnlAccent.Name = "pnlAccent";
            pnlAccent.Size = new Size(960, 2);
            pnlAccent.TabIndex = 3;
            // 
            // tlpMain
            // 
            tlpMain.AutoSize = true;
            tlpMain.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            tlpMain.ColumnCount = 1;
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpMain.Controls.Add(pnlCardPeriodo, 0, 0);
            tlpMain.Controls.Add(pnlCardRutas, 0, 1);
            tlpMain.Controls.Add(pnlCardEjecucion, 0, 2);
            tlpMain.Dock = DockStyle.Top;
            tlpMain.Location = new Point(0, 56);
            tlpMain.Name = "tlpMain";
            tlpMain.Padding = new Padding(16);
            tlpMain.RowCount = 3;
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.Size = new Size(960, 568);
            tlpMain.TabIndex = 1;
            // 
            // pnlCardPeriodo
            // 
            pnlCardPeriodo.BackColor = SystemColors.Window;
            pnlCardPeriodo.BorderStyle = BorderStyle.FixedSingle;
            pnlCardPeriodo.Controls.Add(pnlAccentBarPeriodo);
            pnlCardPeriodo.Controls.Add(lblCardTituloPeriodo);
            pnlCardPeriodo.Controls.Add(tlpPeriodo);
            pnlCardPeriodo.Dock = DockStyle.Fill;
            pnlCardPeriodo.Location = new Point(22, 22);
            pnlCardPeriodo.Margin = new Padding(6);
            pnlCardPeriodo.MinimumSize = new Size(0, 116);
            pnlCardPeriodo.Name = "pnlCardPeriodo";
            pnlCardPeriodo.Padding = new Padding(16);
            pnlCardPeriodo.Size = new Size(916, 116);
            pnlCardPeriodo.TabIndex = 0;
            // 
            // pnlAccentBarPeriodo
            // 
            pnlAccentBarPeriodo.BackColor = SystemColors.Highlight;
            pnlAccentBarPeriodo.Dock = DockStyle.Left;
            pnlAccentBarPeriodo.Location = new Point(16, 16);
            pnlAccentBarPeriodo.Name = "pnlAccentBarPeriodo";
            pnlAccentBarPeriodo.Size = new Size(4, 82);
            pnlAccentBarPeriodo.TabIndex = 2;
            // 
            // lblCardTituloPeriodo
            // 
            lblCardTituloPeriodo.AutoSize = true;
            lblCardTituloPeriodo.Font = new Font("Segoe UI Semibold", 10F);
            lblCardTituloPeriodo.ForeColor = SystemColors.WindowText;
            lblCardTituloPeriodo.Location = new Point(20, 10);
            lblCardTituloPeriodo.Name = "lblCardTituloPeriodo";
            lblCardTituloPeriodo.Size = new Size(68, 23);
            lblCardTituloPeriodo.TabIndex = 0;
            lblCardTituloPeriodo.Text = "Período";
            // 
            // tlpPeriodo
            // 
            tlpPeriodo.AutoSize = true;
            tlpPeriodo.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            tlpPeriodo.ColumnCount = 3;
            tlpPeriodo.ColumnStyles.Add(new ColumnStyle());
            tlpPeriodo.ColumnStyles.Add(new ColumnStyle());
            tlpPeriodo.ColumnStyles.Add(new ColumnStyle());
            tlpPeriodo.Controls.Add(lblAnio, 0, 0);
            tlpPeriodo.Controls.Add(cmbAnio, 0, 1);
            tlpPeriodo.Controls.Add(lblMes, 1, 0);
            tlpPeriodo.Controls.Add(cmbMes, 1, 1);
            tlpPeriodo.Controls.Add(lblQuincena, 2, 0);
            tlpPeriodo.Controls.Add(cmbQuincena, 2, 1);
            tlpPeriodo.Location = new Point(28, 38);
            tlpPeriodo.Name = "tlpPeriodo";
            tlpPeriodo.RowCount = 3;
            tlpPeriodo.RowStyles.Add(new RowStyle());
            tlpPeriodo.RowStyles.Add(new RowStyle());
            tlpPeriodo.RowStyles.Add(new RowStyle(SizeType.Absolute, 20F));
            tlpPeriodo.Size = new Size(358, 76);
            tlpPeriodo.TabIndex = 1;
            // 
            // lblAnio
            // 
            lblAnio.AutoSize = true;
            lblAnio.Location = new Point(0, 0);
            lblAnio.Margin = new Padding(0, 0, 24, 8);
            lblAnio.Name = "lblAnio";
            lblAnio.Size = new Size(39, 20);
            lblAnio.TabIndex = 0;
            lblAnio.Text = "Año:";
            // 
            // cmbAnio
            // 
            cmbAnio.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbAnio.FormattingEnabled = true;
            cmbAnio.Location = new Point(0, 28);
            cmbAnio.Margin = new Padding(0, 0, 24, 0);
            cmbAnio.Name = "cmbAnio";
            cmbAnio.Size = new Size(80, 28);
            cmbAnio.TabIndex = 1;
            cmbAnio.SelectedIndexChanged += Periodo_SelectedIndexChanged;
            // 
            // lblMes
            // 
            lblMes.AutoSize = true;
            lblMes.Location = new Point(104, 0);
            lblMes.Margin = new Padding(0, 0, 24, 8);
            lblMes.Name = "lblMes";
            lblMes.Size = new Size(39, 20);
            lblMes.TabIndex = 2;
            lblMes.Text = "Mes:";
            // 
            // cmbMes
            // 
            cmbMes.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbMes.FormattingEnabled = true;
            cmbMes.Location = new Point(104, 28);
            cmbMes.Margin = new Padding(0, 0, 24, 0);
            cmbMes.Name = "cmbMes";
            cmbMes.Size = new Size(120, 28);
            cmbMes.TabIndex = 3;
            cmbMes.SelectedIndexChanged += Periodo_SelectedIndexChanged;
            // 
            // lblQuincena
            // 
            lblQuincena.AutoSize = true;
            lblQuincena.Location = new Point(248, 0);
            lblQuincena.Margin = new Padding(0, 0, 0, 8);
            lblQuincena.Name = "lblQuincena";
            lblQuincena.Size = new Size(74, 20);
            lblQuincena.TabIndex = 4;
            lblQuincena.Text = "Quincena:";
            // 
            // cmbQuincena
            // 
            cmbQuincena.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbQuincena.FormattingEnabled = true;
            cmbQuincena.Location = new Point(248, 28);
            cmbQuincena.Margin = new Padding(0);
            cmbQuincena.Name = "cmbQuincena";
            cmbQuincena.Size = new Size(110, 28);
            cmbQuincena.TabIndex = 5;
            cmbQuincena.SelectedIndexChanged += Periodo_SelectedIndexChanged;
            // 
            // pnlCardRutas
            // 
            pnlCardRutas.BackColor = SystemColors.Window;
            pnlCardRutas.BorderStyle = BorderStyle.FixedSingle;
            pnlCardRutas.Controls.Add(pnlAccentBarRutas);
            pnlCardRutas.Controls.Add(lblCardTituloRutas);
            pnlCardRutas.Controls.Add(lblCarpeta);
            pnlCardRutas.Controls.Add(txtCarpetaFuentes);
            pnlCardRutas.Controls.Add(btnSeleccionarCarpeta);
            pnlCardRutas.Controls.Add(lblPlantilla);
            pnlCardRutas.Controls.Add(txtPlantilla);
            pnlCardRutas.Controls.Add(btnSeleccionarPlantilla);
            pnlCardRutas.Controls.Add(lblCarpetaSalida);
            pnlCardRutas.Controls.Add(txtCarpetaSalida);
            pnlCardRutas.Controls.Add(btnSeleccionarSalida);
            pnlCardRutas.Dock = DockStyle.Fill;
            pnlCardRutas.Location = new Point(22, 150);
            pnlCardRutas.Margin = new Padding(6);
            pnlCardRutas.MinimumSize = new Size(0, 188);
            pnlCardRutas.Name = "pnlCardRutas";
            pnlCardRutas.Padding = new Padding(16);
            pnlCardRutas.Size = new Size(916, 188);
            pnlCardRutas.TabIndex = 1;
            // 
            // pnlAccentBarRutas
            // 
            pnlAccentBarRutas.BackColor = SystemColors.Highlight;
            pnlAccentBarRutas.Dock = DockStyle.Left;
            pnlAccentBarRutas.Location = new Point(16, 16);
            pnlAccentBarRutas.Name = "pnlAccentBarRutas";
            pnlAccentBarRutas.Size = new Size(4, 154);
            pnlAccentBarRutas.TabIndex = 10;
            // 
            // lblCardTituloRutas
            // 
            lblCardTituloRutas.AutoSize = true;
            lblCardTituloRutas.Font = new Font("Segoe UI Semibold", 10F);
            lblCardTituloRutas.ForeColor = SystemColors.WindowText;
            lblCardTituloRutas.Location = new Point(20, 10);
            lblCardTituloRutas.Name = "lblCardTituloRutas";
            lblCardTituloRutas.Size = new Size(53, 23);
            lblCardTituloRutas.TabIndex = 0;
            lblCardTituloRutas.Text = "Rutas";
            // 
            // lblCarpeta
            // 
            lblCarpeta.AutoSize = true;
            lblCarpeta.Location = new Point(20, 50);
            lblCarpeta.Name = "lblCarpeta";
            lblCarpeta.Size = new Size(116, 20);
            lblCarpeta.TabIndex = 1;
            lblCarpeta.Text = "Carpeta fuentes:";
            // 
            // txtCarpetaFuentes
            // 
            txtCarpetaFuentes.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtCarpetaFuentes.BackColor = SystemColors.Window;
            txtCarpetaFuentes.Location = new Point(153, 48);
            txtCarpetaFuentes.Name = "txtCarpetaFuentes";
            txtCarpetaFuentes.PlaceholderText = "Sin seleccionar";
            txtCarpetaFuentes.ReadOnly = true;
            txtCarpetaFuentes.Size = new Size(695, 27);
            txtCarpetaFuentes.TabIndex = 2;
            toolTipRutas.SetToolTip(txtCarpetaFuentes, "Sin seleccionar");
            txtCarpetaFuentes.TextChanged += Ruta_TextChanged;
            txtCarpetaFuentes.DoubleClick += Ruta_DoubleClick;
            // 
            // btnSeleccionarCarpeta
            // 
            btnSeleccionarCarpeta.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSeleccionarCarpeta.Location = new Point(856, 46);
            btnSeleccionarCarpeta.Name = "btnSeleccionarCarpeta";
            btnSeleccionarCarpeta.Size = new Size(32, 28);
            btnSeleccionarCarpeta.TabIndex = 3;
            btnSeleccionarCarpeta.Text = "...";
            toolTipRutas.SetToolTip(btnSeleccionarCarpeta, "Examinar…");
            btnSeleccionarCarpeta.UseVisualStyleBackColor = true;
            btnSeleccionarCarpeta.Click += btnSeleccionarCarpeta_Click;
            // 
            // lblPlantilla
            // 
            lblPlantilla.AutoSize = true;
            lblPlantilla.Location = new Point(20, 86);
            lblPlantilla.Name = "lblPlantilla";
            lblPlantilla.Size = new Size(65, 20);
            lblPlantilla.TabIndex = 4;
            lblPlantilla.Text = "Plantilla:";
            // 
            // txtPlantilla
            // 
            txtPlantilla.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtPlantilla.BackColor = SystemColors.Window;
            txtPlantilla.Location = new Point(153, 84);
            txtPlantilla.Name = "txtPlantilla";
            txtPlantilla.PlaceholderText = "Sin seleccionar";
            txtPlantilla.ReadOnly = true;
            txtPlantilla.Size = new Size(695, 27);
            txtPlantilla.TabIndex = 5;
            toolTipRutas.SetToolTip(txtPlantilla, "Sin seleccionar");
            txtPlantilla.TextChanged += Ruta_TextChanged;
            txtPlantilla.DoubleClick += Ruta_DoubleClick;
            // 
            // btnSeleccionarPlantilla
            // 
            btnSeleccionarPlantilla.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSeleccionarPlantilla.Location = new Point(856, 82);
            btnSeleccionarPlantilla.Name = "btnSeleccionarPlantilla";
            btnSeleccionarPlantilla.Size = new Size(32, 28);
            btnSeleccionarPlantilla.TabIndex = 6;
            btnSeleccionarPlantilla.Text = "...";
            toolTipRutas.SetToolTip(btnSeleccionarPlantilla, "Examinar…");
            btnSeleccionarPlantilla.UseVisualStyleBackColor = true;
            btnSeleccionarPlantilla.Click += btnSeleccionarPlantilla_Click;
            // 
            // lblCarpetaSalida
            // 
            lblCarpetaSalida.AutoSize = true;
            lblCarpetaSalida.Location = new Point(20, 122);
            lblCarpetaSalida.Name = "lblCarpetaSalida";
            lblCarpetaSalida.Size = new Size(107, 20);
            lblCarpetaSalida.TabIndex = 7;
            lblCarpetaSalida.Text = "Carpeta salida:";
            // 
            // txtCarpetaSalida
            // 
            txtCarpetaSalida.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            txtCarpetaSalida.BackColor = SystemColors.Window;
            txtCarpetaSalida.Location = new Point(153, 120);
            txtCarpetaSalida.Name = "txtCarpetaSalida";
            txtCarpetaSalida.PlaceholderText = "Sin seleccionar";
            txtCarpetaSalida.ReadOnly = true;
            txtCarpetaSalida.Size = new Size(695, 27);
            txtCarpetaSalida.TabIndex = 8;
            toolTipRutas.SetToolTip(txtCarpetaSalida, "Sin seleccionar");
            txtCarpetaSalida.TextChanged += Ruta_TextChanged;
            txtCarpetaSalida.DoubleClick += Ruta_DoubleClick;
            // 
            // btnSeleccionarSalida
            // 
            btnSeleccionarSalida.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnSeleccionarSalida.Location = new Point(856, 118);
            btnSeleccionarSalida.Name = "btnSeleccionarSalida";
            btnSeleccionarSalida.Size = new Size(32, 28);
            btnSeleccionarSalida.TabIndex = 9;
            btnSeleccionarSalida.Text = "...";
            toolTipRutas.SetToolTip(btnSeleccionarSalida, "Examinar…");
            btnSeleccionarSalida.UseVisualStyleBackColor = true;
            btnSeleccionarSalida.Click += btnSeleccionarSalida_Click;
            // 
            // pnlCardEjecucion
            // 
            pnlCardEjecucion.BackColor = SystemColors.Window;
            pnlCardEjecucion.BorderStyle = BorderStyle.FixedSingle;
            pnlCardEjecucion.Controls.Add(pnlAccentBarEjecucion);
            pnlCardEjecucion.Controls.Add(lblCardTituloEjecucion);
            pnlCardEjecucion.Controls.Add(lblAse);
            pnlCardEjecucion.Controls.Add(cmbAse);
            pnlCardEjecucion.Controls.Add(chkCincoAse);
            pnlCardEjecucion.Controls.Add(btnEjecutar);
            pnlCardEjecucion.Controls.Add(progressBar);
            pnlCardEjecucion.Controls.Add(btnLimpiar);
            pnlCardEjecucion.Controls.Add(lblProgresoPct);
            pnlCardEjecucion.Controls.Add(btnAbrirSalida);
            pnlCardEjecucion.Controls.Add(lblAseActual);
            pnlCardEjecucion.Controls.Add(btnVerLogs);
            pnlCardEjecucion.Dock = DockStyle.Fill;
            pnlCardEjecucion.Location = new Point(22, 350);
            pnlCardEjecucion.Margin = new Padding(6);
            pnlCardEjecucion.MinimumSize = new Size(0, 196);
            pnlCardEjecucion.Name = "pnlCardEjecucion";
            pnlCardEjecucion.Padding = new Padding(16);
            pnlCardEjecucion.Size = new Size(916, 196);
            pnlCardEjecucion.TabIndex = 2;
            // 
            // pnlAccentBarEjecucion
            // 
            pnlAccentBarEjecucion.BackColor = SystemColors.Highlight;
            pnlAccentBarEjecucion.Dock = DockStyle.Left;
            pnlAccentBarEjecucion.Location = new Point(16, 16);
            pnlAccentBarEjecucion.Name = "pnlAccentBarEjecucion";
            pnlAccentBarEjecucion.Size = new Size(4, 162);
            pnlAccentBarEjecucion.TabIndex = 11;
            // 
            // lblCardTituloEjecucion
            // 
            lblCardTituloEjecucion.AutoSize = true;
            lblCardTituloEjecucion.Font = new Font("Segoe UI Semibold", 10F);
            lblCardTituloEjecucion.ForeColor = SystemColors.WindowText;
            lblCardTituloEjecucion.Location = new Point(20, 10);
            lblCardTituloEjecucion.Name = "lblCardTituloEjecucion";
            lblCardTituloEjecucion.Size = new Size(82, 23);
            lblCardTituloEjecucion.TabIndex = 0;
            lblCardTituloEjecucion.Text = "Ejecución";
            // 
            // lblAse
            // 
            lblAse.AutoSize = true;
            lblAse.Location = new Point(20, 46);
            lblAse.Name = "lblAse";
            lblAse.Size = new Size(38, 20);
            lblAse.TabIndex = 1;
            lblAse.Text = "ASE:";
            // 
            // cmbAse
            // 
            cmbAse.DropDownStyle = ComboBoxStyle.DropDownList;
            cmbAse.FormattingEnabled = true;
            cmbAse.Location = new Point(58, 42);
            cmbAse.Name = "cmbAse";
            cmbAse.Size = new Size(150, 28);
            cmbAse.TabIndex = 2;
            // 
            // chkCincoAse
            // 
            chkCincoAse.AutoSize = true;
            chkCincoAse.Location = new Point(220, 45);
            chkCincoAse.Name = "chkCincoAse";
            chkCincoAse.Size = new Size(129, 24);
            chkCincoAse.TabIndex = 3;
            chkCincoAse.Text = "Procesar todos";
            chkCincoAse.UseVisualStyleBackColor = true;
            chkCincoAse.CheckedChanged += chkCincoAse_CheckedChanged;
            // 
            // btnEjecutar
            // 
            btnEjecutar.BackColor = SystemColors.Highlight;
            btnEjecutar.FlatAppearance.BorderSize = 0;
            btnEjecutar.FlatStyle = FlatStyle.Flat;
            btnEjecutar.ForeColor = SystemColors.HighlightText;
            btnEjecutar.Location = new Point(761, 139);
            btnEjecutar.Name = "btnEjecutar";
            btnEjecutar.Size = new Size(140, 28);
            btnEjecutar.TabIndex = 4;
            btnEjecutar.Text = "▶ Ejecutar";
            btnEjecutar.UseVisualStyleBackColor = false;
            btnEjecutar.Click += btnEjecutar_Click;
            // 
            // progressBar
            // 
            progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Location = new Point(26, 91);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(822, 14);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.TabIndex = 8;
            progressBar.TabStop = false;
            // 
            // btnLimpiar
            // 
            btnLimpiar.BackColor = SystemColors.Window;
            btnLimpiar.FlatAppearance.BorderColor = SystemColors.ControlDark;
            btnLimpiar.FlatStyle = FlatStyle.Flat;
            btnLimpiar.ForeColor = SystemColors.WindowText;
            btnLimpiar.Location = new Point(28, 139);
            btnLimpiar.Name = "btnLimpiar";
            btnLimpiar.Size = new Size(110, 28);
            btnLimpiar.TabIndex = 5;
            btnLimpiar.Text = "Limpiar";
            btnLimpiar.UseVisualStyleBackColor = false;
            btnLimpiar.Click += btnLimpiar_Click;
            // 
            // lblProgresoPct
            // 
            lblProgresoPct.Location = new Point(857, 81);
            lblProgresoPct.Name = "lblProgresoPct";
            lblProgresoPct.Size = new Size(44, 32);
            lblProgresoPct.TabIndex = 9;
            lblProgresoPct.Text = "0 %";
            lblProgresoPct.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // btnAbrirSalida
            // 
            btnAbrirSalida.BackColor = SystemColors.Window;
            btnAbrirSalida.FlatAppearance.BorderColor = SystemColors.ControlDark;
            btnAbrirSalida.FlatStyle = FlatStyle.Flat;
            btnAbrirSalida.ForeColor = SystemColors.WindowText;
            btnAbrirSalida.Location = new Point(633, 139);
            btnAbrirSalida.Name = "btnAbrirSalida";
            btnAbrirSalida.Size = new Size(120, 28);
            btnAbrirSalida.TabIndex = 6;
            btnAbrirSalida.Text = "Abrir salida";
            btnAbrirSalida.UseVisualStyleBackColor = false;
            btnAbrirSalida.Click += btnAbrirSalida_Click;
            // 
            // lblAseActual
            // 
            lblAseActual.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            lblAseActual.AutoEllipsis = true;
            lblAseActual.Location = new Point(61, 108);
            lblAseActual.Name = "lblAseActual";
            lblAseActual.Size = new Size(834, 24);
            lblAseActual.TabIndex = 10;
            lblAseActual.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // btnVerLogs
            // 
            btnVerLogs.BackColor = SystemColors.Window;
            btnVerLogs.Enabled = false;
            btnVerLogs.FlatAppearance.BorderColor = SystemColors.ControlDark;
            btnVerLogs.FlatStyle = FlatStyle.Flat;
            btnVerLogs.ForeColor = SystemColors.WindowText;
            btnVerLogs.Location = new Point(144, 139);
            btnVerLogs.Name = "btnVerLogs";
            btnVerLogs.Size = new Size(120, 28);
            btnVerLogs.TabIndex = 7;
            btnVerLogs.Text = "Ver logs";
            btnVerLogs.UseVisualStyleBackColor = false;
            btnVerLogs.Click += btnVerLogs_Click;
            // 
            // statusStrip
            // 
            statusStrip.ImageScalingSize = new Size(20, 20);
            statusStrip.Items.AddRange(new ToolStripItem[] { toolStripStatusLabel, toolStripVersion });
            statusStrip.Location = new Point(0, 654);
            statusStrip.Name = "statusStrip";
            statusStrip.Size = new Size(960, 26);
            statusStrip.TabIndex = 2;
            statusStrip.Text = "statusStrip";
            // 
            // toolStripStatusLabel
            // 
            toolStripStatusLabel.Name = "toolStripStatusLabel";
            toolStripStatusLabel.Size = new Size(899, 20);
            toolStripStatusLabel.Spring = true;
            toolStripStatusLabel.Text = "Listo";
            toolStripStatusLabel.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // toolStripVersion
            // 
            toolStripVersion.Name = "toolStripVersion";
            toolStripVersion.Size = new Size(46, 20);
            toolStripVersion.Text = "v1.0.0";
            // 
            // toolTipRutas
            // 
            toolTipRutas.AutoPopDelay = 10000;
            toolTipRutas.InitialDelay = 400;
            toolTipRutas.ReshowDelay = 100;
            toolTipRutas.ShowAlways = true;
            // 
            // openFileDialogPlantilla
            // 
            openFileDialogPlantilla.Filter = "Archivos Excel (*.xlsx)|*.xlsx";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(892, 627);
            label1.Margin = new Padding(0, 0, 24, 8);
            label1.Name = "label1";
            label1.Size = new Size(46, 20);
            label1.TabIndex = 6;
            label1.Text = "v1.0.0";
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            ClientSize = new Size(960, 680);
            Controls.Add(label1);
            Controls.Add(statusStrip);
            Controls.Add(tlpMain);
            Controls.Add(pnlHeader);
            MinimumSize = new Size(940, 660);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Remuneración Quincenal UAESP";
            pnlHeader.ResumeLayout(false);
            pnlHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)picLogo).EndInit();
            tlpMain.ResumeLayout(false);
            pnlCardPeriodo.ResumeLayout(false);
            pnlCardPeriodo.PerformLayout();
            tlpPeriodo.ResumeLayout(false);
            tlpPeriodo.PerformLayout();
            pnlCardRutas.ResumeLayout(false);
            pnlCardRutas.PerformLayout();
            pnlCardEjecucion.ResumeLayout(false);
            pnlCardEjecucion.PerformLayout();
            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
    }
}
