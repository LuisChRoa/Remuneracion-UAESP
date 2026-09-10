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
        private System.Windows.Forms.PictureBox picLogo;
        private System.Windows.Forms.Panel pnlAccent;

        // --- Contenedor principal ---
        private System.Windows.Forms.TableLayoutPanel tlpMain;

        // --- Cards (Panels; G3: renombre grp* → pnlCard*) ---
        private System.Windows.Forms.Panel pnlCardPeriodo;
        private System.Windows.Forms.Panel pnlCardRutas;
        private System.Windows.Forms.Panel pnlCardEjecucion;
        private System.Windows.Forms.Panel pnlCardResultado;
        private System.Windows.Forms.Panel pnlAccentBarPeriodo;
        private System.Windows.Forms.Panel pnlAccentBarRutas;
        private System.Windows.Forms.Panel pnlAccentBarEjecucion;
        private System.Windows.Forms.Panel pnlAccentBarResultado;
        private System.Windows.Forms.Label lblCardTituloPeriodo;
        private System.Windows.Forms.Label lblCardTituloRutas;
        private System.Windows.Forms.Label lblCardTituloEjecucion;
        private System.Windows.Forms.Label lblCardTituloResultado;

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
        private System.Windows.Forms.Button btnCopiarLog;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label lblProgresoPct;
        private System.Windows.Forms.Label lblAseActual;

        // --- Resultado (colapsable) ---
        private System.Windows.Forms.TableLayoutPanel tlpResultadoResumen;
        private System.Windows.Forms.Label lblEstadoHumano;
        private System.Windows.Forms.Label lblResumenUnaLinea;
        private System.Windows.Forms.CheckBox chkVerDetalle;
        private System.Windows.Forms.Panel pnlDetalleTecnico;
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
            pnlHeader = new Panel();
            lblTitulo = new Label();
            lblSubtitulo = new Label();
            lblVersionBadge = new Label();
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
            btnLimpiar = new Button();
            btnAbrirSalida = new Button();
            btnCopiarLog = new Button();
            progressBar = new ProgressBar();
            lblProgresoPct = new Label();
            lblAseActual = new Label();
            pnlCardResultado = new Panel();
            pnlDetalleTecnico = new Panel();
            lblLineasLog = new Label();
            txtLog = new TextBox();
            tlpResultadoResumen = new TableLayoutPanel();
            lblCardTituloResultado = new Label();
            lblEstadoHumano = new Label();
            lblResumenUnaLinea = new Label();
            chkVerDetalle = new CheckBox();
            pnlAccentBarResultado = new Panel();
            statusStrip = new StatusStrip();
            toolStripStatusLabel = new ToolStripStatusLabel();
            toolStripVersion = new ToolStripStatusLabel();
            toolTipRutas = new ToolTip(components);
            folderBrowserDialog = new FolderBrowserDialog();
            openFileDialogPlantilla = new OpenFileDialog();
            pnlHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)picLogo).BeginInit();
            tlpMain.SuspendLayout();
            pnlCardPeriodo.SuspendLayout();
            tlpPeriodo.SuspendLayout();
            pnlCardRutas.SuspendLayout();
            pnlCardEjecucion.SuspendLayout();
            pnlCardResultado.SuspendLayout();
            pnlDetalleTecnico.SuspendLayout();
            tlpResultadoResumen.SuspendLayout();
            statusStrip.SuspendLayout();
            SuspendLayout();
            // 
            // pnlHeader
            // 
            pnlHeader.BackColor = SystemColors.Window;
            pnlHeader.Controls.Add(lblTitulo);
            pnlHeader.Controls.Add(lblSubtitulo);
            pnlHeader.Controls.Add(lblVersionBadge);
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
            // lblVersionBadge
            // 
            lblVersionBadge.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            lblVersionBadge.AutoSize = true;
            lblVersionBadge.BorderStyle = BorderStyle.FixedSingle;
            lblVersionBadge.ForeColor = SystemColors.GrayText;
            lblVersionBadge.Location = new Point(712, 14);
            lblVersionBadge.MinimumSize = new Size(64, 20);
            lblVersionBadge.Name = "lblVersionBadge";
            lblVersionBadge.Padding = new Padding(4);
            lblVersionBadge.Size = new Size(64, 30);
            lblVersionBadge.TabIndex = 2;
            lblVersionBadge.Text = "v1.0.0";
            lblVersionBadge.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // picLogo
            // 
            picLogo.AccessibleName = "Slogan FNTecnologia";
            picLogo.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            picLogo.BackColor = SystemColors.Window;
            picLogo.Location = new Point(784, 4);
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
            tlpMain.ColumnCount = 1;
            tlpMain.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpMain.Controls.Add(pnlCardPeriodo, 0, 0);
            tlpMain.Controls.Add(pnlCardRutas, 0, 1);
            tlpMain.Controls.Add(pnlCardEjecucion, 0, 2);
            tlpMain.Controls.Add(pnlCardResultado, 0, 3);
            tlpMain.Dock = DockStyle.Fill;
            tlpMain.Location = new Point(0, 56);
            tlpMain.Name = "tlpMain";
            tlpMain.Padding = new Padding(16);
            tlpMain.RowCount = 4;
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.RowStyles.Add(new RowStyle());
            tlpMain.Size = new Size(960, 598);
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
            pnlCardPeriodo.MinimumSize = new Size(0, 104);
            pnlCardPeriodo.Name = "pnlCardPeriodo";
            pnlCardPeriodo.Padding = new Padding(16);
            pnlCardPeriodo.Size = new Size(916, 104);
            pnlCardPeriodo.TabIndex = 0;
            // 
            // pnlAccentBarPeriodo
            // 
            pnlAccentBarPeriodo.BackColor = SystemColors.Highlight;
            pnlAccentBarPeriodo.Dock = DockStyle.Left;
            pnlAccentBarPeriodo.Location = new Point(16, 16);
            pnlAccentBarPeriodo.Name = "pnlAccentBarPeriodo";
            pnlAccentBarPeriodo.Size = new Size(4, 70);
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
            tlpPeriodo.Location = new Point(20, 38);
            tlpPeriodo.Name = "tlpPeriodo";
            tlpPeriodo.RowCount = 2;
            tlpPeriodo.RowStyles.Add(new RowStyle());
            tlpPeriodo.RowStyles.Add(new RowStyle());
            tlpPeriodo.Size = new Size(358, 56);
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
            pnlCardRutas.Location = new Point(22, 138);
            pnlCardRutas.Margin = new Padding(6);
            pnlCardRutas.MinimumSize = new Size(0, 176);
            pnlCardRutas.Name = "pnlCardRutas";
            pnlCardRutas.Padding = new Padding(16);
            pnlCardRutas.Size = new Size(916, 176);
            pnlCardRutas.TabIndex = 1;
            // 
            // pnlAccentBarRutas
            // 
            pnlAccentBarRutas.BackColor = SystemColors.Highlight;
            pnlAccentBarRutas.Dock = DockStyle.Left;
            pnlAccentBarRutas.Location = new Point(16, 16);
            pnlAccentBarRutas.Name = "pnlAccentBarRutas";
            pnlAccentBarRutas.Size = new Size(4, 142);
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
            txtCarpetaFuentes.Location = new Point(128, 48);
            txtCarpetaFuentes.Name = "txtCarpetaFuentes";
            txtCarpetaFuentes.PlaceholderText = "Sin seleccionar";
            txtCarpetaFuentes.ReadOnly = true;
            txtCarpetaFuentes.Size = new Size(720, 27);
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
            txtPlantilla.Location = new Point(128, 84);
            txtPlantilla.Name = "txtPlantilla";
            txtPlantilla.PlaceholderText = "Sin seleccionar";
            txtPlantilla.ReadOnly = true;
            txtPlantilla.Size = new Size(720, 27);
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
            txtCarpetaSalida.Location = new Point(128, 120);
            txtCarpetaSalida.Name = "txtCarpetaSalida";
            txtCarpetaSalida.PlaceholderText = "Sin seleccionar";
            txtCarpetaSalida.ReadOnly = true;
            txtCarpetaSalida.Size = new Size(720, 27);
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
            pnlCardEjecucion.Controls.Add(btnLimpiar);
            pnlCardEjecucion.Controls.Add(btnAbrirSalida);
            pnlCardEjecucion.Controls.Add(btnCopiarLog);
            pnlCardEjecucion.Controls.Add(progressBar);
            pnlCardEjecucion.Controls.Add(lblProgresoPct);
            pnlCardEjecucion.Controls.Add(lblAseActual);
            pnlCardEjecucion.Dock = DockStyle.Fill;
            pnlCardEjecucion.Location = new Point(22, 326);
            pnlCardEjecucion.Margin = new Padding(6);
            pnlCardEjecucion.MinimumSize = new Size(0, 168);
            pnlCardEjecucion.Name = "pnlCardEjecucion";
            pnlCardEjecucion.Padding = new Padding(16);
            pnlCardEjecucion.Size = new Size(916, 168);
            pnlCardEjecucion.TabIndex = 2;
            // 
            // pnlAccentBarEjecucion
            // 
            pnlAccentBarEjecucion.BackColor = SystemColors.Highlight;
            pnlAccentBarEjecucion.Dock = DockStyle.Left;
            pnlAccentBarEjecucion.Location = new Point(16, 16);
            pnlAccentBarEjecucion.Name = "pnlAccentBarEjecucion";
            pnlAccentBarEjecucion.Size = new Size(4, 134);
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
            chkCincoAse.Size = new Size(152, 24);
            chkCincoAse.TabIndex = 3;
            chkCincoAse.Text = "Procesar los 5 ASE";
            chkCincoAse.UseVisualStyleBackColor = true;
            chkCincoAse.CheckedChanged += chkCincoAse_CheckedChanged;
            // 
            // btnEjecutar
            // 
            btnEjecutar.BackColor = SystemColors.Highlight;
            btnEjecutar.FlatAppearance.BorderSize = 0;
            btnEjecutar.FlatStyle = FlatStyle.Flat;
            btnEjecutar.ForeColor = SystemColors.HighlightText;
            btnEjecutar.Location = new Point(20, 78);
            btnEjecutar.Name = "btnEjecutar";
            btnEjecutar.Size = new Size(140, 28);
            btnEjecutar.TabIndex = 4;
            btnEjecutar.Text = "▶ Ejecutar";
            btnEjecutar.UseVisualStyleBackColor = false;
            btnEjecutar.Click += btnEjecutar_Click;
            // 
            // btnLimpiar
            // 
            btnLimpiar.BackColor = SystemColors.Window;
            btnLimpiar.FlatAppearance.BorderColor = SystemColors.ControlDark;
            btnLimpiar.FlatStyle = FlatStyle.Flat;
            btnLimpiar.ForeColor = SystemColors.WindowText;
            btnLimpiar.Location = new Point(172, 78);
            btnLimpiar.Name = "btnLimpiar";
            btnLimpiar.Size = new Size(110, 28);
            btnLimpiar.TabIndex = 5;
            btnLimpiar.Text = "Limpiar";
            btnLimpiar.UseVisualStyleBackColor = false;
            btnLimpiar.Click += btnLimpiar_Click;
            // 
            // btnAbrirSalida
            // 
            btnAbrirSalida.BackColor = SystemColors.Window;
            btnAbrirSalida.FlatAppearance.BorderColor = SystemColors.ControlDark;
            btnAbrirSalida.FlatStyle = FlatStyle.Flat;
            btnAbrirSalida.ForeColor = SystemColors.WindowText;
            btnAbrirSalida.Location = new Point(294, 78);
            btnAbrirSalida.Name = "btnAbrirSalida";
            btnAbrirSalida.Size = new Size(120, 28);
            btnAbrirSalida.TabIndex = 6;
            btnAbrirSalida.Text = "Abrir salida";
            btnAbrirSalida.UseVisualStyleBackColor = false;
            btnAbrirSalida.Click += btnAbrirSalida_Click;
            // 
            // btnCopiarLog
            // 
            btnCopiarLog.BackColor = SystemColors.Window;
            btnCopiarLog.Enabled = false;
            btnCopiarLog.FlatAppearance.BorderColor = SystemColors.ControlDark;
            btnCopiarLog.FlatStyle = FlatStyle.Flat;
            btnCopiarLog.ForeColor = SystemColors.WindowText;
            btnCopiarLog.Location = new Point(426, 78);
            btnCopiarLog.Name = "btnCopiarLog";
            btnCopiarLog.Size = new Size(120, 28);
            btnCopiarLog.TabIndex = 7;
            btnCopiarLog.Text = "Copiar log";
            btnCopiarLog.UseVisualStyleBackColor = false;
            btnCopiarLog.Click += btnCopiarLog_Click;
            // 
            // progressBar
            // 
            progressBar.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            progressBar.Location = new Point(72, 118);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(544, 14);
            progressBar.Style = ProgressBarStyle.Continuous;
            progressBar.TabIndex = 8;
            progressBar.TabStop = false;
            // 
            // lblProgresoPct
            // 
            lblProgresoPct.Location = new Point(20, 118);
            lblProgresoPct.Name = "lblProgresoPct";
            lblProgresoPct.Size = new Size(44, 15);
            lblProgresoPct.TabIndex = 9;
            lblProgresoPct.Text = "0 %";
            lblProgresoPct.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblAseActual
            // 
            lblAseActual.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            lblAseActual.AutoEllipsis = true;
            lblAseActual.Location = new Point(624, 118);
            lblAseActual.Name = "lblAseActual";
            lblAseActual.Size = new Size(264, 15);
            lblAseActual.TabIndex = 10;
            lblAseActual.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // pnlCardResultado
            // 
            pnlCardResultado.AutoSize = true;
            pnlCardResultado.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            pnlCardResultado.BackColor = SystemColors.Window;
            pnlCardResultado.BorderStyle = BorderStyle.FixedSingle;
            pnlCardResultado.Controls.Add(pnlDetalleTecnico);
            pnlCardResultado.Controls.Add(tlpResultadoResumen);
            pnlCardResultado.Controls.Add(pnlAccentBarResultado);
            pnlCardResultado.Dock = DockStyle.Fill;
            pnlCardResultado.Location = new Point(22, 506);
            pnlCardResultado.Margin = new Padding(6);
            pnlCardResultado.MinimumSize = new Size(0, 52);
            pnlCardResultado.Name = "pnlCardResultado";
            pnlCardResultado.Padding = new Padding(16, 12, 16, 12);
            pnlCardResultado.Size = new Size(916, 203);
            pnlCardResultado.TabIndex = 3;
            // 
            // pnlDetalleTecnico
            // 
            pnlDetalleTecnico.Controls.Add(lblLineasLog);
            pnlDetalleTecnico.Controls.Add(txtLog);
            pnlDetalleTecnico.Dock = DockStyle.Top;
            pnlDetalleTecnico.Location = new Point(20, 39);
            pnlDetalleTecnico.MaximumSize = new Size(0, 150);
            pnlDetalleTecnico.MinimumSize = new Size(0, 150);
            pnlDetalleTecnico.Name = "pnlDetalleTecnico";
            pnlDetalleTecnico.Padding = new Padding(0, 20, 0, 0);
            pnlDetalleTecnico.Size = new Size(878, 150);
            pnlDetalleTecnico.TabIndex = 4;
            pnlDetalleTecnico.Visible = false;
            // 
            // lblLineasLog
            // 
            lblLineasLog.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            lblLineasLog.AutoSize = true;
            lblLineasLog.ForeColor = SystemColors.GrayText;
            lblLineasLog.Location = new Point(810, 0);
            lblLineasLog.Name = "lblLineasLog";
            lblLineasLog.Size = new Size(59, 20);
            lblLineasLog.TabIndex = 0;
            lblLineasLog.Text = "0 líneas";
            // 
            // txtLog
            // 
            txtLog.Dock = DockStyle.Fill;
            txtLog.Font = new Font("Consolas", 9F);
            txtLog.Location = new Point(0, 20);
            txtLog.Multiline = true;
            txtLog.Name = "txtLog";
            txtLog.ReadOnly = true;
            txtLog.ScrollBars = ScrollBars.Both;
            txtLog.Size = new Size(878, 130);
            txtLog.TabIndex = 1;
            txtLog.TabStop = false;
            txtLog.WordWrap = false;
            // 
            // tlpResultadoResumen
            // 
            tlpResultadoResumen.AutoSize = true;
            tlpResultadoResumen.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            tlpResultadoResumen.ColumnCount = 4;
            tlpResultadoResumen.ColumnStyles.Add(new ColumnStyle());
            tlpResultadoResumen.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 250F));
            tlpResultadoResumen.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            tlpResultadoResumen.ColumnStyles.Add(new ColumnStyle());
            tlpResultadoResumen.Controls.Add(lblCardTituloResultado, 0, 0);
            tlpResultadoResumen.Controls.Add(lblEstadoHumano, 1, 0);
            tlpResultadoResumen.Controls.Add(lblResumenUnaLinea, 2, 0);
            tlpResultadoResumen.Controls.Add(chkVerDetalle, 3, 0);
            tlpResultadoResumen.Dock = DockStyle.Top;
            tlpResultadoResumen.Location = new Point(20, 12);
            tlpResultadoResumen.Name = "tlpResultadoResumen";
            tlpResultadoResumen.RowCount = 1;
            tlpResultadoResumen.RowStyles.Add(new RowStyle());
            tlpResultadoResumen.Size = new Size(878, 27);
            tlpResultadoResumen.TabIndex = 0;
            // 
            // lblCardTituloResultado
            // 
            lblCardTituloResultado.AutoSize = true;
            lblCardTituloResultado.Font = new Font("Segoe UI Semibold", 10F);
            lblCardTituloResultado.ForeColor = SystemColors.WindowText;
            lblCardTituloResultado.Location = new Point(0, 0);
            lblCardTituloResultado.Margin = new Padding(0, 0, 12, 0);
            lblCardTituloResultado.Name = "lblCardTituloResultado";
            lblCardTituloResultado.Size = new Size(86, 23);
            lblCardTituloResultado.TabIndex = 0;
            lblCardTituloResultado.Text = "Resultado";
            lblCardTituloResultado.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblEstadoHumano
            // 
            lblEstadoHumano.AutoEllipsis = true;
            lblEstadoHumano.Dock = DockStyle.Fill;
            lblEstadoHumano.ForeColor = SystemColors.WindowText;
            lblEstadoHumano.Location = new Point(101, 0);
            lblEstadoHumano.Margin = new Padding(3, 0, 12, 0);
            lblEstadoHumano.Name = "lblEstadoHumano";
            lblEstadoHumano.Size = new Size(235, 27);
            lblEstadoHumano.TabIndex = 1;
            lblEstadoHumano.Text = "Listo — elija período y rutas y pulse Ejecutar.";
            lblEstadoHumano.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // lblResumenUnaLinea
            // 
            lblResumenUnaLinea.AutoEllipsis = true;
            lblResumenUnaLinea.Dock = DockStyle.Fill;
            lblResumenUnaLinea.ForeColor = SystemColors.GrayText;
            lblResumenUnaLinea.Location = new Point(351, 0);
            lblResumenUnaLinea.Margin = new Padding(3, 0, 12, 0);
            lblResumenUnaLinea.Name = "lblResumenUnaLinea";
            lblResumenUnaLinea.Size = new Size(361, 27);
            lblResumenUnaLinea.TabIndex = 2;
            lblResumenUnaLinea.Text = "Aún no hay ejecución en esta sesión.";
            lblResumenUnaLinea.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // chkVerDetalle
            // 
            chkVerDetalle.Anchor = AnchorStyles.Right;
            chkVerDetalle.AutoSize = true;
            chkVerDetalle.Location = new Point(724, 3);
            chkVerDetalle.Margin = new Padding(0, 3, 0, 0);
            chkVerDetalle.Name = "chkVerDetalle";
            chkVerDetalle.Size = new Size(154, 24);
            chkVerDetalle.TabIndex = 3;
            chkVerDetalle.Text = "Ver detalle técnico";
            chkVerDetalle.UseVisualStyleBackColor = true;
            chkVerDetalle.CheckedChanged += chkVerDetalle_CheckedChanged;
            // 
            // pnlAccentBarResultado
            // 
            pnlAccentBarResultado.BackColor = SystemColors.Highlight;
            pnlAccentBarResultado.Dock = DockStyle.Left;
            pnlAccentBarResultado.Location = new Point(16, 12);
            pnlAccentBarResultado.Name = "pnlAccentBarResultado";
            pnlAccentBarResultado.Size = new Size(4, 177);
            pnlAccentBarResultado.TabIndex = 5;
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
            // Form1
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            ClientSize = new Size(960, 680);
            Controls.Add(tlpMain);
            Controls.Add(pnlHeader);
            Controls.Add(statusStrip);
            MinimumSize = new Size(940, 660);
            Name = "Form1";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "Remuneración Quincenal UAESP";
            pnlHeader.ResumeLayout(false);
            pnlHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)picLogo).EndInit();
            tlpMain.ResumeLayout(false);
            tlpMain.PerformLayout();
            pnlCardPeriodo.ResumeLayout(false);
            pnlCardPeriodo.PerformLayout();
            tlpPeriodo.ResumeLayout(false);
            tlpPeriodo.PerformLayout();
            pnlCardRutas.ResumeLayout(false);
            pnlCardRutas.PerformLayout();
            pnlCardEjecucion.ResumeLayout(false);
            pnlCardEjecucion.PerformLayout();
            pnlCardResultado.ResumeLayout(false);
            pnlCardResultado.PerformLayout();
            pnlDetalleTecnico.ResumeLayout(false);
            pnlDetalleTecnico.PerformLayout();
            tlpResultadoResumen.ResumeLayout(false);
            tlpResultadoResumen.PerformLayout();
            statusStrip.ResumeLayout(false);
            statusStrip.PerformLayout();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion
    }
}
