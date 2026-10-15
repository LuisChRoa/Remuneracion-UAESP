using Remuneracion.Core.Models;

namespace Remuneracion.WinForms;

// Plan 42 (T3-UI, R-UI-1, §2.3 Visual Design Intent): superficie de RECONOCIMIENTO EXPLÍCITO
// de divergencias declaradas. Es el BLOQUE DOMINANTE del cierre (no un toast ni una línea de
// log): primero QUÉ (hoja/celda + observado vs esperado + Δ), después DE QUIÉN ES (dueño
// pendiente + cita de evidencia), después QUÉ SE PIDE (acto de aceptación por esta corrida).
// Anti-costumbre (R-COSTUMBRE, D-D): nada pre-marcado, nada persistente entre corridas; cada
// corrida construye un diálogo nuevo y el cierre espera el acto. Construido en código como
// FormDetalleLog; solo recursos verificados por grep (SystemColors, Segoe UI/Segoe UI Semibold,
// FlatStyle.Flat, MessageBox por catálogo en el llamador).
public sealed class FormReconocimientoDivergencia : Form
{
    private readonly List<CheckBox> _aceptaciones = [];
    private readonly Button _btnAceptar;

    // Ids de entrada aceptados en ESTA corrida (los mismos que el CLI pasa por flag). Vacío si
    // el diálogo no cerró con aceptación.
    public IReadOnlyList<string> EntradasAceptadas { get; private set; } = [];

    public FormReconocimientoDivergencia(ClasificacionDivergenciasDeclaradas clasificacion)
    {
        ArgumentNullException.ThrowIfNull(clasificacion);
        if (!clasificacion.SoloDeclaradas)
        {
            throw new ArgumentException(
                "El reconocimiento solo se ofrece cuando lo único pendiente es divergencia declarada.",
                nameof(clasificacion));
        }

        Text = "Divergencia declarada — reconocimiento para esta corrida";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterParent;
        ClientSize = new Size(740, 560);
        MinimumSize = new Size(560, 420);
        MaximizeBox = false;
        MinimizeBox = false;
        ShowInTaskbar = false;
        BackColor = SystemColors.Control;
        Font = new Font("Segoe UI", 9F);
        AutoScroll = true;

        var filas = ReconocimientoDivergencia.AgruparPorDivergencia(clasificacion);
        var entradas = clasificacion.EntradasDeclaradas;

        var lblTitulo = new Label
        {
            Dock = DockStyle.Top,
            Height = 28,
            TextAlign = ContentAlignment.MiddleLeft,
            Font = new Font("Segoe UI Semibold", 10F),
            ForeColor = SystemColors.WindowText,
            Padding = new Padding(16, 4, 16, 0),
            Text = "La corrida encontró una diferencia conocida"
        };

        var lblSubtitulo = new Label
        {
            Dock = DockStyle.Top,
            Height = 48,
            ForeColor = SystemColors.GrayText,
            Padding = new Padding(16, 0, 16, 0),
            Text = "Revise el monto, el responsable y la evidencia antes de decidir. "
                + "Nada queda guardado: cada corrida vuelve a presentar esta diferencia."
        };

        var txtDetalle = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            WordWrap = true,
            BackColor = SystemColors.Window,
            ForeColor = SystemColors.WindowText,
            Margin = new Padding(16, 0, 16, 0),
            Text = ConstruirDetalle(filas, entradas),
            TabIndex = 0
        };

        var pnlAceptacion = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 40 + entradas.Count * 44,
            Padding = new Padding(16, 8, 16, 0),
            BackColor = SystemColors.Control
        };

        int tab = 1;
        // Dock.Top apila en orden inverso al agregado: se agregan en reverso para que el orden
        // visual siga al del catálogo.
        foreach (var entradaId in entradas.Reverse())
        {
            var entrada = ReconocimientoDivergencia.BuscarEntrada(entradaId);
            var filaRepresentativa = filas.First(v => string.Equals(v.EntradaId, entradaId, StringComparison.Ordinal));
            var chk = new CheckBox
            {
                Dock = DockStyle.Top,
                Height = 44,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.WindowText,
                AccessibleName = $"Aceptación de la divergencia {entradaId}",
                Tag = entradaId,
                Text = ReconocimientoDivergencia.TextoAceptacion(entradaId, filaRepresentativa.Delta, entrada?.Dueno),
                Checked = false,
                TabIndex = tab++
            };
            chk.CheckedChanged += Aceptacion_CheckedChanged;
            _aceptaciones.Add(chk);
            pnlAceptacion.Controls.Add(chk);
        }

        var lblHuella = new Label
        {
            Dock = DockStyle.Top,
            Height = 40,
            ForeColor = SystemColors.GrayText,
            Text = "La aceptación queda registrada en el reporte y en el log de esta corrida."
        };
        pnlAceptacion.Controls.Add(lblHuella);

        var pnlBotones = new Panel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            Padding = new Padding(8),
            BackColor = SystemColors.Control
        };

        _btnAceptar = new Button
        {
            Text = "Aceptar y cerrar",
            Width = 150,
            Height = 28,
            Dock = DockStyle.Right,
            FlatStyle = FlatStyle.Flat,
            BackColor = SystemColors.Highlight,
            ForeColor = SystemColors.HighlightText,
            UseVisualStyleBackColor = false,
            Enabled = false,
            TabIndex = tab++
        };
        _btnAceptar.FlatAppearance.BorderSize = 0;
        _btnAceptar.Click += btnAceptar_Click;

        var btnVolver = new Button
        {
            Text = "Volver sin cerrar",
            Width = 150,
            Height = 28,
            Dock = DockStyle.Right,
            DialogResult = DialogResult.Cancel,
            FlatStyle = FlatStyle.Flat,
            BackColor = SystemColors.Window,
            ForeColor = SystemColors.WindowText,
            UseVisualStyleBackColor = false,
            TabIndex = tab
        };
        btnVolver.FlatAppearance.BorderColor = SystemColors.ControlDark;

        pnlBotones.Controls.Add(_btnAceptar);
        pnlBotones.Controls.Add(btnVolver);

        Controls.Add(txtDetalle);
        Controls.Add(lblSubtitulo);
        Controls.Add(lblTitulo);
        Controls.Add(pnlAceptacion);
        Controls.Add(pnlBotones);

        // Foco inicial en el bloque de reconocimiento (primer acto), no en los botones.
        Shown += (_, _) => ActiveControl = (Control?)_aceptaciones.FirstOrDefault() ?? btnVolver;
        AcceptButton = _btnAceptar;
        CancelButton = btnVolver;
    }

    private void Aceptacion_CheckedChanged(object? sender, EventArgs e)
    {
        _btnAceptar.Enabled = _aceptaciones.Count > 0 && _aceptaciones.All(c => c.Checked);
    }

    private void btnAceptar_Click(object? sender, EventArgs e)
    {
        if (!_aceptaciones.All(c => c.Checked))
        {
            return;
        }

        EntradasAceptadas = _aceptaciones
            .Select(c => (string)c.Tag!)
            .ToList();
        DialogResult = DialogResult.OK;
        Close();
    }

    private static string ConstruirDetalle(
        IReadOnlyList<VeredictoDivergenciaValidacion> filas,
        IReadOnlyList<string> entradas)
    {
        var lineas = new List<string>();
        var numero = 1;
        foreach (var entradaId in entradas)
        {
            var entrada = ReconocimientoDivergencia.BuscarEntrada(entradaId);
            lineas.Add($"Diferencia {entradaId} ({numero} de {entradas.Count})");
            foreach (var fila in filas.Where(v => string.Equals(v.EntradaId, entradaId, StringComparison.Ordinal)))
            {
                lineas.Add(ReconocimientoDivergencia.TextoQue(fila));
            }

            foreach (var lineaDueno in ReconocimientoDivergencia.TextoDeQuienEs(entrada).Split(Environment.NewLine))
            {
                lineas.Add(lineaDueno);
            }

            lineas.Add(string.Empty);
            numero++;
        }

        return string.Join(Environment.NewLine, lineas);
    }
}
