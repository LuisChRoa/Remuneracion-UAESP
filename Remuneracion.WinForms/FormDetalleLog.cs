using System;
using System.Drawing;
using System.Windows.Forms;

namespace Remuneracion.WinForms
{
    /// <summary>
    /// HU-19 (§2.7): diálogo modal de solo lectura que muestra el detalle técnico de la última
    /// ejecución (el buffer de log en memoria de <see cref="Form1"/>). Reemplaza al panel inline de
    /// la card Resultado eliminada. Construido en código: sin lógica de negocio ni persistencia.
    /// </summary>
    public sealed class FormDetalleLog : Form
    {
        private readonly TextBox _txtContenido;
        private readonly Button _btnCopiar;

        public FormDetalleLog(string contenido)
        {
            Text = "Detalle técnico — logs";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ClientSize = new Size(760, 480);
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            BackColor = SystemColors.Control;
            Font = new Font("Segoe UI", 9F);

            var contenidoSeguro = contenido ?? string.Empty;

            var lblConteo = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = SystemColors.GrayText,
                Text = $"{ContarLineas(contenidoSeguro)} líneas"
            };

            _txtContenido = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                Font = new Font("Consolas", 9F),
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                TabStop = false,
                Text = contenidoSeguro
            };

            var pnlInferior = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 44,
                Padding = new Padding(8)
            };

            _btnCopiar = new Button
            {
                Text = "Copiar",
                Width = 90,
                Height = 28,
                Dock = DockStyle.Right,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                UseVisualStyleBackColor = false,
                TabIndex = 1
            };
            _btnCopiar.FlatAppearance.BorderColor = SystemColors.ControlDark;
            _btnCopiar.Click += btnCopiar_Click;

            var btnCerrar = new Button
            {
                Text = "Cerrar",
                Width = 90,
                Height = 28,
                Dock = DockStyle.Right,
                DialogResult = DialogResult.Cancel,
                FlatStyle = FlatStyle.Flat,
                BackColor = SystemColors.Window,
                ForeColor = SystemColors.WindowText,
                UseVisualStyleBackColor = false,
                TabIndex = 0
            };
            btnCerrar.FlatAppearance.BorderColor = SystemColors.ControlDark;

            // Dock: se agrega primero el Fill y luego los bordes; dentro del panel,
            // "Cerrar" se agrega último para quedar más a la derecha que "Copiar".
            pnlInferior.Controls.Add(_btnCopiar);
            pnlInferior.Controls.Add(btnCerrar);

            Controls.Add(_txtContenido);
            Controls.Add(lblConteo);
            Controls.Add(pnlInferior);

            // Esc cierra el diálogo (no hay acción primaria implícita con Enter).
            CancelButton = btnCerrar;
        }

        private void btnCopiar_Click(object? sender, EventArgs e)
        {
            if (_txtContenido.TextLength == 0)
            {
                MessageBox.Show(
                    this,
                    "Todavía no hay contenido para copiar.",
                    "Detalle técnico",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            try
            {
                Clipboard.SetText(_txtContenido.Text);
            }
            catch (Exception)
            {
                MessageBox.Show(
                    this,
                    "No se pudo copiar el contenido al portapapeles.",
                    "Detalle técnico",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

        private static int ContarLineas(string contenido)
        {
            if (string.IsNullOrEmpty(contenido))
            {
                return 0;
            }

            int lineas = contenido.Split('\n').Length;
            if (contenido.EndsWith("\n", StringComparison.Ordinal))
            {
                lineas--;
            }

            return lineas;
        }
    }
}
