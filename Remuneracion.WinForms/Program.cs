namespace Remuneracion.WinForms
{
    internal static class Program
    {
        [STAThread]
        static void Main()
        {
            ApplicationConfiguration.Initialize();

            // Plan 43 (O-1): composición ÚNICA del host (oráculo de validaciones + emisor del reporte
            // durable SIEMPRE presentes). El constructor por defecto de Form1 usa la MISMA
            // composición, así que ninguna vía de arranque puede cerrar en verde sin validar.
            var (procesador, procesadorPeriodo, locator) = ComposicionServicios.Crear();

            Application.Run(new Form1(procesador, procesadorPeriodo, locator));
        }
    }
}
