using System;
using System.Windows.Forms;

namespace ExcelTemplateManager
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            try
            {
                Application.SetHighDpiMode(HighDpiMode.SystemAware);
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

                // Gestione globale delle eccezioni non gestite
                Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(
                    (sender, e) => MessageBox.Show($"Si è verificato un errore: {e.Exception.Message}", 
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error));

                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(
                    (sender, e) => MessageBox.Show($"Si è verificato un errore critico: {((Exception)e.ExceptionObject).Message}", 
                    "Errore critico", MessageBoxButtons.OK, MessageBoxIcon.Error));

                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Si è verificato un errore durante l'avvio: {ex.Message}", 
                    "Errore di avvio", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}