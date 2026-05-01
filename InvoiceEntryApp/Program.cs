using System;
using System.Windows.Forms;

namespace InvoiceEntryApp
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (!AppSettingsService.EnsureExcelPathConfigured(null))
            {
                return;
            }

            Application.Run(new InvoiceEntryForm());
        }
    }
}
