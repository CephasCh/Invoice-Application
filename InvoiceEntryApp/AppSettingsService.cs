using System;
using System.IO;
using System.Windows.Forms;

namespace InvoiceEntryApp
{
    internal static class AppSettingsService
    {
        private const string SettingsDirectoryName = "InvoiceEntryApp";
        private const string SettingsFileName = "excel-path.txt";

        private static readonly string SettingsDirectoryPath =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), SettingsDirectoryName);

        private static readonly string SettingsFilePath =
            Path.Combine(SettingsDirectoryPath, SettingsFileName);

        public static string GetConfiguredExcelPath()
        {
            if (!File.Exists(SettingsFilePath))
            {
                return string.Empty;
            }

            return File.ReadAllText(SettingsFilePath).Trim();
        }

        public static bool EnsureExcelPathConfigured(IWin32Window owner)
        {
            string existingPath = GetConfiguredExcelPath();
            if (!string.IsNullOrWhiteSpace(existingPath))
            {
                return true;
            }

            MessageBox.Show(
                owner,
                "Select the Excel workbook that should store invoice data. You can choose an existing .xlsx file or create a new one.",
                "Choose Excel File",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);

            using (SaveFileDialog dialog = new SaveFileDialog())
            {
                dialog.Title = "Select Excel Workbook";
                dialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
                dialog.DefaultExt = "xlsx";
                dialog.AddExtension = true;
                dialog.OverwritePrompt = false;
                dialog.FileName = "InvoiceData.xlsx";

                string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (!string.IsNullOrWhiteSpace(documentsPath) && Directory.Exists(documentsPath))
                {
                    dialog.InitialDirectory = documentsPath;
                }

                if (dialog.ShowDialog(owner) != DialogResult.OK)
                {
                    MessageBox.Show(
                        owner,
                        "The application needs an Excel file before it can save invoices.",
                        "Excel File Required",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return false;
                }

                SaveExcelPath(dialog.FileName);
                return true;
            }
        }

        public static void SaveExcelPath(string excelPath)
        {
            Directory.CreateDirectory(SettingsDirectoryPath);
            File.WriteAllText(SettingsFilePath, excelPath ?? string.Empty);
        }
    }
}
