using System;
using System.IO;
using ClosedXML.Excel;

namespace InvoiceEntryApp
{
    internal sealed class ExcelInvoiceSaver
    {
        private const string WorksheetName = "InvoiceData";

        private static readonly string[] Headers =
        {
            "Saved Timestamp",
            "Invoice No",
            "Invoice Date",
            "Company Name",
            "Company Address",
            "Company City",
            "Company State",
            "Company Pin Code",
            "Company Contact No",
            "Company Tin",
            "Customer Name",
            "Customer Address",
            "Customer City",
            "Customer State",
            "Customer Pin Code",
            "Customer Contact No",
            "Customer Tin",
            "Item No",
            "Description",
            "Quantity",
            "Price",
            "Line Total",
            "SubTotal",
            "GST",
            "Total"
        };

        public void Save(InvoiceSubmission submission)
        {
            string excelFilePath = AppSettingsService.GetConfiguredExcelPath();
            if (string.IsNullOrWhiteSpace(excelFilePath))
            {
                throw new InvalidOperationException("No Excel file is configured for invoice saving.");
            }

            string directory = Path.GetDirectoryName(excelFilePath);
            if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            using (XLWorkbook workbook = File.Exists(excelFilePath) ? new XLWorkbook(excelFilePath) : new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.Worksheets.Contains(WorksheetName)
                    ? workbook.Worksheet(WorksheetName)
                    : workbook.Worksheets.Add(WorksheetName);

                EnsureHeaders(worksheet);

                int nextRow = worksheet.LastRowUsed()?.RowNumber() + 1 ?? 2;
                string savedTimestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                string invoiceDate = submission.InvoiceDate.ToString("yyyy-MM-dd");

                foreach (InvoiceItem item in submission.Items)
                {
                    worksheet.Cell(nextRow, 1).Value = savedTimestamp;
                    worksheet.Cell(nextRow, 2).Value = submission.InvoiceNo;
                    worksheet.Cell(nextRow, 3).Value = invoiceDate;
                    worksheet.Cell(nextRow, 4).Value = submission.Company.Name;
                    worksheet.Cell(nextRow, 5).Value = submission.Company.Address;
                    worksheet.Cell(nextRow, 6).Value = submission.Company.City;
                    worksheet.Cell(nextRow, 7).Value = submission.Company.State;
                    worksheet.Cell(nextRow, 8).Value = submission.Company.PinCode;
                    worksheet.Cell(nextRow, 9).Value = submission.Company.ContactNo;
                    worksheet.Cell(nextRow, 10).Value = submission.Company.Tin;
                    worksheet.Cell(nextRow, 11).Value = submission.Customer.Name;
                    worksheet.Cell(nextRow, 12).Value = submission.Customer.Address;
                    worksheet.Cell(nextRow, 13).Value = submission.Customer.City;
                    worksheet.Cell(nextRow, 14).Value = submission.Customer.State;
                    worksheet.Cell(nextRow, 15).Value = submission.Customer.PinCode;
                    worksheet.Cell(nextRow, 16).Value = submission.Customer.ContactNo;
                    worksheet.Cell(nextRow, 17).Value = submission.Customer.Tin;
                    worksheet.Cell(nextRow, 18).Value = item.ItemNo;
                    worksheet.Cell(nextRow, 19).Value = item.Description;
                    worksheet.Cell(nextRow, 20).Value = item.Quantity;
                    worksheet.Cell(nextRow, 21).Value = item.Price;
                    worksheet.Cell(nextRow, 22).Value = item.LineTotal;
                    worksheet.Cell(nextRow, 23).Value = submission.SubTotal;
                    worksheet.Cell(nextRow, 24).Value = submission.Gst;
                    worksheet.Cell(nextRow, 25).Value = submission.Total;
                    nextRow++;
                }

                worksheet.Range(1, 1, 1, Headers.Length).Style.Font.Bold = true;
                worksheet.Columns().AdjustToContents();

                if (File.Exists(excelFilePath))
                {
                    workbook.Save();
                }
                else
                {
                    workbook.SaveAs(excelFilePath);
                }
            }
        }

        private static void EnsureHeaders(IXLWorksheet worksheet)
        {
            for (int index = 0; index < Headers.Length; index++)
            {
                worksheet.Cell(1, index + 1).Value = Headers[index];
            }
        }
    }
}
