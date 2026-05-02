# Invoice Application

A Windows desktop invoice entry application built with **C# WinForms** that recreates a classic ERP-style invoice entry workflow and saves submitted invoice data directly into an **Excel `.xlsx` workbook**.

This project was rebuilt from scratch to match a reference invoice-entry desktop app as closely as possible while adding one practical enhancement:

- persistent Excel-based invoice storage

## What it does

The application provides a fixed-layout invoice entry screen with:

- company details
- customer details
- invoice number and invoice date
- editable item grid
- subtotal, GST, and total fields
- submit and cancel actions

On `Submit`, the app validates the form and appends the invoice data to an Excel workbook using a **one-row-per-item** format.

## Key features

- Windows desktop UI built with WinForms
- ERP-style invoice entry layout
- editable `DataGridView` for item rows
- automatic subtotal and total calculation
- Excel saving with [`ClosedXML`](https://github.com/ClosedXML/ClosedXML)
- no Excel installation required
- first-run Excel file selection dialog
- saved workbook path reused on later launches
- single-file `.exe` publish output available for distribution

## First-run behavior

When the application is opened for the first time, it shows a standard Windows file dialog asking the user to choose or create the Excel workbook that will store invoice data.

That selected path is saved locally and reused for future launches.

## Excel save model

Each valid item row is saved as a separate row in the workbook, repeating the invoice-level, company-level, and customer-level information for that item.

### Worksheet

- default sheet name: `InvoiceData`

### Excel columns

- Saved Timestamp
- Invoice No
- Invoice Date
- Company Name
- Company Address
- Company City
- Company State
- Company Pin Code
- Company Contact No
- Company Tin
- Customer Name
- Customer Address
- Customer City
- Customer State
- Customer Pin Code
- Customer Contact No
- Customer Tin
- Item No
- Description
- Quantity
- Price
- Line Total
- SubTotal
- GST
- Total

## Tech stack

- C#
- .NET 8 Windows
- Windows Forms
- ClosedXML

## Project structure

```text
InvoiceEntryApp/
  AppSettingsService.cs
  ExcelInvoiceSaver.cs
  InvoiceEntryForm.cs
  InvoiceModels.cs
  InvoiceValidator.cs
  Program.cs
```

## Download and run

If you download this repository as a ZIP, run the rebuilt application from:

```text
Release/InvoiceEntryApp.exe
```

Important:

- `Release/InvoiceEntryApp.exe` is the rebuilt application
- `Reference/Order Entry.exe` is only the original reference demo asset and is not the app you should use

On first launch, the rebuilt app will ask you to choose or create the Excel workbook used for invoice storage.

## Run the application

If you already have a published build, run:

```text
publish/InvoiceEntryApp-SingleFile/InvoiceEntryApp.exe
```

On launch:

1. choose or create the Excel workbook
2. enter invoice details
3. click `Submit`

## Open in Visual Studio

1. Open `InvoiceEntryApp/InvoiceEntryApp.csproj`
2. Allow package restore if prompted
3. Build or publish the project

## Publish from command line

Example:

```powershell
dotnet publish .\InvoiceEntryApp\InvoiceEntryApp.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true
```

## Notes

- The published `.exe` is intended for Windows.
- The workbook path is no longer hardcoded into the application.
- The app is designed for a stable, automation-friendly UI flow.

## Reference assets

The `Reference/` folder contains the original demo assets used to guide the rebuild.
