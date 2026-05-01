# InvoiceEntryApp

This project rebuilds the reference invoice entry desktop app as a Windows Forms application and adds permanent Excel saving on `Submit`.

## Open and edit

1. Open `D:\projects\ERP application\InvoiceEntryApp\InvoiceEntryApp.csproj` in Visual Studio 2022 or later.
2. If Visual Studio asks to restore packages, allow it.

## Build the executable

1. From Visual Studio, build or publish the project.
2. In this workspace, the published runnable output is already available at:
   `D:\projects\ERP application\publish\InvoiceEntryApp\InvoiceEntryApp.exe`

## Excel file selection

On first launch, the application now shows a standard Windows dialog that lets the user choose or create the `.xlsx` workbook used for invoice storage.

The selected path is saved locally for future launches in the user's local app data folder.

## Package dependency

The app uses the following NuGet package for Excel writing:
- `ClosedXML` `0.104.2`

If you need to restore manually from the command line, use a NuGet-enabled build environment and restore packages before building.
