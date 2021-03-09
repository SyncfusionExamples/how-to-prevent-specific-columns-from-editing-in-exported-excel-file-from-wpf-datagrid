# How to prevent specific columns from editing in exported excel file from WPF DataGrid (SfDataGrid)?

This sample show cases how to prevent the specific columns from editing in exported excel file from [WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid)?

# About the sample

You can prevent editing of some columns in [WPF DataGrid](https://www.syncfusion.com/wpf-ui-controls/datagrid) (SfDataGrid). You can get the same behavior in exported excel file by protecting the worksheet and set the `Locked` property to false if you want to edit the columns.

```c#
private static void OnExecuteExportToExcel(object sender, ExecutedRoutedEventArgs args)
{
    var dataGrid = args.Source as SfDataGrid;
    if (dataGrid == null) return;
    try
    {
        GetDataGridStyles(dataGrid);

        // Creating an instance for ExcelExportingOptions which is passed as a parameter to the ExportToExcel method.
        ExcelExportingOptions options = new ExcelExportingOptions();
        options.AllowOutlining = true;
        options.ExportingEventHandler = ExportingHandler;
        // Exports Datagrid to Excel and returns ExcelEngine.
        var excelEngine = dataGrid.ExportToExcel(dataGrid.View, options);
        // Gets the exported workbook from the ExcelEngine
        var workBook = excelEngine.Excel.Workbooks[0];
        var workSheet = workBook.Worksheets[0];
        var gridColumns = dataGrid.Columns.Where(col => col.AllowEditing).ToList();
        foreach(var column in workSheet.Columns)
        {
            if(gridColumns.Any(gridCol => gridCol.HeaderText == column.DisplayText))
                workSheet.Range[column.AddressLocal].CellStyle.Locked = false;
        }
        workSheet.Protect("syncfusion");

        // Saving the workbook.
        SaveFileDialog sfd = new SaveFileDialog
        {
            FilterIndex = 2,
            Filter = "Excel 97 to 2003 Files(*.xls)|*.xls|Excel 2007 to 2010 Files(*.xlsx)|*.xlsx"
        };

        if (sfd.ShowDialog() == true)
        {
            using (Stream stream = sfd.OpenFile())
            {
                if (sfd.FilterIndex == 1)
                    workBook.Version = ExcelVersion.Excel97to2003;
                else
                    workBook.Version = ExcelVersion.Excel2010;
                workBook.SaveAs(stream);
            }

            //Message box confirmation to view the created spreadsheet.
            if (MessageBox.Show("Do you want to view the workbook?", "Workbook has been created",
                                MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
            {
                //Launching the Excel file using the default Application.[MS Excel Or Free ExcelViewer]
                System.Diagnostics.Process.Start(sfd.FileName);
            }
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show(ex.Message);
    }
}
```
## Requirements to run the demo
 Visual Studio 2015 and above versions
