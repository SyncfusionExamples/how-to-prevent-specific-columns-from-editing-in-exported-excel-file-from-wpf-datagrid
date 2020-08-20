using Microsoft.Win32;
using Syncfusion.UI.Xaml.Grid;
using Syncfusion.UI.Xaml.Grid.Converter;
using Syncfusion.XlsIO;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace SfDataGridDemo
{
    public static class Commands
    {
        static Commands()
        {
            CommandManager.RegisterClassCommandBinding(typeof(SfDataGrid), new CommandBinding(ExportToExcel, OnExecuteExportToExcel, OnCanExecuteExportToExcel));
        }

        static Color gridHeaderBackgroundColor, gridHeaderForeGroundColor, gridCellBackgroundColor;

        #region ExportToExcel Command

        public static RoutedCommand ExportToExcel = new RoutedCommand("ExportToExcel", typeof(SfDataGrid));

    private static void GetDataGridStyles(SfDataGrid dataGrid)
    {
        var gridHeaderCellControl = dataGrid.FindResource(typeof(GridHeaderCellControl)) as Style;
        var gridCell = dataGrid.FindResource(typeof(GridCell)) as Style;

        if (gridHeaderCellControl == null || gridCell == null)
            return;

        foreach (Setter setter in gridHeaderCellControl.Setters)
        {
            if (setter.Property == GridHeaderCellControl.BackgroundProperty)
                gridHeaderBackgroundColor = (Color)ColorConverter.ConvertFromString(setter.Value.ToString());
            else if (setter.Property == GridHeaderCellControl.ForegroundProperty)
                gridHeaderForeGroundColor = (Color)ColorConverter.ConvertFromString(setter.Value.ToString());
        }

        foreach (Setter setter in gridCell.Setters)
        {
            if (setter.Property== GridCell.BackgroundProperty)
                gridCellBackgroundColor = (Color)ColorConverter.ConvertFromString(setter.Value.ToString());
        }
    }
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

        private static void ExportingHandler(object sender, GridExcelExportingEventArgs e)
        {
            if (e.CellType == ExportCellType.HeaderCell)
            {
                e.CellStyle.BackGroundBrush = new SolidColorBrush(gridHeaderBackgroundColor);
                e.CellStyle.ForeGroundBrush = new SolidColorBrush(gridHeaderForeGroundColor);
            }
            else if (e.CellType == ExportCellType.RecordCell)
            {
                e.CellStyle.BackGroundBrush = new SolidColorBrush(gridCellBackgroundColor);
            }
            e.Handled = true;
        }

        private static void OnCanExecuteExportToExcel(object sender, CanExecuteRoutedEventArgs args)
        {
            args.CanExecute = true;
        }
        #endregion
    }
}
