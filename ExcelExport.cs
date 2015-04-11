using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Views.Grid;
using DevExpress.XtraGrid.Columns;

namespace HiimelOyun.App.Lib.Utils.Export
{
    public class ExcelExport
    {
        public static void toExcel(string filename, Control control)
        {
            if (control is GridControl) 
            {
                GridView gridView = (GridView)((GridControl)control).MainView;
                toExcel(filename, gridView, null);
            }
            else if (control is ListView) 
            {
                toExcel(filename, (ListView)control, null);
            }
        }

        public static void toExcel(string filename, GridView gridView, DataSet ds)
        {
            int row = 1;
            int col = 1;
            int tableIndex = 0;
            int colLength = 0;

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = null;
            Microsoft.Office.Interop.Excel.Range Cel = null;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            if (ds == null)
            {
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(tableIndex + 1);

                colLength = gridView.Columns.Count;
                GridColumn column = null;
                for (int i = 0; i < gridView.DataRowCount; i++)
                {
                    if (row == 2) i = 0;
                    object item = gridView.GetRow(i);
                    for (col = 1; col <= colLength; )
                    {
                        column = gridView.Columns[col - 1];
                        if (row == 1)
                        {
                            xlWorkSheet.Cells[row, col] = "" + column.Caption;

                            Cel = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                            Cel.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                , Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                , Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                                , 1);
                            Cel.Interior.Color = ColorTranslator.ToOle(Color.Silver);
                            Cel.Font.Color = ColorTranslator.ToOle(Color.Black);
                            Cel.Font.Bold = true;
                            Cel.ColumnWidth = column.Width / 4.5;
                        }
                        else
                        {
                            try
                            {
                                xlWorkSheet.Cells[row, col] = "" + AppUtil.GetPropertyValue(item, column.FieldName);
                            }
                            catch { }
                        }
                        col++;
                    }
                    row++;
                }
            }
            else
                for (tableIndex = 0; tableIndex < ds.Tables.Count; tableIndex++)
                {
                    DataTable dt = ds.Tables[tableIndex];
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(tableIndex + 1);

                    if (gridView != null && gridView.Columns.Count > 0)
                    {
                        colLength = gridView.Columns.Count;
                        GridColumn column = null;
                        for (int i = 0; i < gridView.DataRowCount; i++)
                        {
                            if (row == 2) i = 0;
                            DataRow dr = dt.Rows[i];
                            for (col = 1; col <= colLength; )
                            {
                                column = gridView.Columns[col - 1];
                                if (row == 1)
                                {
                                    xlWorkSheet.Cells[row, col] = "" + column.Caption;

                                    Cel = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                                    Cel.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                        , Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                        , Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                                        , 1);
                                    Cel.Interior.Color = ColorTranslator.ToOle(Color.Silver);
                                    Cel.Font.Color = ColorTranslator.ToOle(Color.Black);
                                    Cel.Font.Bold = true;
                                    Cel.ColumnWidth = ("" + column.Caption).Length * 4.5;
                                }
                                else
                                {
                                    xlWorkSheet.Cells[row, col] = "" + dr[column.FieldName];
                                }
                                col++;
                            }
                            row++;
                        }
                    }
                    else
                    {
                        colLength = dt.Columns.Count;
                        DataColumn column = null;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (row == 2) i = 0;
                            DataRow dr = dt.Rows[i];
                            for (col = 1; col <= colLength; )
                            {
                                column = dt.Columns[col - 1];
                                if (row == 1)
                                {
                                    xlWorkSheet.Cells[row, col] = "" + column.ColumnName;

                                    Cel = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                                    Cel.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                        , Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                        , Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                                        , 1);
                                    Cel.Interior.Color = ColorTranslator.ToOle(Color.Silver);
                                    Cel.Font.Color = ColorTranslator.ToOle(Color.Black);
                                    Cel.Font.Bold = true;
                                    Cel.ColumnWidth = ("" + column.ColumnName).Length * 4.5;
                                }
                                else
                                {
                                    xlWorkSheet.Cells[row, col] = "" + dr[column.ColumnName];
                                }
                                col++;
                            }
                            row++;
                        }
                    }
                }

            xlWorkBook.SaveAs(filename
                , Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                , misValue, misValue, misValue, misValue
                , Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive
                , misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Process p;
            ProcessStartInfo pInfo;
            try
            {
                pInfo = new ProcessStartInfo();
                pInfo.Verb = "open";
                pInfo.FileName = filename;
                pInfo.UseShellExecute = true;
                pInfo.WindowStyle = ProcessWindowStyle.Maximized;

                p = Process.Start(pInfo);
            }
            catch { }
        }

        public static void toExcel(string filename, ListView listView, DataSet ds)
        {
            int row = 1;
            int col = 1;
            int tableIndex = 0;
            int colLength = 0;
            
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = null;
            Microsoft.Office.Interop.Excel.Range Cel = null;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);

            if (ds == null)
            {
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(tableIndex + 1);

                colLength = listView.Columns.Count;
                ColumnHeader column = null;
                for (int i = 0; i < listView.Items.Count; i++ )
                {
                    if (row == 2) i = 0;
                    ListViewItem li = listView.Items[i];
                    for (col = 1; col <= colLength; )
                    {
                        column = listView.Columns[col - 1];
                        if (row == 1)
                        {
                            xlWorkSheet.Cells[row, col] = "" + column.Text;

                            Cel = (Microsoft.Office.Interop.Excel.Range) xlWorkSheet.Cells[row, col];
                            Cel.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                , Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                , Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                                , 1);
                            Cel.Interior.Color = ColorTranslator.ToOle(Color.Silver);
                            Cel.Font.Color = ColorTranslator.ToOle(Color.Black);
                            Cel.Font.Bold = true;
                            Cel.ColumnWidth = column.Width / 4.5;
                        }
                        else
                        {
                            try
                            {
                                if (col == 1)
                                    xlWorkSheet.Cells[row, col] = "" + li.Text;
                                else
                                    xlWorkSheet.Cells[row, col] = "" + li.SubItems[col - 1].Text;
                            }
                            catch { }
                        }                        
                        col++;
                    }
                    row++;
                }
            }
            else
                for (tableIndex = 0; tableIndex < ds.Tables.Count; tableIndex++)
                {
                    DataTable dt = ds.Tables[tableIndex];
                    xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(tableIndex + 1);

                    if (listView != null && listView.Columns.Count > 0)
                    {
                        colLength = listView.Columns.Count;
                        ColumnHeader column = null;
                        for (int i = 0; i < listView.Items.Count; i++ )
                        {
                            if (row == 2) i = 0;
                            DataRow dr = dt.Rows[i];
                            for (col = 1; col <= colLength; )
                            {
                                column = listView.Columns[col - 1];
                                if (row == 1)
                                {
                                    xlWorkSheet.Cells[row, col] = "" + column.Text;

                                    Cel = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                                    Cel.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                        , Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                        , Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                                        , 1);
                                    Cel.Interior.Color = ColorTranslator.ToOle(Color.Silver);
                                    Cel.Font.Color = ColorTranslator.ToOle(Color.Black);
                                    Cel.Font.Bold = true;
                                    Cel.ColumnWidth = ("" + column.Text).Length * 4.5;
                                }
                                else
                                {
                                    xlWorkSheet.Cells[row, col] = "" + dr[column.Tag.ToString()];
                                }
                                col++;
                            }
                            row++;
                        }
                    }
                    else
                    {
                        colLength = dt.Columns.Count;
                        DataColumn column = null;
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            if (row == 2) i = 0;
                            DataRow dr = dt.Rows[i];
                            for (col = 1; col <= colLength; )
                            {
                                column = dt.Columns[col - 1];
                                if (row == 1)
                                {
                                    xlWorkSheet.Cells[row, col] = "" + column.ColumnName;

                                    Cel = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                                    Cel.BorderAround(Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous
                                        , Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin
                                        , Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic
                                        , 1);
                                    Cel.Interior.Color = ColorTranslator.ToOle(Color.Silver);
                                    Cel.Font.Color = ColorTranslator.ToOle(Color.Black);
                                    Cel.Font.Bold = true;
                                    Cel.ColumnWidth = ("" + column.ColumnName).Length * 4.5;
                                }
                                else
                                {
                                    xlWorkSheet.Cells[row, col] = "" + dr[column.ColumnName];
                                }
                                col++;
                            }
                            row++;
                        }
                    }
                }

            xlWorkBook.SaveAs(filename
                , Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal
                , misValue, misValue, misValue, misValue
                , Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive
                , misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            Process p;
            ProcessStartInfo pInfo;
            try
            {
                pInfo = new ProcessStartInfo();
                pInfo.Verb = "open";
                pInfo.FileName = filename;
                pInfo.UseShellExecute = true;
                pInfo.WindowStyle = ProcessWindowStyle.Maximized;

                p = Process.Start(pInfo);
            }
            catch { }
        }

        public static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
