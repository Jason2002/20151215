using System;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Newegg.Ozzo.RMA.Service.Common;
using Excel = Microsoft.Office.Interop.Excel;

namespace Newegg.Ozzo.RMA.Client.WinModel.Common
{
    public enum IMEXMode
    {
        IMEXExportMode = 1,
        IMEXImportMode = 2,
        IMEXLinkedMode = 3
    }

    public class ExcelAccessToolkit
    {
        private const string connectionString = "Provider={0};Data Source={1};Extended Properties=\"Excel {2};HDR={3};IMEX={4}\"";
        private const string TabString = "\t";

        private static string GetConnectionStringByFile(string fileName, bool firstRowAsHeader, IMEXMode mode)
        {
            string fileType = Path.GetExtension(fileName);
            switch (fileType)
            {
                case ".xls":
                    return string.Format(connectionString, "Microsoft.Jet.OLEDB.4.0", fileName, "8.0", firstRowAsHeader ? "YES" : "NO", (int)mode);
                case ".xlsx":
                    return string.Format(connectionString, "Microsoft.ACE.OLEDB.12.0", fileName, "12.0", firstRowAsHeader ? "YES" : "NO", (int)mode);
                default:
                    return string.Format(connectionString, "Microsoft.Jet.OLEDB.4.0", fileName, "8.0", firstRowAsHeader ? "YES" : "NO", (int)mode);
            }
        }

        public static DataSet GetExcelDataByAdo(string fileName, bool hasHeader)
        {
            try
            {
                string connStr = GetConnectionStringByFile(fileName, hasHeader, IMEXMode.IMEXExportMode);

                using (OleDbConnection conn = new OleDbConnection(connStr))
                {
                    conn.Open();
                    DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    DataSet ds = new DataSet();
                    OleDbDataAdapter da = new OleDbDataAdapter();

                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        string sheet = (string)dt.Rows[i]["TABLE_NAME"];

                        if (sheet.Contains("$") && !sheet.Replace("'", "").EndsWith("$"))
                            continue;

                        da.SelectCommand = new OleDbCommand(string.Format("SELECT * FROM [{0}]", sheet), conn);
                        da.Fill(ds, sheet);
                    }

                    return ds;
                }
            }
            catch
            {
                throw new Exception("The selected file is not a valid excel file.");
            }
        }

        public static bool DataSet2ExcelCom(string fileName, DataSet ds, bool hasHeader)
        {
            if (ds == null || ds.Tables.Count <= 0)
            {
                return false;
            }

            Excel.ApplicationClass xlsApp = new Excel.ApplicationClass();
            xlsApp.DisplayAlerts = true;
            Excel.Workbook xlsWorkBook = xlsApp.Workbooks.Add(Type.Missing);

            try
            {
                int sheetCount = xlsWorkBook.Worksheets.Count;
                int tableCount = ds.Tables.Count;
                if (sheetCount < tableCount)
                {
                    xlsWorkBook.Worksheets.Add(Type.Missing, xlsWorkBook.Worksheets[sheetCount], tableCount - sheetCount, Excel.XlSheetType.xlWorksheet);
                }

                Excel.Worksheet xlsWorkSheetPrevios = null;
                for (int i = 0; i < tableCount; i++)
                {
                    Excel.Worksheet xlsWorkSheet = (Excel.Worksheet)xlsWorkBook.Worksheets[i + 1];
                    string sheetName = ds.Tables[i].TableName.Replace("'", "").TrimEnd('$');
                    try
                    {
                        xlsWorkSheet.Name = sheetName;
                    }
                    catch
                    {
                        Excel.Worksheet dupSheet = ((Excel.Worksheet)xlsWorkBook.Worksheets[sheetName]);
                        if (dupSheet != null)
                        {
                            dupSheet.Delete();
                        }

                        Excel.Worksheet newSheet;
                        if (xlsWorkSheetPrevios != null)
                        {
                            newSheet = (Excel.Worksheet)xlsWorkBook.Worksheets.Add(Type.Missing, xlsWorkSheetPrevios, 1, Excel.XlSheetType.xlWorksheet);
                        }
                        else
                        {
                            newSheet = (Excel.Worksheet)xlsWorkBook.Worksheets.Add(Type.Missing, Type.Missing, 1, Excel.XlSheetType.xlWorksheet);
                        }
                        newSheet.Name = sheetName;
                        xlsWorkSheet = newSheet;
                    }

                    xlsWorkSheet.Activate();
                    //xlsWorkSheet.EnableCalculation = false;
                    StringBuilder stringBuffer = new StringBuilder();

                    if (hasHeader)
                    {
                        for (int h = 0; h < ds.Tables[i].Columns.Count; h++)
                        {
                            stringBuffer.Append(ds.Tables[i].Columns[h].Caption);
                            if (h < ds.Tables[i].Columns.Count - 1)
                            {
                                stringBuffer.Append(TabString);
                            }
                        }
                        stringBuffer.Append(Environment.NewLine);
                    }

                    for (int r = 0; r < ds.Tables[i].Rows.Count; r++)
                    {
                        for (int c = 0; c < ds.Tables[i].Columns.Count; c++)
                        {
                            stringBuffer.Append(ds.Tables[i].Rows[r][c] != null
                                ? MakeSafeExcelString(ds.Tables[i].Rows[r][c].ToString())
                                : "");
                            if (c < ds.Tables[i].Columns.Count - 1)
                            {
                                stringBuffer.Append(TabString);
                            }
                        }

                        if (r < ds.Tables[i].Rows.Count - 1)
                        {
                            stringBuffer.Append(Environment.NewLine);
                        }
                    }

                    Clipboard.SetDataObject("");
                    Clipboard.SetDataObject(stringBuffer.ToString());
                    Excel.Range range = ((Excel.Range)(xlsWorkSheet).Cells[1, 1]);
                    range.Select();
                    xlsWorkSheet.Paste(Type.Missing, Type.Missing);
                    range.Select();

                    if (hasHeader)
                    {
                        Excel.Range headerRange = xlsWorkSheet.Range[xlsWorkSheet.Cells[1, 1], xlsWorkSheet.Cells[1, ds.Tables[i].Columns.Count]];
                        headerRange.Font.Bold = true;
                        headerRange.Font.Size = 12;
                        headerRange.ColumnWidth = 20;
                        headerRange.HorizontalAlignment = 3;
                        //headerRange.EntireColumn.AutoFit();
                    }

                    Clipboard.SetDataObject("");
                    xlsWorkSheetPrevios = xlsWorkSheet;
                }

                ((Excel.Worksheet)xlsWorkBook.Worksheets[1]).Activate();
                xlsWorkBook.SaveAs(fileName, Excel.XlSaveAction.xlSaveChanges, Type.Missing, Type.Missing, false, false,
                                Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                xlsWorkBook.Close(false, null, null);
            }
            catch (Exception ex)
            {
                throw new BizUIException(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(xlsWorkBook);
                xlsWorkBook = null;
                xlsApp.Quit();
                Marshal.ReleaseComObject(xlsApp);
                xlsApp = null;
                GC.Collect();
            }

            return true;
        }

        public static void DataSet2ExcelAdo(string fileName, DataSet ds)
        {
            string connStr = GetConnectionStringByFile(fileName, true, IMEXMode.IMEXLinkedMode);
            using (OleDbConnection conn = new OleDbConnection(connStr))
            {
                conn.Open();
                foreach (DataTable dt in ds.Tables)
                {
                    string fieldsCreateString = "";
                    string fieldsInsertString = "";
                    string tableName = "";
                    int rowsCount = dt.Rows.Count;
                    int colNum = dt.Columns.Count;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        fieldsCreateString += "[" + dc.Caption + "] VarChar,";
                        fieldsInsertString += "[" + dc.Caption + "],";
                    }
                    fieldsCreateString = fieldsCreateString.TrimEnd(',');
                    fieldsInsertString = fieldsInsertString.TrimEnd(',');
                    tableName = dt.TableName.Replace("$", "");

                    string sqlCreate = "CREATE TABLE [" + tableName + "] (" + fieldsCreateString + ")";
                    OleDbCommand cmd = new OleDbCommand(sqlCreate, conn);
                    cmd.ExecuteNonQuery();

                    for (int r = 0; r < rowsCount; r++)
                    {
                        string sqlValues = "";
                        for (int c = 0; c < colNum; c++)
                        {
                            sqlValues += "'" + dt.Rows[r][c] + "',";
                        }
                        sqlValues = sqlValues.TrimEnd(',');
                        cmd.CommandText = "INSERT INTO [" + tableName + "] (" + fieldsInsertString + ") VALUES (" + sqlValues + ")";
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        private static string MakeSafeExcelString(string str)
        {
            if (str.IsNullOrEmptyEx())
            {
                return "";
            }

            if (str.StartsWith("="))
            {
                str = " " + str;
            }

            return StringHelper.ReplaceNewLine(str, " ").Replace(TabString, "");
        }
    }
}
