using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using Excel = Microsoft.Office.Interop.Excel;


namespace UpdatePartsFile
{
    internal static class Program
    {

        //*************************************************************************
        // GLOBAL PROGRAM VARIABLES AND CONSTANTS
        //************************************************************************
        
        public static Excel.Workbook? g_wb = null;
        public static string gstrFileMode = "";
        public static string gstrAbapAppServer = "P01";
        public static string gstrPlant = "0300";
        public static string gstrCustomerNo = "";
        public static string gstrFileToOpen = "";
        public static string gstrPartsFile_0300 = "C:\\QwDataTrans2\\SAP_PARTS_0300.xlsm";
        public static string gstrPartsFile_0310 = "C:\\QwDataTrans2\\SAP_PARTS_0310.xlsm";
        public static string gstrPartsSheetName = "SapParts";
        public static string gstrPartsTableName = "tblSapParts";
        

        public static string gstrVersionNumber = "2.3";
        public static bool gblnOfflineMode = false;
        public static DateTime gdtmLastUpdate = DateTime.MinValue;
        public static bool gblnAutoMode = false;

        // Add a unique mutex name for this application
        private static readonly string MutexName = "QwDataTransfer2_SingleInstanceMutex";

        // Import user32.dll functions for window activation
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_RESTORE = 9;


        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            using (Mutex mutex = new Mutex(true, MutexName, out bool isNewInstance))
            {
                if (!isNewInstance)
                {
                    // Bring the existing instance to the foreground
                    BringOtherInstanceToFront();

                    //MessageBox.Show(
                    //    "Multiple copies of the program cannot be started.",
                    //    "QwDataTransfer2",
                    //    MessageBoxButtons.OK,
                    //    MessageBoxIcon.Warning
                    //);
                    return;
                }

                ApplicationConfiguration.Initialize();

                if (args.Length > 0)
                {
                    gstrFileToOpen = args[0];
                    gstrAbapAppServer = args[1];
                    gstrPlant = args[2];
                    gblnAutoMode = true;
                    
                }
                Application.Run(new frmMain());
            }

        }
        private static void BringOtherInstanceToFront()
        {
            try
            {
                // Get the current process
                Process current = Process.GetCurrentProcess();
                // Find the other process with the same name but different Id
                foreach (Process process in Process.GetProcessesByName(current.ProcessName))
                {
                    if (process.Id != current.Id)
                    {
                        IntPtr hWnd = process.MainWindowHandle;
                        if (hWnd != IntPtr.Zero)
                        {
                            // Restore if minimized
                            ShowWindow(hWnd, SW_RESTORE);
                            // Bring to foreground
                            SetForegroundWindow(hWnd);
                        }
                        break;
                    }
                }
            }
            catch
            {
                // Ignore any errors in bringing to front
            }
        }
        public static bool IsFileOpen(string docPath)
        {

            bool fileStatus = false;
            // get a reference to the currently open workbook
            Excel.Workbook? wb = GetOpenWorkbook(docPath);
            if (wb != null)
            {
                fileStatus = true;
            }
            else
            {
                fileStatus = false;
            }

            return fileStatus;

        }
        public static class Marshal2
        {
            //*************************************************************************
            // copied from stackoverflow.com
            // https://stackoverflow.com/questions/58010510/no-definition-found-for-getactiveobject-from-system-runtime-interopservices-mars
            // we need this to get a reference to an active Excel application object
            // so we can use it to get the active workbook and worksheet
            //*************************************************************************
            internal const string OLEAUT32 = "oleaut32.dll";
            internal const string OLE32 = "ole32.dll";

            [SecurityCritical]  // auto-generated_required
            public static object GetActiveObject(string progID)
            {
                object obj = null;
                Guid clsid;

                // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
                // CLSIDFromProgIDEx doesn't exist.
                try
                {
                    CLSIDFromProgIDEx(progID, out clsid);
                }
                //            catch
                catch (Exception)
                {
                    CLSIDFromProgID(progID, out clsid);
                }

                GetActiveObject(ref clsid, nint.Zero, out obj);
                return obj;
            }

            //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
            [DllImport(OLE32, PreserveSig = false)]
            [ResourceExposure(ResourceScope.None)]
            [SuppressUnmanagedCodeSecurity]
            [SecurityCritical]  // auto-generated
            private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

            //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
            [DllImport(OLE32, PreserveSig = false)]
            [ResourceExposure(ResourceScope.None)]
            [SuppressUnmanagedCodeSecurity]
            [SecurityCritical]  // auto-generated
            private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] string progId, out Guid clsid);

            //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
            [DllImport(OLEAUT32, PreserveSig = false)]
            [ResourceExposure(ResourceScope.None)]
            [SuppressUnmanagedCodeSecurity]
            [SecurityCritical]  // auto-generated
            private static extern void GetActiveObject(ref Guid rclsid, nint reserved, [MarshalAs(UnmanagedType.Interface)] out object ppunk);

        }
        public static Excel.Workbook? GetOpenWorkbook(string fullPath)
        {
            try
            {
                // Get the running Excel application
                Excel.Application excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
                foreach (Excel.Workbook wb in excelApp.Workbooks)
                {
                    // check if workbook file name contains a url
                    if (wb.FullName.StartsWith("https://lizprominent.sharepoint.com", StringComparison.OrdinalIgnoreCase))
                    {

                        string fileName = Path.GetFileName(fullPath);
                        string url = wb.FullName;

                        if (url.IndexOf(fileName, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            return wb;
                        }
                    }
                    if (string.Equals(System.IO.Path.GetFullPath(wb.FullName), System.IO.Path.GetFullPath(fullPath), StringComparison.OrdinalIgnoreCase))
                    {
                        return wb;
                    }
                }
            }
            catch
            {
                // Excel is not running or not accessible
            }
            return null;
        }

        public static Excel.Workbook? OpenWorkbook(string fullPath)
        {
            if (string.IsNullOrWhiteSpace(fullPath) || !File.Exists(fullPath))
                throw new FileNotFoundException("The specified Excel file does not exist.", fullPath);

            Excel.Application? excelApp = null;
            bool createdNewInstance = false;
            try
            {
                // Try to get the running Excel application
                try
                {
                    excelApp = (Excel.Application)Marshal2.GetActiveObject("Excel.Application");
                }
                catch
                {
                    // Excel is not running, so create a new instance
                    excelApp = new Excel.Application();
                    createdNewInstance = true;
                }

                excelApp.Visible = true; // Set to false if you want Excel hidden

                // Open the workbook in the application instance
                Excel.Workbook wb = excelApp.Workbooks.Open(fullPath, ReadOnly: false);
                return wb;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening Excel workbook: " + ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Clean up if Excel was started but workbook failed to open
                if (excelApp != null && createdNewInstance)
                {
                    try { excelApp.Quit(); } catch { }
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                return null;
            }
        }
        public static void SaveSapParts(Excel.Worksheet ws, DataTable dt)
        {

            try
            {

                Excel.ListObject table = ws.ListObjects["tblSapParts"];
                


                // set the cursor to wait
                Cursor.Current = Cursors.WaitCursor;
                Application.DoEvents();
                // unprotect the sheet
                ws.Unprotect();


              

                // Store the old last row of the table
                int oldLastRow = table.Range.Row + table.ListRows.Count;

                // Calculate new range for the table (header + data rows)
                int headerRow = table.Range.Row;
                int firstCol = table.Range.Column;
                int lastCol = table.Range.Column + table.Range.Columns.Count - 1;
                int newRowCount = dt.Rows.Count;

                // If there are no rows, keep only the header
                int newLastRow = newRowCount > 0 ? headerRow + newRowCount : headerRow;

                var newRange = ws.Range[
                    ws.Cells[headerRow, firstCol],
                    ws.Cells[newLastRow, lastCol]
                ];
                table.Resize(newRange);


                // Example: Set columns 1 and 3 as text (1-based index)
                int[] textColumns = { 1, 13 };
                foreach (int col in textColumns)
                {
                    Excel.Range colRange = table.DataBodyRange.Columns[col];
                    colRange.NumberFormat = "@";
                }



                // Replace data in the table with data from the DataTable

                // Prepare a 2D object array
                object[,] values = new object[dt.Rows.Count, dt.Columns.Count];
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        if (textColumns.Contains(c))
                            values[r, c] = dt.Rows[r][c]?.ToString();
                        else
                            values[r, c] = dt.Rows[r][c];
                    }
                }

                // Assign the array to the range in one operation
                table.DataBodyRange.Value2 = values;

                // Delete rows below the new table range if any
                int lastTableRow = table.Range.Row + table.Range.Rows.Count - 1;
                int lastUsedRow = ws.UsedRange.Rows.Count;

                // Only clear if there are rows below the table
                if (lastTableRow < lastUsedRow)
                {
                    Excel.Range clearRange = ws.Range[
                        ws.Cells[lastTableRow + 1, table.Range.Column],
                        ws.Cells[lastUsedRow, table.Range.Column + table.Range.Columns.Count - 1]
                    ];
                    clearRange.Clear(); // or clearRange.ClearContents();
                }

                // save the excel workbook
                g_wb.Save();


            
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error writing data to Excel: " + ex.Message, "Excel Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
                Application.DoEvents();
            }


        }
    }
}