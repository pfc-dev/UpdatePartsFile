using Microsoft.WindowsAPICodePack.Dialogs;
using SapData;
using System;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using Excel = Microsoft.Office.Interop.Excel;


namespace UpdatePartsFile
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            // set the default sap system
            cboSapSystem.Text = Program.gstrAbapAppServer;

            cboPlant.SelectedItem = Program.gstrPlant;
            lblDateLastUpdated.Text = "";

            // if the program was launched from the Excel Workbook, then
            // the global varial will already be set. 
            txtMain.Text = "Start Parameters: ";
            txtMain.Text += Environment.NewLine + "SAP System: " + Program.gstrAbapAppServer;
            txtMain.Text += Environment.NewLine + "SAP Plant: " + Program.gstrPlant;


            this.Show();
            Application.DoEvents();

            if ( Program.gblnAutoMode == true )
            {


                // load the workbook and update it automatically
                Program.g_wb = Program.GetOpenWorkbook(Program.gstrFileToOpen);
                // fill in the list of worksheet names on the combo box



                if (Program.g_wb != null)
                {


                    // set the last modified date of the file
                    Program.gdtmLastUpdate = File.GetLastWriteTime(Program.gstrFileToOpen);
                    lblDateLastUpdated.Text = Program.gdtmLastUpdate.ToString();
                    txtMain.Text += Environment.NewLine + "File to Update: " + Program.gstrFileToOpen;
                    Application.DoEvents();
                    // check the version of the file
                    string version = CheckFileVersion();
                    if (version != "OK")
                    {
                        // if the version is not compatible, then exit the program.
                        lblMessage.Text = "File is not compatible with this program version.";
                        this.Cursor = Cursors.Default;
                        // display a message box to the user
                        MessageBox.Show("File is not compatible with this program version.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    txtMain.Text += Environment.NewLine + "File Version: " + version;
                    Application.DoEvents();

                    // now update the file

                    UpdateFile();
                    QuitApplication();
                }
            }

        }
        private void btnChoseFile_Click(object sender, EventArgs e)
        {
            var ofd = new CommonOpenFileDialog();
            ofd.Filters.Add(new CommonFileDialogFilter("Excel Files", "*.xlsx;*.xlsm"));
            if (ofd.ShowDialog() == CommonFileDialogResult.Ok)
            {



                if (Program.IsFileOpen(ofd.FileName))
                {
                    Program.g_wb = Program.GetOpenWorkbook(ofd.FileName);
                }
                else
                {
                    Program.g_wb = Program.OpenWorkbook(ofd.FileName);
                }


                if (Program.g_wb == null)
                {
                    lblMessage.Text = "File could not be opened";
                    this.Cursor = Cursors.Default;
                    return;
                }


                // read custom properties from the file 
                // and check the version number.
                string version = CheckFileVersion();
                if (version != "OK")
                {
                    // if the version is not compatible, then exit the program.
                    lblMessage.Text = "File is not compatible with this program version.";
                    this.Cursor = Cursors.Default;
                    return;
                }

            }
            else
            {
                lblMessage.Text = "User cancelled";
                return;

            }


            // put the file name into the file textbox
            txtFile.Text = Program.g_wb.FullName;
            Program.gstrFileToOpen = ofd.FileName;

            // set the last modified date of the file
            Program.gdtmLastUpdate = File.GetLastWriteTime(Program.gstrFileToOpen);

            // update the screen with the last updated date
            lblDateLastUpdated.Text = Program.gdtmLastUpdate.ToString();

            this.Cursor = Cursors.Default;
            lblMessage.Text = "File loaded.";

        }
        private void cboSapSystem_SelectedValueChanged(object sender, EventArgs e)
        {
            lblMessage.Text = "checking SAP connection, please wait...";
            this.Cursor = Cursors.WaitCursor;
            Application.DoEvents();

            Program.gstrAbapAppServer = cboSapSystem.Text;
            //close the current connection.
            if (SapRfc.sapSystem != "")
                SapRfc.CloseSapDest(SapRfc.sapSystem);
            // open the new connection'
            SapRfc.OpenSapDest(Program.gstrAbapAppServer);
            if (SapRfc.SapDestOk())
            {
                lblMessage.Text = "SAP Connection: 'OK'";

                txtMain.Text += Environment.NewLine + "SAP Connection: 'OK'";
                

                //lblSapMode.Text = Program.gstrAbapAppServer + " Online Mode";
                // set the global variable to indicate online mode
                Program.gblnOfflineMode = false;
            }
            else
            {
                lblMessage.Text = "Unable to connect to SAP.";
                Program.gblnOfflineMode = true;
            }

            this.Cursor = Cursors.Default;
        }
        private void btnQuit_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "Do you want to close the application?",
                "Exit Application",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {

                QuitApplication();

            }
            // If No, do nothing (the popup will close automatically)
        }

        private void QuitApplication()
        {
            // Release Excel workbook COM object
            if (Program.g_wb != null)
            {
                //try
                //{
                //    Program.g_wb.Close(false); // Optionally close the workbook (false = don't save)
                //}
                //catch { /* Ignore errors if already closed */ }

                Marshal.ReleaseComObject(Program.g_wb);
                Program.g_wb = null;

            }

            // Force garbage collection to clean up any remaining COM references
            GC.Collect();
            GC.WaitForPendingFinalizers();

            Application.Exit();
        }

        private string CheckFileVersion()
        {
            //// check the version of the excel file,  it must be at least 1.0
            //// if not, then display a message and exit the program. 

            if (Program.g_wb is null) { return "0"; }

            var customProps = Program.g_wb.CustomDocumentProperties;
            string version = "";

            try
            {
                var prop = customProps["VersionNumber"];
                version = prop.Value.ToString();
            }
            catch
            {
                // Property not found or error
                version = "";
            }
            if (version == null || version == "")
            {
                lblMessage.Text = "Cannot find version number in the file.";
                return "E1";
            }
            if (version != Program.gstrVersionNumber)
            {
                lblMessage.Text = "This file is not compatible with this program version.";
                return "E2";
            }
            else
            {
                return "OK";
            }


        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
           
            UpdateFile();

        }
        private void UpdateFile()
        {
            string plant = "";
            if (cboPlant.SelectedItem is not null)
            {
                try
                {
                    plant = (string)cboPlant.SelectedItem;
                    lblMessage.Text = "Reading SAP Data, please wait...";
                    txtMain.Text += Environment.NewLine + "Reading SAP Data, please wait...";
                    this.Cursor = Cursors.WaitCursor;
                    Application.DoEvents();

                    // Run the long operation in a background thread
                    //var dt = await Task.Run(() => SapRfc.GetSapParts(plant));
                    var dt = SapRfc.GetSapParts(plant);
                    dt.AcceptChanges();

                    //dgvSapParts.DataSource = dt;
                    lblMessage.Text = dt.Rows.Count + " Records read from SAP.";
                    txtMain.Text += Environment.NewLine + dt.Rows.Count + " Records read from SAP.";
                    Application.DoEvents();
                    Excel.Worksheet? ws = null;
                    string targetWs = Program.gstrPartsSheetName;
                    foreach (Excel.Worksheet sheet in Program.g_wb.Worksheets)
                    {
                        if (sheet.Name.Equals(targetWs, StringComparison.OrdinalIgnoreCase))
                        {
                            ws = sheet;
                            break;
                        }
                    }





                    // now write the data to the excel file.
                    if (dt.Rows.Count > 0)
                    {
                        lblMessage.Text = "Writing data to Excel, please wait...";
                        txtMain.Text += Environment.NewLine + "Writing data to Excel, please wait...";
                        Application.DoEvents();
                        Program.SaveSapParts(ws, dt);
                        lblMessage.Text = " Excel file updated.";
                        txtMain.Text += Environment.NewLine + "Excel file updated.";
                        Application.DoEvents();
                    }
                    else
                    {
                        lblMessage.Text = "No records to write to Excel.";
                    }

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }

            }
        }

    }
}
