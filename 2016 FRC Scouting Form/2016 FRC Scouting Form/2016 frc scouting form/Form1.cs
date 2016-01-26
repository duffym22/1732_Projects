using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace _2016_FRC_Scouting_Form
{
    public partial class Form1 : Form
    {
        private const String _DATA_DIRECTORY = "C:\\2016_FRC_Scouting\\";
        private const String _EXCEL_FILENAME = "2016_FRC_Scouting_Data.xlsx";
        private const String dataSheet = "Scouting_Data";
        private Microsoft.Office.Interop.Excel.Application _xlApp;
        private Workbook _xlwb;
        private Worksheet _xlws;
        private object misValue = System.Reflection.Missing.Value;
        private DataGrid dataGridView1;


        public Form1()
        {
            InitializeComponent();
        }

        private void btn_submitData_Click(object sender, EventArgs e)
        {
            if (initExcel()) //initialize Excel object
            {
                //check if file already exists
                //File will be stored locally at C:\2016_FRC_Scouting\
                if (!verifyExistingDataFile())
                {
                    //if not exist - create new file (force creation of file in directory where executable is run from)
                    createNewDataFile();
                }
                //access file if not open (if file not exist, will be created in condition above)
                openDataFile();
                addDataRow();
                clearExcelObjects();


                //gather data from form
                //dataGridView1.DataSource = gatherData();  
                //set into specific format
            }
        }

        private void addDataRow()
        {
            //Range cell = _xlws
        }

        private void openDataFile()
        {
            _xlwb = _xlApp.Workbooks.Open(_DATA_DIRECTORY + _EXCEL_FILENAME);
            _xlws = _xlwb.Worksheets.get_Item(dataSheet);
        }

        private void clearExcelObjects()
        {
            releaseExcelObject(_xlws);
            releaseExcelObject(_xlwb);
            releaseExcelObject(_xlApp);
        }

        private void releaseExcelObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                HandleException(ex, "Exception occured while releasing excel object");
            }
            finally
            {
                GC.Collect();
            }
        }

        private void createNewDataFile()
        {
            try
            {
                _xlwb = _xlApp.Workbooks.Add(Type.Missing);
                _xlws = _xlwb.ActiveSheet;
                _xlws.Name = dataSheet; 
                _xlws = (Worksheet)_xlwb.Worksheets[1];

                _xlws.Cells[1, 1] = "Team#";
                _xlws.Cells[1, 2] = "Match#";
                _xlws.Cells[1, 3] = "Scout_Name";
                _xlws.Cells[1, 4] = "Alliance";
                _xlws.Cells[1, 5] = "Auto_Defense_Reached";
                _xlws.Cells[1, 6] = "Auto_Defense_Crossed";
                _xlws.Cells[1, 7] = "Auto_Low_Goal_Scored";
                _xlws.Cells[1, 8] = "Auto_High_Goal_Scored";
                _xlws.Cells[1, 9] = "Auto_Starting_Position";
                _xlws.Cells[1, 10] = "Auto_Ending_Position";
                _xlws.Cells[1, 11] = "Tele_Portcullis";
                _xlws.Cells[1, 12] = "Tele_Fries";
                _xlws.Cells[1, 13] = "Tele_Rampart";
                _xlws.Cells[1, 14] = "Tele_Moat";
                _xlws.Cells[1, 15] = "Tele_Drawbridge";
                _xlws.Cells[1, 16] = "Tele_SallyPort";
                _xlws.Cells[1, 17] = "Tele_RockWall";
                _xlws.Cells[1, 18] = "Tele_RoughTerrain";
                _xlws.Cells[1, 19] = "Tele_LowBar";
                _xlws.Cells[1, 20] = "Tele_Low_Goal_Scored";
                _xlws.Cells[1, 21] = "Tele_High_Goal_Scored";
                _xlws.Cells[1, 22] = "End_Challenged";
                _xlws.Cells[1, 23] = "End_Scaled";
                _xlws.Cells[1, 24] = "Notes";

                _xlwb.SaveAs(_DATA_DIRECTORY + _EXCEL_FILENAME, XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                _xlwb.Close(true, misValue, misValue);
                _xlApp.Quit();
            }
            catch (Exception ex)
            {
                HandleException(ex);
            }
        }

        private Boolean verifyExistingDataFile()
        {
            Boolean rtval = false;
            if(!Directory.Exists(_DATA_DIRECTORY)) //create data directory if it doesn't exist
            {
                Directory.CreateDirectory(_DATA_DIRECTORY);
            }
            else
            {
                if (!File.Exists(_DATA_DIRECTORY + _EXCEL_FILENAME))
                    MessageBox.Show("File does not exist - creating file");
                else
                {
                    rtval = true;
                    MessageBox.Show("File exists - continuing");
                }
            }
            return rtval;

        }

        private Boolean initExcel()
        {
            Boolean rtval = false;
            if (_xlApp != null) //object already initialized - don't initialize again
                rtval = true;
            else
            {
                _xlApp = new Microsoft.Office.Interop.Excel.Application();
                if (_xlApp == null)
                    MessageBox.Show("Error - Cannot continue - Microsoft Excel is not properly installed");
                else
                {
                    _xlApp.Visible = false;
                    _xlApp.DisplayAlerts = false;
                    rtval = true;
                }
            }
            return rtval;
        }

        private void btn_clearData_Click(object sender, EventArgs e)
        {
            clearAllCheckboxes(this);
            clearAllRadioButtons(this);
            clearAllTextBoxes(this);
            clearAllRichTextBoxes(this);
        }

        private void clearAllRichTextBoxes(Control ctrl)
        {
            RichTextBox rtxt = ctrl as RichTextBox;
            if (rtxt == null)
            {
                foreach (Control child in ctrl.Controls)
                {
                    clearAllRichTextBoxes(child); //recursive
                }
            }
            else
            {
                rtxt.Text = String.Empty;
            }
        }

        private void clearAllTextBoxes(Control ctrl)
        {
            System.Windows.Forms.TextBox txt = ctrl as System.Windows.Forms.TextBox;
            if (txt == null)
            {
                foreach (Control child in ctrl.Controls)
                {
                    clearAllTextBoxes(child); //recursive
                }
            }
            else
            {
                txt.Text = String.Empty;
            }
        }

        private void clearAllRadioButtons(Control ctrl)
        {
            RadioButton rdoBtn = ctrl as RadioButton;
            if (rdoBtn == null)
            {
                foreach (Control child in ctrl.Controls)
                {
                    clearAllRadioButtons(child); //recursive
                }
            }
            else
            {
                rdoBtn.Checked = false;
            }
        }

        private void clearAllCheckboxes(Control ctrl)
        {
            System.Windows.Forms.CheckBox chkBox = ctrl as System.Windows.Forms.CheckBox;
            if (chkBox == null )
            {
                foreach (Control child in ctrl.Controls)
                {
                    clearAllCheckboxes(child); //recursive
                }
            }
            else
            {
                chkBox.Checked = false;
            }
        }

        private void HandleException(Exception ex, String message = "")
        {
            if(message.Equals(""))
                MessageBox.Show("Exception thrown:" + Environment.NewLine + ex.Message + Environment.NewLine + ex.InnerException + Environment.NewLine + ex.StackTrace);
            else
                MessageBox.Show(message + Environment.NewLine + ex.Message + Environment.NewLine + ex.InnerException + Environment.NewLine + ex.StackTrace);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
