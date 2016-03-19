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
using System.Diagnostics;
using System.Threading;
using System.Collections;
using System.Linq;

namespace _2016_FRC_Scouting_Form
{
    enum DATA_ROWS
    {
        Team_Num = 1,
        Match_Num = 2,
        Scout_Name = 3,
        Alliance = 4,
        Auto_Defense_Reached = 5,
        Auto_Defense_Crossed = 6,
        Auto_Low_Goal_Scored = 7,
        Auto_High_Goal_Scored = 8,
        Auto_Starting_Position = 9,
        Auto_Ending_Position = 10,
        Tele_Portcullis = 11,
        Tele_Fries = 12,
        Tele_Rampart = 13,
        Tele_Moat = 14,
        Tele_Drawbridge = 15,
        Tele_SallyPort = 16,
        Tele_RockWall = 17,
        Tele_RoughTerrain = 18,
        Tele_LowBar = 19,
        Tele_Low_Goal_Scored = 20,
        Tele_High_Goal_Scored = 21,
        Tele_Missed_High_Goal = 22,
        Robot_Disabled = 23,
        Time_Disabled = 24,
        End_Challenged = 25,
        End_Scaled = 26,
        Notes = 27

    };

    public partial class Form1 : Form
    {
        private const String _DATA_DIRECTORY = "C:\\2016_FRC_Scouting\\";
        private const String _EXCEL_FILENAME = "2016_FRC_Scouting_Data.xlsx";
        private const String dataSheet = "Scouting_Data";
        private Microsoft.Office.Interop.Excel.Application _xlApp;
        private Workbook _xlwb;
        private Worksheet _xlws;
        private object misValue = System.Reflection.Missing.Value;
        private Stopwatch timer = new Stopwatch();
        private Boolean _datagridInit = false;

        public Form1()
        {
            InitializeComponent();
        }

        #region Team_Properties
        internal int _Team_Num { get; set; }
        internal int _Match_Num { get; set; }
        internal String _Scout_Name { get; set; }
        internal String _Team_Alliance { get; set; }
        internal int _Auto_Defense_Reached { get; set; }
        internal int _Auto_Defense_Crossed { get; set; }
        internal int _Auto_Low_Goal_Scored { get; set; }
        internal int _Auto_High_Goal_Scored { get; set; }
        internal String _Auto_Starting_Position { get; set; }
        internal String _Auto_Ending_Position { get; set; }
        internal int _Tele_Portcullis { get; set; }
        internal int _Tele_Fries { get; set; }
        internal int _Tele_Rampart { get; set; }
        internal int _Tele_Moat { get; set; }
        internal int _Tele_Drawbridge { get; set; }
        internal int _Tele_SallyPort { get; set; }
        internal int _Tele_RockWall { get; set; }
        internal int _Tele_RoughTerrain { get; set; }
        internal int _Tele_LowBar { get; set; }
        internal int _Tele_Low_Goal_Scored { get; set; }
        internal int _Tele_High_Goal_Scored { get; set; }
        internal int _Tele_High_Goal_Missed { get; set; }
        internal int  _Robot_Disabled { get; set; }
        internal string _Time_Disabled { get; set; }
        internal int _End_Challenged { get; set; }
        internal int _End_Scaled { get; set; }
        internal String _Notes { get; set; }
        #endregion

        #region Search_Properties
        internal int _Total_High_Goals { get; set; }
        internal int _Total_High_Goals_Missed { get; set; }
        internal int _Total_Low_Goals { get; set; }
        internal int _Total_Auto_Crossing { get; set; }
        internal int _Total_Portcullis_Attempts { get; set; }
        internal int _Total_Rampart_Attempts { get; set; }
        internal int _Total_Drawbridge_Attempts { get; set; }
        internal int _Total_Freedom_Fries_Attempts { get; set; }
        internal int _Total_Moat_Attempts { get; set; }
        internal int _Total_SallyPort_Attempts { get; set; }
        internal int _Total_RockWall_Attempts { get; set; }
        internal int _Total_RoughTerrain_Attempts { get; set; }
        internal int _Total_LowBar_Attempts { get; set; }
        internal int _Total_Challenge_Attempts { get; set; }
        internal int _Total_Scale_Attempts { get; set; }
        #endregion


        private void btn_submitData_Click(object sender, EventArgs e)
        {
            Thread t;
            initializeProperties();
            if (initExcel()) //initialize Excel object
            {

                if (!verifyExistingDataFile())                  //Check if file already exists - File will be stored locally at C:\2016_FRC_Scouting\
                    createNewDataFile();                        //if not exist - create new file (force creation of file in directory where executable is run from)                
                openDataFile();                                 //access file if not open (if file not exist, will be created in condition above)
                gatherData();                                   //gather data from form
                if (_Team_Num.Equals(-1) || _Match_Num.Equals(-1))
                {
                    if (_Team_Num.Equals(-1))
                    {
                        t = new Thread(() => setStatusBar("Please specify a team number", 10000));
                        t.Start();
                        txt_teamNum.Focus();
                    }
                    else if (_Match_Num.Equals(-1))
                    {
                        t = new Thread(() => setStatusBar("Please specify the match number", 10000));
                        t.Start();
                        txt_matchNum.Focus();
                    }
                }
                else if (_Match_Num.ToString().Length >= 3)
                {
                    t = new Thread(() => setStatusBar("Match number too large, typo?", 10000));
                    t.Start();
                    txt_matchNum.Focus();
                }
                else
                {
                    addDataRow();
                    clearALLData();
                    t = new Thread(() => setStatusBar("Form submitted successfully"));
                    t.Start();
                    txt_teamNum.Focus();
                }
            }
        }

        private void setStatusBar(string msg, int timeToDisplay = 2500)
        {
            toolStripStatusLabel1.Text = msg;
            System.Threading.Thread.Sleep(2500);
            toolStripStatusLabel1.Text = "";
        }

        private void gatherData()
        {
            int result;
            try
            {
                _Team_Num = Int32.TryParse(txt_teamNum.Text, out result) ? result : -1;
                _Match_Num = Int32.TryParse(txt_matchNum.Text, out result) ? result : -1;
                _Scout_Name = txt_scoutName.Text;
                _Team_Alliance = rdo_allianceRed.Checked ? "Red" : "Blue";

                _Auto_Defense_Reached = rdo_reached.Checked ? 1 : 0;
                _Auto_Defense_Crossed = rdo_crossed.Checked ? 1 : 0;
                _Auto_Low_Goal_Scored = chk_lowScore.Checked ? 1 : 0;
                _Auto_High_Goal_Scored = chk_highScore.Checked ? 1 : 0;
                _Auto_Starting_Position = rdo_startNeutral.Checked ? "Neutral Zone" : "Courtyard";
                _Auto_Ending_Position = rdo_endNeutral.Checked ? "Neutral Zone" : "Courtyard";

                _Tele_Portcullis = Int32.TryParse(txt_portcullis.Text, out result) ? result : 0;
                _Tele_Fries = Int32.TryParse(txt_fries.Text, out result) ? result : 0;
                _Tele_Rampart = Int32.TryParse(txt_rampart.Text, out result) ? result : 0;
                _Tele_Moat = Int32.TryParse(txt_moat.Text, out result) ? result : 0;
                _Tele_Drawbridge = Int32.TryParse(txt_drawbridge.Text, out result) ? result : 0;
                _Tele_SallyPort = Int32.TryParse(txt_sallyPort.Text, out result) ? result : 0;
                _Tele_RockWall = Int32.TryParse(txt_rockWall.Text, out result) ? result : 0;
                _Tele_RoughTerrain = Int32.TryParse(txt_roughTerrain.Text, out result) ? result : 0;
                _Tele_LowBar = Int32.TryParse(txt_lowBar.Text, out result) ? result : 0;
                _Tele_Low_Goal_Scored = Int32.TryParse(txt_lowGoalsScored.Text, out result) ? result : 0;
                _Tele_High_Goal_Scored = Int32.TryParse(txt_highGoalsScored.Text, out result) ? result : 0;
                _Tele_High_Goal_Scored = Int32.TryParse(txt_highGoalsScored.Text, out result) ? result : 0;

                _Robot_Disabled = chk_robotDisabled.Checked ? 1 : 0;
                _Time_Disabled = mtb_timeDisabled.Enabled ? mtb_timeDisabled.Text : "N/A";
                _End_Challenged = rdo_Challenged.Checked ? 1 : 0;
                _End_Scaled = rdo_Scaled.Checked ? 1 : 0;
                _Notes = rtb_Notes.Text;
            }
            catch (Exception ex)
            {
                HandleException(ex, "Issue reading data from form");
            }
        }

        private void initializeProperties()
        {
            _Team_Num = 0;
            _Match_Num = 0;
            _Scout_Name = "";
            _Team_Alliance = "";

            _Auto_Defense_Reached = 0;
            _Auto_Defense_Crossed = 0;
            _Auto_Low_Goal_Scored = 0;
            _Auto_High_Goal_Scored = 0;
            _Auto_Starting_Position = "";
            _Auto_Ending_Position = "";

            _Tele_Portcullis = 0;
            _Tele_Fries = 0;
            _Tele_Rampart = 0;
            _Tele_Moat = 0;
            _Tele_Drawbridge = 0;
            _Tele_SallyPort = 0;
            _Tele_RockWall = 0;
            _Tele_RoughTerrain = 0;
            _Tele_LowBar = 0;
            _Tele_Low_Goal_Scored = 0;
            _Tele_High_Goal_Scored = 0;

            _Robot_Disabled = 0;
            _Time_Disabled = "";
            _End_Challenged = 0;
            _End_Scaled = 0;
            _Notes = "";
        }

        private void addDataRow()
        {
            try
            {
                //get last data row
                Range last = _xlws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
                int lastRow = last.Row;
                lastRow++;
                _xlws.Cells[lastRow, DATA_ROWS.Team_Num] = _Team_Num;
                _xlws.Cells[lastRow, DATA_ROWS.Match_Num] = _Match_Num;
                _xlws.Cells[lastRow, DATA_ROWS.Scout_Name] = _Scout_Name;
                _xlws.Cells[lastRow, DATA_ROWS.Alliance] = _Team_Alliance;
                _xlws.Cells[lastRow, DATA_ROWS.Auto_Defense_Reached] = _Auto_Defense_Reached;
                _xlws.Cells[lastRow, DATA_ROWS.Auto_Defense_Crossed] = _Auto_Defense_Crossed;
                _xlws.Cells[lastRow, DATA_ROWS.Auto_Low_Goal_Scored] = _Auto_Low_Goal_Scored;
                _xlws.Cells[lastRow, DATA_ROWS.Auto_High_Goal_Scored] = _Auto_High_Goal_Scored;
                _xlws.Cells[lastRow, DATA_ROWS.Auto_Starting_Position] = _Auto_Starting_Position;
                _xlws.Cells[lastRow, DATA_ROWS.Auto_Ending_Position] = _Auto_Ending_Position;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Portcullis] = _Tele_Portcullis;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Fries] = _Tele_Fries;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Rampart] = _Tele_Rampart;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Moat] = _Tele_Moat;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Drawbridge] = _Tele_Drawbridge;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_SallyPort] = _Tele_SallyPort;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_RockWall] = _Tele_RockWall;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_RoughTerrain] = _Tele_RoughTerrain;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_LowBar] = _Tele_LowBar;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Low_Goal_Scored] = _Tele_Low_Goal_Scored;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_High_Goal_Scored] = _Tele_High_Goal_Scored;
                _xlws.Cells[lastRow, DATA_ROWS.Tele_Missed_High_Goal] = _Tele_High_Goal_Missed;
                _xlws.Cells[lastRow, DATA_ROWS.Robot_Disabled] = _Robot_Disabled;
                _xlws.Cells[lastRow, DATA_ROWS.Time_Disabled] = _Time_Disabled;
                _xlws.Cells[lastRow, DATA_ROWS.End_Challenged] = _End_Challenged;
                _xlws.Cells[lastRow, DATA_ROWS.End_Scaled] = _End_Scaled;
                _xlws.Cells[lastRow, DATA_ROWS.Notes] = _Notes;


                _xlwb.Save();
                _xlwb.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Issue writing data row to worksheet");
            }

        }

        private void openDataFile()
        {
            try
            {
                _xlwb = _xlApp.Workbooks.Open(_DATA_DIRECTORY + _EXCEL_FILENAME);
                _xlws = _xlwb.Worksheets.get_Item(dataSheet);
            }
            catch (Exception ex)
            {
                HandleException(ex, "Issue opening the data file");
            }
        }

        private void clearExcelObjects()
        {
            releaseExcelObject(_xlws);
            releaseExcelObject(_xlwb);
            releaseExcelObject(_xlApp);
            _xlApp = null;
            _xlwb = null;
            _xlws = null;
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

                _xlws.Cells[1, DATA_ROWS.Team_Num] = "Team#";
                _xlws.Cells[1, DATA_ROWS.Match_Num] = "Match#";
                _xlws.Cells[1, DATA_ROWS.Scout_Name] = "Scout_Name";
                _xlws.Cells[1, DATA_ROWS.Alliance] = "Alliance";
                _xlws.Cells[1, DATA_ROWS.Auto_Defense_Reached] = "Auto_Defense_Reached";
                _xlws.Cells[1, DATA_ROWS.Auto_Defense_Crossed] = "Auto_Defense_Crossed";
                _xlws.Cells[1, DATA_ROWS.Auto_Low_Goal_Scored] = "Auto_Low_Goal_Scored";
                _xlws.Cells[1, DATA_ROWS.Auto_High_Goal_Scored] = "Auto_High_Goal_Scored";
                _xlws.Cells[1, DATA_ROWS.Auto_Starting_Position] = "Auto_Starting_Position";
                _xlws.Cells[1, DATA_ROWS.Auto_Ending_Position] = "Auto_Ending_Position";
                _xlws.Cells[1, DATA_ROWS.Tele_Portcullis] = "Tele_Portcullis";
                _xlws.Cells[1, DATA_ROWS.Tele_Fries] = "Tele_Fries";
                _xlws.Cells[1, DATA_ROWS.Tele_Rampart] = "Tele_Rampart";
                _xlws.Cells[1, DATA_ROWS.Tele_Moat] = "Tele_Moat";
                _xlws.Cells[1, DATA_ROWS.Tele_Drawbridge] = "Tele_Drawbridge";
                _xlws.Cells[1, DATA_ROWS.Tele_SallyPort] = "Tele_SallyPort";
                _xlws.Cells[1, DATA_ROWS.Tele_RockWall] = "Tele_RockWall";
                _xlws.Cells[1, DATA_ROWS.Tele_RoughTerrain] = "Tele_RoughTerrain";
                _xlws.Cells[1, DATA_ROWS.Tele_LowBar] = "Tele_LowBar";
                _xlws.Cells[1, DATA_ROWS.Tele_Low_Goal_Scored] = "Tele_Low_Goal_Scored";
                _xlws.Cells[1, DATA_ROWS.Tele_High_Goal_Scored] = "Tele_High_Goal_Scored";
                _xlws.Cells[1, DATA_ROWS.Tele_Missed_High_Goal] = "Tele_Missed_High_Goal";
                _xlws.Cells[1, DATA_ROWS.Robot_Disabled] = "Robot_Disabled";
                _xlws.Cells[1, DATA_ROWS.Time_Disabled] = "Time_Disabled";
                _xlws.Cells[1, DATA_ROWS.End_Challenged] = "End_Challenged";
                _xlws.Cells[1, DATA_ROWS.End_Scaled] = "End_Scaled";
                _xlws.Cells[1, DATA_ROWS.Notes] = "Notes";

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
            if (!Directory.Exists(_DATA_DIRECTORY)) //create data directory if it doesn't exist
            {
                Directory.CreateDirectory(_DATA_DIRECTORY);
            }
            else
            {
                if (File.Exists(_DATA_DIRECTORY + _EXCEL_FILENAME))
                    rtval = true;
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

        private void clearALLData()
        {
            clearAllCheckboxes(this);
            clearAllRadioButtons(this);
            clearAllTextBoxes(this);
            clearAllMaskedTextBoxes(this);
            clearAllRichTextBoxes(this);
        }

        private void btn_clearData_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Are you sure you want to clear the form?", "WHAT ARE YOU DOING!?!?!?!?!", MessageBoxButtons.YesNo);
            if (result1.Equals(DialogResult.Yes))
                clearALLData();
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

        private void clearAllMaskedTextBoxes(Control ctrl)
        {
            System.Windows.Forms.MaskedTextBox txt = ctrl as System.Windows.Forms.MaskedTextBox;
            if (txt == null)
            {
                foreach (Control child in ctrl.Controls)
                {
                    clearAllMaskedTextBoxes(child); //recursive
                }
            }
            else
            {
                txt.Text = String.Empty;
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
                rdoBtn.TabStop = true;
            }
        }

        private void clearAllCheckboxes(Control ctrl)
        {
            System.Windows.Forms.CheckBox chkBox = ctrl as System.Windows.Forms.CheckBox;
            if (chkBox == null)
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
            if (message.Equals(""))
                MessageBox.Show("Exception thrown:" + Environment.NewLine + ex.Message + Environment.NewLine + ex.InnerException + Environment.NewLine + ex.StackTrace);
            else
                MessageBox.Show(message + Environment.NewLine + ex.Message + Environment.NewLine + ex.InnerException + Environment.NewLine + ex.StackTrace);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            _datagridInit = false;
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            closeWorkbooks();
            if (_xlApp != null)
                clearExcelObjects();
        }

        private void closeWorkbooks()
        {
            try
            {
                _xlwb.Close(true);
            }
            catch (Exception ex)
            {
                return;
            }

        }

        private void chk_robotDisabled_CheckedChanged(object sender, EventArgs e)
        {
            mtb_timeDisabled.Enabled = chk_robotDisabled.Checked;
        }

        private void btn_clearSearch_Click(object sender, EventArgs e)
        {
            dgv_Search.Rows.Clear();
            clearTeamStats();
        }

        private void clearTeamStats()
        {
            txt_teamNum.Text = String.Empty;
            txt_totalAutoCrossing.Text = String.Empty;
            txt_totalHighGoals.Text = String.Empty;
            txt_totalHighGoalsMissed.Text = String.Empty;
            txt_totalLowGoals.Text = String.Empty;
            txt_totalChallenge.Text = String.Empty;
            txt_totalScales.Text = String.Empty;
            txt_totalPortcullis.Text = String.Empty;
            txt_totalRampart.Text = String.Empty;
            txt_totalDrawbridge.Text = String.Empty;
            txt_totalFreedomFries.Text = String.Empty;
            txt_totalMoat.Text = String.Empty;
            txt_totalSallyPort.Text = String.Empty;
            txt_totalRockWall.Text = String.Empty;
            txt_totalRoughTerrain.Text = String.Empty;
            txt_totalLowBar.Text = String.Empty;
        }

        private void btn_Search_Click(object sender, EventArgs e)
        {
            int teamResult;

            Thread t;
            if (_xlApp == null)
                initExcel();
            if (_xlwb == null || _xlws == null)
                openDataFile();
            else
            {
                resetExcel();
            }

            initializeProperties();     //initialize properties so they are empty when we use them    


            //what happens if we have been entering data and then want to quickly search a team? - Exception thrown @ GetCells()
            //need to save and close file then re-open so it applies lastCell to worksheet?


            if (Int32.TryParse(txt_teamNum.Text, out teamResult))
                _Team_Num = teamResult;
            if (_Team_Num.Equals(0))
            {
                t = new Thread(() => setStatusBar("Please specify a team number", 10000));
                t.Start();
                return;
            }
            else
            {
                dgv_Search.Rows.Clear();
                dgv_Search.Columns.Clear();
                initializeSearchDatagridView();
            }


            //get maximum search range
            Range last = _xlws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int lastRow = last.Row;
            Range data = _xlApp.get_Range("A1", String.Format("A" + lastRow.ToString()));
            //get rows that have the data we want
            ArrayList rows = searchData(data);

            initializeRobotStats();     //set search/stats properties to 0
            gatherSearchData(rows);     //gather data from rows and store in array
            displaySearchData();        //display data on datagridview
            t = new Thread(() => setStatusBar("Search completed successfully", 10000));
            t.Start();



        }

        private void resetExcel()
        {
            closeWorkbooks();
            clearExcelObjects();
            initExcel();
            openDataFile();
        }

        private void initializeRobotStats()
        {
            _Total_High_Goals = 0;
            _Total_High_Goals_Missed = 0;
            _Total_Low_Goals = 0;
            _Total_Auto_Crossing = 0;
            _Total_Scale_Attempts = 0;
            _Total_Portcullis_Attempts = 0;
            _Total_Rampart_Attempts = 0;
            _Total_Drawbridge_Attempts = 0;
            _Total_Freedom_Fries_Attempts = 0;
            _Total_Moat_Attempts = 0;
            _Total_SallyPort_Attempts = 0;
            _Total_RockWall_Attempts = 0;
            _Total_RoughTerrain_Attempts = 0;
            _Total_LowBar_Attempts = 0;
            _Total_Challenge_Attempts = 0;
            _Total_Scale_Attempts = 0;
        }

        private void displaySearchData()
        {
            txt_totalAutoCrossing.Text = _Total_Auto_Crossing.ToString();
            txt_totalHighGoals.Text = _Total_High_Goals.ToString();
            txt_totalHighGoalsMissed.Text = _Total_High_Goals_Missed.ToString();
            txt_totalLowGoals.Text = _Total_Low_Goals.ToString();
            txt_totalChallenge.Text = _Total_Challenge_Attempts.ToString();
            txt_totalScales.Text = _Total_Scale_Attempts.ToString();
            txt_totalPortcullis.Text = _Total_Portcullis_Attempts.ToString();
            txt_totalRampart.Text = _Total_Rampart_Attempts.ToString();
            txt_totalDrawbridge.Text = _Total_Drawbridge_Attempts.ToString();
            txt_totalFreedomFries.Text = _Total_Freedom_Fries_Attempts.ToString();
            txt_totalMoat.Text = _Total_Moat_Attempts.ToString();
            txt_totalSallyPort.Text = _Total_SallyPort_Attempts.ToString();
            txt_totalRockWall.Text = _Total_RockWall_Attempts.ToString();
            txt_totalRoughTerrain.Text = _Total_RoughTerrain_Attempts.ToString();
            txt_totalLowBar.Text = _Total_LowBar_Attempts.ToString();
        }

        private void gatherSearchData(ArrayList rows, Boolean displayToGrid = true)
        {
            foreach (int item in rows)
            {
                Range range1 = _xlws.Rows[item]; //For all columns in row selected
                _Team_Num = (int)_xlws.Cells[item, DATA_ROWS.Team_Num].Value2;
                _Match_Num = (int)_xlws.Cells[item, DATA_ROWS.Match_Num].Value2;
                _Scout_Name = (String)_xlws.Cells[item, DATA_ROWS.Scout_Name].Value2;
                _Team_Alliance = (String)_xlws.Cells[item, DATA_ROWS.Alliance].Value2;
                _Auto_Defense_Reached = (int)_xlws.Cells[item, DATA_ROWS.Auto_Defense_Reached].Value2;
                _Auto_Defense_Crossed = (int)_xlws.Cells[item, DATA_ROWS.Auto_Defense_Crossed].Value2;
                if (_Auto_Defense_Crossed != 0)
                    _Total_Auto_Crossing++;
                _Auto_Low_Goal_Scored = (int)_xlws.Cells[item, DATA_ROWS.Auto_Low_Goal_Scored].Value2;
                _Auto_High_Goal_Scored = (int)_xlws.Cells[item, DATA_ROWS.Auto_High_Goal_Scored].Value2;
                _Auto_Starting_Position = (String)_xlws.Cells[item, DATA_ROWS.Auto_Starting_Position].Value2;
                _Auto_Ending_Position = (String)_xlws.Cells[item, DATA_ROWS.Auto_Ending_Position].Value2;

                _Tele_Portcullis = (int)_xlws.Cells[item, DATA_ROWS.Tele_Portcullis].Value2;
                _Total_Portcullis_Attempts += _Tele_Portcullis;

                _Tele_Fries = (int)_xlws.Cells[item, DATA_ROWS.Tele_Fries].Value2;
                _Total_Freedom_Fries_Attempts += _Tele_Fries;

                _Tele_Rampart = (int)_xlws.Cells[item, DATA_ROWS.Tele_Rampart].Value2;
                _Total_Rampart_Attempts += _Tele_Rampart;

                _Tele_Moat = (int)_xlws.Cells[item, DATA_ROWS.Tele_Moat].Value2;
                _Total_Moat_Attempts += _Tele_Moat;

                _Tele_Drawbridge = (int)_xlws.Cells[item, DATA_ROWS.Tele_Drawbridge].Value2;
                _Total_Drawbridge_Attempts += _Tele_Drawbridge;

                _Tele_SallyPort = (int)_xlws.Cells[item, DATA_ROWS.Tele_SallyPort].Value2;
                _Total_SallyPort_Attempts += _Tele_SallyPort;

                _Tele_RockWall = (int)_xlws.Cells[item, DATA_ROWS.Tele_RockWall].Value2;
                _Total_RockWall_Attempts += _Tele_RockWall;

                _Tele_RoughTerrain = (int)_xlws.Cells[item, DATA_ROWS.Tele_RoughTerrain].Value2;
                _Total_RoughTerrain_Attempts += _Tele_RoughTerrain;

                _Tele_LowBar = (int)_xlws.Cells[item, DATA_ROWS.Tele_LowBar].Value2;
                _Total_LowBar_Attempts += _Tele_LowBar;

                _Tele_Low_Goal_Scored = (int)_xlws.Cells[item, DATA_ROWS.Tele_Low_Goal_Scored].Value2;
                _Total_Low_Goals += _Tele_Low_Goal_Scored;

                _Tele_High_Goal_Scored = (int)_xlws.Cells[item, DATA_ROWS.Tele_High_Goal_Scored].Value2;
                _Total_High_Goals += _Tele_High_Goal_Scored;

                _Tele_High_Goal_Missed = (int)_xlws.Cells[item, DATA_ROWS.Tele_Missed_High_Goal].Value2;
                _Total_High_Goals_Missed += _Tele_High_Goal_Missed;

                _Robot_Disabled = (int)_xlws.Cells[item, DATA_ROWS.Robot_Disabled].Value2;
                _Time_Disabled = (String)_xlws.Cells[item, DATA_ROWS.Time_Disabled].Text.ToString();
                _End_Challenged = (int)_xlws.Cells[item, DATA_ROWS.End_Challenged].Value2;
                if (_End_Challenged != 0)
                    _Total_Challenge_Attempts++;
                _End_Scaled = (int)_xlws.Cells[item, DATA_ROWS.End_Scaled].Value2;
                if (_End_Scaled != 0)
                    _Total_Scale_Attempts++;
                _Notes = (String)_xlws.Cells[item, DATA_ROWS.Notes].Value2;
                if (displayToGrid)
                    dgv_Search.Rows.Add(_Team_Num, _Match_Num, _Scout_Name, _Team_Alliance, _Auto_Defense_Reached, _Auto_Defense_Crossed, _Auto_Low_Goal_Scored, _Auto_High_Goal_Scored, _Auto_Starting_Position, _Auto_Ending_Position, _Tele_Portcullis, _Tele_Fries, _Tele_Rampart, _Tele_Moat, _Tele_Drawbridge, _Tele_SallyPort, _Tele_RockWall, _Tele_RoughTerrain, _Tele_LowBar, _Tele_Low_Goal_Scored, _Tele_High_Goal_Scored, _Tele_High_Goal_Missed, _Robot_Disabled, _Time_Disabled, _End_Challenged, _End_Scaled, _Notes);
            }
        }

        private void initializeSearchDatagridView()
        {
            dgv_Search.Columns.Add("Team_Num", "Team_Num");
            dgv_Search.Columns.Add("Match_Num", "Match_Num");
            dgv_Search.Columns.Add("Scout_Name", "Scout_Name");
            dgv_Search.Columns.Add("Team_Alliance", "Team_Alliance");
            dgv_Search.Columns.Add("Auto_Defense_Reached", "Auto_Defense_Reached");
            dgv_Search.Columns.Add("Auto_Defense_Crossed", "Auto_Defense_Crossed");
            dgv_Search.Columns.Add("Auto_Low_Goal_Scored", "Auto_Low_Goal_Scored");
            dgv_Search.Columns.Add("Auto_High_Goal_Scored", "Auto_High_Goal_Scored");
            dgv_Search.Columns.Add("Auto_Starting_Position", "Auto_Starting_Position");
            dgv_Search.Columns.Add("Auto_Ending_Position", "Auto_Ending_Position");
            dgv_Search.Columns.Add("Tele_Portcullis", "Tele_Portcullis");
            dgv_Search.Columns.Add("Tele_Fries", "Tele_Fries");
            dgv_Search.Columns.Add("Tele_Rampart", "Tele_Rampart");
            dgv_Search.Columns.Add("Tele_Moat", "Tele_Moat");
            dgv_Search.Columns.Add("Tele_Drawbridge", "Tele_Drawbridge");
            dgv_Search.Columns.Add("Tele_SallyPort", "Tele_SallyPort");
            dgv_Search.Columns.Add("Tele_RockWall", "Tele_RockWall");
            dgv_Search.Columns.Add("Tele_RoughTerrain", "Tele_RoughTerrain");
            dgv_Search.Columns.Add("Tele_LowBar", "Tele_LowBar");
            dgv_Search.Columns.Add("Tele_Low_Goal_Scored", "Tele_Low_Goal_Scored");
            dgv_Search.Columns.Add("Tele_High_Goal_Scored", "Tele_High_Goal_Scored");
            dgv_Search.Columns.Add("Tele_Missed_High_Goal", "Tele_Missed_High_Goal");
            dgv_Search.Columns.Add("Robot_Disabled", "Robot_Disabled");
            dgv_Search.Columns.Add("Time_Disabled", "Time_Disabled");
            dgv_Search.Columns.Add("End_Challenged", "End_Challenged");
            dgv_Search.Columns.Add("End_Scaled", "End_Scaled");
            dgv_Search.Columns.Add("Notes", "Notes");
            _datagridInit = true;
        }

        private void initializeDataGridAggregateView()
        {
            dgv_Search.Columns.Add("Team_Num", "Team_Num");
            dgv_Search.Columns.Add("Total_High_Goals", "Total_High_Goals");
            dgv_Search.Columns.Add("Total_High_Goals_Missed", "Total_High_Goals_Missed");
            dgv_Search.Columns.Add("Total_Low_Goals", "Total_Low_Goals");
            dgv_Search.Columns.Add("Total_Auto_Crossing", "Total_Auto_Crossing");
            dgv_Search.Columns.Add("Total_Portcullis_Attempts", "Total_Portcullis_Attempts");
            dgv_Search.Columns.Add("Total_Rampart_Attempts", "Total_Rampart_Attempts");
            dgv_Search.Columns.Add("Total_Drawbridge_Attempts", "Total_Drawbridge_Attempts");
            dgv_Search.Columns.Add("Total_Freedom_Fries_Attempts", "Total_Freedom_Fries_Attempts");
            dgv_Search.Columns.Add("Total_Moat_Attempts", "Total_Moat_Attempts");
            dgv_Search.Columns.Add("Total_SallyPort_Attempts", "Total_SallyPort_Attempts");
            dgv_Search.Columns.Add("Total_RockWall_Attempts", "Total_RockWall_Attempts");
            dgv_Search.Columns.Add("Total_RoughTerrain_Attempts", "Total_RoughTerrain_Attempts");
            dgv_Search.Columns.Add("Total_LowBar_Attempts", "Total_LowBar_Attempts");
            dgv_Search.Columns.Add("Total_Scale_Attempts", "Total_Scale_Attempts");
            dgv_Search.Columns.Add("Total_Challenge_Attempts", "Total_Challenge_Attempts");
            _datagridInit = true;
        }

        private ArrayList searchData(Range excelData)
        {
            ArrayList searchArray = new ArrayList();
            Range currentFind = null;
            Range firstFind = null; //need to track this when we wrap around for search

            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            currentFind = excelData.Find(_Team_Num, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart, XlSearchOrder.xlByRows, XlSearchDirection.xlNext, false, Type.Missing, Type.Missing);

            while (currentFind != null)
            {
                //keep track of the first range you find. 
                if (firstFind == null)
                    firstFind = currentFind;

                //if you didn't move to a new range/find another row, you are done.
                else if (currentFind.get_Address(XlReferenceStyle.xlA1) == firstFind.get_Address(XlReferenceStyle.xlA1))
                    break;

                //Add found rows to array
                searchArray.Add(currentFind.Row);

                //find next row if available
                currentFind = excelData.FindNext(currentFind);
            }

            return searchArray;
        }

        private void btn_showTeamAggregate_Click(object sender, EventArgs e)
        {
            string foundTeams = "";
            ArrayList rows;

            //get all team numbers
            if (_xlApp == null)
                initExcel();
            if (_xlwb == null || _xlws == null)
                openDataFile();
            else
                resetExcel();

            dgv_Search.Rows.Clear();
            dgv_Search.Columns.Clear();
            clearTeamStats();
            initializeDataGridAggregateView();

            //get all team numbers from excel sheet
            Range last = _xlws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int lastRow = last.Row;
            Range data = _xlApp.get_Range("A2", String.Format("A" + lastRow.ToString()));

            int temp;
            //store all unique team numbers in string
            foreach (Range item in data.Cells)
            {
                temp = (int)item.Value;
                if (!foundTeams.Contains(temp.ToString()))
                    foundTeams += item.Value + ";";
            }

            //create <key,value> list with team number as key and value is array of columns (7) 
            String[] teams = foundTeams.Split(';');
            Array.Resize(ref teams, teams.Length - 1);

            //var dict = teams.ToDictionary(item => new int[7]);
            var dict = teams.ToDictionary(v => v, v => new int[] { 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 });

            /* integer array stored in value of dictionary is organized as such
             * Total High Goals, Total High Goals Missed, Total Low Goals, 
             * Total AUTO crossings, Total Crossings of...Portcullis, Rampart, 
             * Drawbridge, Fries, Moat, Sally Port, Rock Wall, Rough Terrain, 
             * Low Bar, Total scales, Total challenges
            */
            Dictionary<int, int[]> finalSet = new Dictionary<int, int[]>();
            foreach (var pair in dict)
            {
                _Team_Num = Int32.Parse(pair.Key);
                rows = searchData(data);
                gatherSearchData(rows, false);      //gather data from rows and store in array
                finalSet.Add(_Team_Num, new int[] {_Total_High_Goals, _Total_High_Goals_Missed, _Total_Low_Goals, _Total_Auto_Crossing, 
                _Total_Portcullis_Attempts, _Total_Rampart_Attempts, _Total_Drawbridge_Attempts, _Total_Freedom_Fries_Attempts, 
                _Total_Moat_Attempts, _Total_SallyPort_Attempts, _Total_RockWall_Attempts, _Total_RoughTerrain_Attempts, 
                _Total_LowBar_Attempts, _Total_Scale_Attempts, _Total_Challenge_Attempts});
                initializeRobotStats();             //set search/stats properties to 0
            }

            //display data to datagrid
            foreach (var pair in finalSet)
            {
                int teamNum = pair.Key;
                dgv_Search.Rows.Add(pair.Key, pair.Value[0], pair.Value[1], pair.Value[2], pair.Value[3], pair.Value[4], pair.Value[5], pair.Value[6], pair.Value[7], pair.Value[8], pair.Value[9], pair.Value[10], pair.Value[11], pair.Value[12], pair.Value[13], pair.Value[14]);
            }

        }

        private void txt_teamNum_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar.Equals(Convert.ToChar(Keys.Enter)))
            {
                if (tabControl1.SelectedTab.Text.Equals("Search"))
                    btn_Search_Click(sender, e);
            }
        }
    }
}
