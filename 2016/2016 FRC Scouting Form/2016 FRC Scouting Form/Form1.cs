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
        End_Challenged = 22,
        End_Scaled = 23,
        Notes = 24

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
        private DataGrid dataGridView1;

        internal int pTeam_Num;
        internal int pMatch_Num;
        internal string pScout_Name;
        internal string pAlliance;
        internal bool pAuto_Defense_Reached;
        internal bool pAuto_Defense_Crossed;
        internal bool pAuto_Low_Goal_Scored;
        internal bool pAuto_High_Goal_Scored;
        internal string pAuto_Starting_Position;
        internal string pAuto_Ending_Position;
        internal int pTele_Portcullis;
        internal int pTele_Fries;
        internal int pTele_Rampart;
        internal int pTele_Moat;
        internal int pTele_Drawbridge;
        internal int pTele_SallyPort;
        internal int pTele_RockWall;
        internal int pTele_RoughTerrain;
        internal int pTele_LowBar;
        internal int pTele_Low_Goal_Scored;
        internal int pTele_High_Goal_Scored;
        internal bool pEnd_Challenged;
        internal bool pEnd_Scaled;
        internal string pNotes;


        public Form1()
        {
            InitializeComponent();
        }

        internal int Team_Num
        {
            get {return pTeam_Num;}
            set {pTeam_Num = value;}
        }

        internal int Match_Num
        {
            get { return pMatch_Num; }
            set { pMatch_Num = value; }
        }

        internal String Scout_Name
        {
            get { return pScout_Name; }
            set { pScout_Name = value; }
        }

        internal String Team_Alliance
        {
            get { return pAlliance; }
            set { pAlliance = value; }
        }

        internal Boolean Auto_Defense_Reached
        {
            get { return pAuto_Defense_Reached; }
            set { pAuto_Defense_Reached = value; }
        }

        internal Boolean Auto_Defense_Crossed
        {
            get { return pAuto_Defense_Crossed; }
            set { pAuto_Defense_Crossed = value; }
        }

        internal Boolean Auto_Low_Goal_Scored
        {
            get { return pAuto_Low_Goal_Scored; }
            set { pAuto_Low_Goal_Scored = value; }
        }

        internal Boolean Auto_High_Goal_Scored
        {
            get { return pAuto_High_Goal_Scored; }
            set { pAuto_High_Goal_Scored = value; }
        }

        internal String Auto_Starting_Position
        {
            get { return pAuto_Starting_Position; }
            set { pAuto_Starting_Position = value; }
        }

        internal String Auto_Ending_Position
        {
            get { return pAuto_Ending_Position; }
            set { pAuto_Ending_Position = value; }
        }

        internal int Tele_Portcullis
        {
            get { return pTele_Portcullis; }
            set { pTele_Portcullis = value; }
        }

        internal int Tele_Fries
        {
            get { return pTele_Fries; }
            set { pTele_Fries = value; }
        }

        internal int Tele_Rampart
        {
            get { return pTele_Rampart; }
            set { pTele_Rampart = value; }
        }

        internal int Tele_Moat
        {
            get { return pTele_Moat; }
            set { pTele_Moat = value; }
        }

        internal int Tele_Drawbridge
        {
            get { return pTele_Drawbridge; }
            set { pTele_Drawbridge = value; }
        }

        internal int Tele_SallyPort
        {
            get { return pTele_SallyPort; }
            set { pTele_SallyPort = value; }
        }

        internal int Tele_RockWall
        {
            get { return pTele_RockWall; }
            set { pTele_RockWall = value; }
        }

        internal int Tele_RoughTerrain
        {
            get { return pTele_RoughTerrain; }
            set { pTele_RoughTerrain = value; }
        }

        internal int Tele_LowBar
        {
            get { return pTele_LowBar; }
            set { pTele_LowBar = value; }
        }

        internal int Tele_Low_Goal_Scored
        {
            get { return pTele_Low_Goal_Scored; }
            set { pTele_Low_Goal_Scored = value; }
        }

        internal int Tele_High_Goal_Scored
        {
            get { return pTele_High_Goal_Scored; }
            set { pTele_High_Goal_Scored = value; }
        }

        internal bool End_Challenged
        {
            get { return pEnd_Challenged; }
            set { pEnd_Challenged = value; }
        }

        internal bool End_Scaled
        {
            get { return pEnd_Scaled; }
            set { pEnd_Scaled = value; }
        }

        internal String Notes
        {
            get { return pNotes; }
            set { pNotes = value; }
        }

        private void btn_submitData_Click(object sender, EventArgs e)
        {
            initializeProperties();
            if (initExcel()) //initialize Excel object
            {
                
                if (!verifyExistingDataFile())                  //Check if file already exists - File will be stored locally at C:\2016_FRC_Scouting\
                    createNewDataFile();                        //if not exist - create new file (force creation of file in directory where executable is run from)                
                openDataFile();                                 //access file if not open (if file not exist, will be created in condition above)
                gatherData();                                   //gather data from form
                addDataRow();
               
                //dataGridView1.DataSource = gatherData();  
                //set into specific format
            }
        }

        private void gatherData()
        {
            Team_Num = Int32.Parse(txt_teamNum.Text);
            Match_Num = Int32.Parse(txt_matchNum.Text);
            Scout_Name = txt_scoutName.Text;
            Team_Alliance = rdo_allianceRed.Checked ? "Red" : "Blue";
            Auto_Defense_Reached = chk_reached.Checked;
            Auto_Defense_Crossed = chk_crossed.Checked;
            Auto_Low_Goal_Scored = chk_lowScore.Checked;
            Auto_High_Goal_Scored = chk_highScore.Checked;
            Auto_Starting_Position = rdo_startNeutral.Checked ? "Neutral Zone" : "Courtyard";
            Auto_Ending_Position = rdo_endNeutral.Checked ? "Neutral Zone" : "Courtyard";
            Tele_Portcullis = Int32.Parse(txt_portcullis.Text);
            Tele_Fries = Int32.Parse(txt_fries.Text);
            Tele_Rampart = Int32.Parse(txt_rampart.Text);
            Tele_Moat = Int32.Parse(txt_moat.Text);
            Tele_Drawbridge = Int32.Parse(txt_drawbridge.Text);
            Tele_SallyPort = Int32.Parse(txt_sallyPort.Text);
            Tele_RockWall = Int32.Parse(txt_rockWall.Text);
            Tele_RoughTerrain = Int32.Parse(txt_roughTerrain.Text);
            Tele_LowBar = Int32.Parse(txt_lowBar.Text);
            Tele_Low_Goal_Scored = Int32.Parse(txt_lowGoalsScored.Text);
            Tele_High_Goal_Scored = Int32.Parse(txt_highGoalsScored.Text);
            End_Challenged = rdo_Challenged.Checked;
            End_Scaled = rdo_Scaled.Checked;
            Notes = rtb_Notes.Text;            
        }

        private void initializeProperties()
        {
            Team_Num = 0;
            Match_Num = 0;
            Scout_Name = "";
            Team_Alliance = "";
            Auto_Defense_Reached = false;
            Auto_Defense_Crossed = false;
            Auto_Low_Goal_Scored = false;
            Auto_High_Goal_Scored = false;
            Auto_Starting_Position = "";
            Auto_Ending_Position = "";
            Tele_Portcullis = 0;
            Tele_Fries = 0;
            Tele_Rampart = 0;
            Tele_Moat = 0;
            Tele_Drawbridge = 0;
            Tele_SallyPort = 0;
            Tele_RockWall = 0;
            Tele_RoughTerrain = 0;
            Tele_LowBar = 0;
            Tele_Low_Goal_Scored = 0;
            Tele_High_Goal_Scored = 0;
            End_Challenged = false;
            End_Scaled = false;
            Notes = "";
        }

        private void addDataRow()
        {
            //get last data row
            Range last = _xlws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);
            int lastRow = last.Row;
            lastRow++;
            _xlws.Cells[lastRow, DATA_ROWS.Team_Num] = Team_Num;
            _xlws.Cells[lastRow, DATA_ROWS.Match_Num] = Match_Num;
            _xlws.Cells[lastRow, DATA_ROWS.Scout_Name] = Scout_Name;
            _xlws.Cells[lastRow, DATA_ROWS.Alliance] = Team_Alliance;
            _xlws.Cells[lastRow, DATA_ROWS.Auto_Defense_Reached] = Auto_Defense_Reached;
            _xlws.Cells[lastRow, DATA_ROWS.Auto_Defense_Crossed] = Auto_Defense_Crossed;
            _xlws.Cells[lastRow, DATA_ROWS.Auto_Low_Goal_Scored] = Auto_Low_Goal_Scored;
            _xlws.Cells[lastRow, DATA_ROWS.Auto_High_Goal_Scored] = Auto_High_Goal_Scored;
            _xlws.Cells[lastRow, DATA_ROWS.Auto_Starting_Position] = Auto_Starting_Position;
            _xlws.Cells[lastRow, DATA_ROWS.Auto_Ending_Position] = Auto_Ending_Position;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_Portcullis] = Tele_Portcullis;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_Fries] = Tele_Fries;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_Rampart] = Tele_Rampart;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_Moat] = Tele_Moat;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_Drawbridge] = Tele_Drawbridge;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_SallyPort] = Tele_SallyPort;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_RockWall] = Tele_RockWall;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_RoughTerrain] = Tele_RoughTerrain;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_LowBar] = Tele_LowBar;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_Low_Goal_Scored] = Tele_Low_Goal_Scored;
            _xlws.Cells[lastRow, DATA_ROWS.Tele_High_Goal_Scored] = Tele_High_Goal_Scored;
            _xlws.Cells[lastRow, DATA_ROWS.End_Challenged] = End_Challenged;
            _xlws.Cells[lastRow, DATA_ROWS.End_Scaled] = End_Scaled;
            _xlws.Cells[lastRow, DATA_ROWS.Notes] = Notes;

            
            _xlwb.Save();
            _xlwb.Close(true, misValue, misValue);

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
            if(!Directory.Exists(_DATA_DIRECTORY)) //create data directory if it doesn't exist
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

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            clearExcelObjects();
        }
    }
}
