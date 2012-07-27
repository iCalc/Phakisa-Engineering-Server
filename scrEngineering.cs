using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using Microsoft.VisualBasic;
using System.IO;
using Analysis = clsAnalysis;
using TB = clsTable;
using DB = clsDBase;
using Base = clsMain;
using General = clsGeneral;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Data.OleDb;
using MetaReportRuntime;
using ICSharpCode.SharpZipLib.Checksums;
using ICSharpCode.SharpZipLib.Zip;

namespace Phakisa
{
    public partial class scrEngineering : Form
    {
        #region Declarations
        int columnnr = 0;
        int intNoOfDays = 0;
        int noOFDay = 0;
        DateTime sheetfhs = new DateTime();
        DateTime sheetlhs = new DateTime();
        int importdone = 0;
        DataTable fixShifts = new DataTable();
        int intStartDay = 0;
        int intEndDay = 0;
        int intStopDay = 0;
        int workedShiftsFixedClockedShift = 0;
        int exitValue = 0;
        string searchEmplNr = string.Empty;
        string searchEmplName = string.Empty;
        string searchEmplGang = string.Empty;
        string Path = string.Empty;
        string strWherePeriod = string.Empty;

        clsBL.clsBL BusinessLanguage = new clsBL.clsBL();
        clsTable.clsTable TB = new clsTable.clsTable();
        clsGeneral.clsGeneral General = new clsGeneral.clsGeneral();
        clsShared Shared = new clsShared();
        clsTableFormulas TBFormulas = new clsTableFormulas();
        clsMain.clsMain Base = new clsMain.clsMain();
        clsAnalysis.clsAnalysis Analysis = new clsAnalysis.clsAnalysis();
        SqlConnection myConn = new SqlConnection();
        SqlConnection AConn = new SqlConnection();
        SqlConnection AAConn = new SqlConnection();
        SqlConnection BaseConn = new SqlConnection();
        System.Collections.Hashtable buttonCollection = new System.Collections.Hashtable();

        Dictionary<string, string> dict = new Dictionary<string, string>();
        Dictionary<string, string> GangTypes = new Dictionary<string, string>();
        Dictionary<string, string> ParameterNames = new Dictionary<string, string>();
        Dictionary<string, string> Employeetypes = new Dictionary<string, string>();
        Dictionary<string, string> Wagecodetypes = new Dictionary<string, string>();
        Dictionary<string, string> dictPrimaryKeyValues = new Dictionary<string, string>();
        Dictionary<string, string> dictGridValues = new Dictionary<string, string>();

        bool blTablenames = true;
        string strEarningsCode = string.Empty;
        string strprevPeriod = string.Empty;
        string prevDatabaseName = string.Empty;
        string strWhere = string.Empty;
        string strWhereSection = string.Empty;
        string strActivity = string.Empty;
        string strMiningIndicator = string.Empty;
        string strMO = string.Empty;
        string strServerPath = string.Empty;
        string strName = string.Empty;
        string strWagecodes = string.Empty;
        string strMetaReportCode = "BSFnupmWkNxm8ZAA1ZhlOgL8fNdMdg4zhJj/j6T0vEyG9aSzk/HPwYcrjmawRGou66hBtseT7qJE+9hbEq9jces6bcGJmtz4Ih8Fic4UIw0Kt2lEffc05nFdiD2aQC0m";

        string dbPath = string.Empty;

        string[] ClockedShifts = new string[5];
        string[] OffShifts = new string[5];
        int intFiller = 0;
        int intCounter = 0;

        List<string> lstGangs = new List<string>();
        List<string> lstParticipGangs = new List<string>();
        List<string> lstNames = new List<string>();
        List<string> lstNewForemen = new List<string>();
        List<string> lstForemen = new List<string>();
        List<string> lstPrimaryKeyColumns = new List<string>();
        List<string> lstColumnNames = new List<string>();
        List<string> lstTableColumns = new List<string>();

        Int64 intProcessCounter = 0;
        StringBuilder strSqlAlter = new StringBuilder();

        DataTable Labour = new DataTable();
        DataTable KPFCostLevel = new DataTable();
        DataTable Designations = new DataTable();
        DataTable SubsectionDept = new DataTable();
        DataTable Clocked = new DataTable();
        DataTable Rates = new DataTable();
        DataTable EmplPen = new DataTable();
        DataTable Configs = new DataTable();
        DataTable Participation = new DataTable();
        DataTable MineParameters = new DataTable();
        DataTable DeptParameters = new DataTable();
        DataTable KPF = new DataTable();
        DataTable HOD = new DataTable();
        DataTable Artisans = new DataTable();
        DataTable Officials = new DataTable();
        DataTable newDataTable = new DataTable();
        DataTable _formulas = new DataTable();

        DataTable Monitor = new DataTable();
        DataTable Calendar = new DataTable();
        DataTable Production = new DataTable();
        DataTable earningsCodes = new DataTable();
        DataTable Status = new DataTable();
        DataTable BonusShifts = new DataTable();

        private ExcelDataReader.ExcelDataReader spreadsheet = null;

        ToolTip tooltip = new ToolTip();
        #endregion

        public scrEngineering()
        {
            InitializeComponent();

        }

        internal void scrEngineeringLoad(string Period, string Region, string BussUnit, string Userid, string MiningType, string BonusType, string Environment)
        {
            #region disable all functions
            //Disable all menu functions.
            foreach (ToolStripMenuItem IT in menuStrip1.Items)
            {
                if (IT.DropDownItems.Count > 0)
                {
                    foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                    {
                        if (ITT.DropDownItems.Count > 0)
                        {
                            foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                            {
                                ITTT.Enabled = false;
                            }
                        }
                        else
                        {
                            ITT.Enabled = false;
                        }
                    }
                }
                else
                {
                    IT.Enabled = false;
                }
            }
            #endregion

            #region declarations
            BusinessLanguage.Period = Period;
            BusinessLanguage.Region = Region;
            BusinessLanguage.BussUnit = BussUnit;
            BusinessLanguage.Userid = Userid;
            BusinessLanguage.MiningType = MiningType;
            BusinessLanguage.BonusType = BonusType;
            txtMiningType.Text = MiningType;
            txtBonusType.Text = BonusType;
            strServerPath = Environment;
            txtDatabaseName.Text = "ENGSER2000";
            //Display dbname in text box
            txtDatabaseName.Text = txtDatabaseName.Text.Trim();
            Base.DBName = txtDatabaseName.Text.Trim();
            Base.Period = BusinessLanguage.Period;

            //Setup the environment BEFORE the databases are moved to the classes.  This is because the environment path forms
            //part of the fisical name of the db

            setEnvironment();

            Base.DBName = txtDatabaseName.Text.Trim();
            TB.DBName = txtDatabaseName.Text.Trim();

            #endregion

            #region Connections
            //Open Connections and create classes

            AAConn = Analysis.AnalysisConnection;
            AAConn.Open();
            BaseConn = Base.BaseConnection;
            BaseConn.Open();

            #endregion

            DataTable useraccess = Base.SelectAccessByUserid(BusinessLanguage.Userid, Base.BaseConnectionString);

            #region Assign useraccess

            //BusinessLanguage.BussUnit = useraccess.Rows[0]["BUSSUNIT"].ToString().Trim();
            BusinessLanguage.Resp = useraccess.Rows[0]["RESP"].ToString().Trim();

            foreach (DataRow dr in useraccess.Rows)
            {
                string strCodeName = dr[6].ToString().Trim();
                foreach (ToolStripMenuItem IT in menuStrip1.Items)
                {
                    if (IT.DropDownItems.Count > 0)
                    {
                        foreach (ToolStripMenuItem ITT in IT.DropDownItems)
                        {
                            if (ITT.DropDownItems.Count > 0)
                            {
                                foreach (ToolStripMenuItem ITTT in ITT.DropDownItems)
                                {
                                    if (ITTT.Name.Trim() == strCodeName)
                                    {
                                        ITTT.Enabled = true;
                                    }
                                }
                            }
                            else
                                if (ITT.Name.Trim() == strCodeName)
                                {
                                    ITT.Enabled = true;
                                }
                        }
                    }
                    else
                    {
                        if (IT.Name.Trim() == strCodeName)
                        {
                            IT.Enabled = true;
                        }

                    }
                }

            }
            #endregion

            #region General
            //Display user details
            txtUserDetails.Text = BusinessLanguage.Userid + " - " + BusinessLanguage.Region + " - " + BusinessLanguage.BussUnit;
            //txtDatabaseName.Text = BusinessLanguage.BussUnit;

            txtPeriod.Text = BusinessLanguage.Period;

            // Set up the delays for the ToolTip.
            tooltip.AutoPopDelay = 5000;
            tooltip.InitialDelay = 1000;
            tooltip.ReshowDelay = 500;
            //Force the ToolTip text to be displayed whether or not the form is active.
            tooltip.ShowAlways = true;

            //Set up the ToolTip text for the Button and Checkbox.
            tooltip.SetToolTip(this.btnImportADTeam, "Clocked Shifts");
            tooltip.SetToolTip(this.tabLabour, "Bonus Shifts");
            tooltip.SetToolTip(this.btnSearch, "Search");

            listBox2.Enabled = false;
            listBox3.Enabled = false;


            #endregion

            #region Status button collection

            //Add the buttons needed for this bonus scheme and that are on the STATUS tab.
            buttonCollection["tabCalendar"] = btnLockCalendar;
            buttonCollection["tabLabour"] = btnLockBonusShifts;
            buttonCollection["tabEmplPen"] = btnLockEmplPen;
            buttonCollection["tabHOD"] = btnLockHOD;
            buttonCollection["Bonus Report Process - Phase 1"] = btnBonusPrints;
            buttonCollection["Input Process"] = btnInputProcess;
            buttonCollection["tabSubSectionDept"] = btnLockSubsectionDept;
            buttonCollection["tabDeptParameters"] = btnLockDeptParameters;
            buttonCollection["tabKPFCostLevel"] = btnLockKPFCostLevel;
            buttonCollection["tabParticipation"] = btnLockParticipation;
            buttonCollection["tabMineParameters"] = btnLockMineParameters;
            buttonCollection["Base Calc Process"] = btnBaseCalcsHeader;
            buttonCollection["HODEarn5"] = btnBaseCalcs;
            buttonCollection["HODEarn30"] = btnHODCalcs;
            buttonCollection["BonusShiftsEarn5"] = btnGangLevelCalcs;
            buttonCollection["BonusShiftsEarn15"] = btnEmployeeCalcs;

            #endregion

            #region BaseData Extracts

            //Extract Base data

            extractConfiguration();
            extractDesignations();
            extractKPFs();
            //extractEarningsCodes();

            #endregion

            //Extract Tab Info
            loadInfo();  

        }

        private void extractDesignations()
        {
            Designations = TB.createDataTableWithAdapterSelectAll(Base.BaseConnectionString, "Designation", " where Miningtype = '" + BusinessLanguage.MiningType +
                                                                                                           "' and Bonustype = '" + BusinessLanguage.BonusType + "'");

            lstNames = TB.loadDistinctValuesFromColumn(Designations, "Designation");


            foreach (string s in lstNames)
            {

                cboHODDesignation.Items.Add(s.Trim());

            }

        }

        private void setEnvironment()
        {

            Base.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Base.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Base.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Base.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Base.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();

            Base.BaseConnectionString = Base.ServerName;
            Base.Directory = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerPath"])).Trim();

            Analysis.Drive = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Drive"];
            Analysis.Integrity = System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Integrity"];
            Analysis.Userid = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Userid"])).Trim();
            Analysis.PWord = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "Password"])).Trim();
            Analysis.ServerName = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Analysis.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();

            Base.ADTeamConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.ClockConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.DBConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.StopeConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.AnalysisConnectionString = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "ServerName"])).Trim();
            Base.BackupPath = Encoding.Unicode.GetString(Convert.FromBase64String(System.Configuration.ConfigurationSettings.AppSettings[strServerPath + "BackupPath"])).Trim();

            #region oleDBConnectionStringBuilder


            if (strServerPath.ToString().Contains("Development") || strServerPath.ToString().Contains("Support"))
            {
                strServerPath = "Development";

                Base.DBConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.StopeConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.AnalysisConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.BaseConnectionString = Environment.MachineName.Trim() + Base.ServerName;
                Base.ADTeamConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Analysis.AnalysisConnectionString = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Analysis.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
                Base.ServerName = Environment.MachineName.Trim() + Base.ServerName.Trim();
            }

            OleDbConnectionStringBuilder builder = new OleDbConnectionStringBuilder();
            builder.ConnectionString = @"Data Source=" + Base.ServerName;
            builder.Add("Provider", "SQLOLEDB.1");
            builder.Add("Initial Catalog", Base.DBName);
            //builder.Add("Persist Security Info", "False");
            builder.Add("User ID", Base.Userid);
            builder.Add("Password", Base.PWord);

            string strdb = Base.DBName;
            //string strPath = Base.Directory.Replace("data\\", "reports\\") + strdb.Replace(BusinessLanguage.Period, "").Replace("1000", "Conn") + ".udl";
            //string strPath = "z:\\icalc\\Harmony\\Kusasalethu\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";
            string strPath = "c:\\iCalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\" + strdb.Replace(BusinessLanguage.Period, "").Replace("4000", "Conn") + ".udl";
            //MessageBox.Show("MEtatreport path en connfile :" + strPath.Trim());

            FileInfo fil = new FileInfo(strPath);

            try
            {
                File.Delete(strPath);
                Application.DoEvents();
            }
            catch (Exception ex)
            {
                MessageBox.Show("delete of udl failed: " + ex.Message);
            }

            switch (strServerPath)
            {
                case "Test":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;


                case "Development":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Integrated Security", "SSPI");
                    builder.Add("Trusted_Connection", "True");
                    break;

                case "Production":
                    builder.Add("Persist Security Info", "True");
                    builder.Add("Trusted_Connection", "True");
                    break;

            }

            //MessageBox.Show("Path: " + strPath);
            bool _check = Shared.CreateUDLFile(strPath, builder);

            if (_check)
            { }
            else
            {
                MessageBox.Show("Error in creation of UDL file", "ERROR", MessageBoxButtons.OK);
            }
            //xxxxxxxxxxxxxxxxxxxxxxxxxxxxx
            #endregion

            myConn.ConnectionString = Base.DBConnectionString;

        }

        static void CreateUDLFile(string FileName, OleDbConnectionStringBuilder builder)
        {
            try
            {
                string conn = Convert.ToString(builder);
                MSDASC.DataLinksClass aaa = new MSDASC.DataLinksClass();
                aaa.WriteStringToStorage(FileName, conn, 1);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in creation of UDL file - " + ex.Message, "ERROR", MessageBoxButtons.OK);
            }
        }

        private void extractEarningsCodes()
        {

            earningsCodes = Base.SelectEarningsCodes(Base.DBConnectionString);

        }

        private void extractPrimaryKeys(clsMain.clsMain main)
        {
            //A threat is started to extract the primary keys of selected tables.
            //The primary keys are stored in clsMain.
            //When the user select one of the selected tables tab on the front-end,
            //the list are passed from clsMain into the primary keys list.
            //No extracts are done from the databases and that makes the audit table fast.
            //ExtractKeys(Base);
            Thread t = new Thread(ExtractKeys);   // Kick off a new thread
            t.Start(main);
        }

        static void ExtractKeys(object main)
        {

            clsMain.clsMain M = (clsMain.clsMain)main;
            M.extractPrimaryKey();

        }

        public void extractDBTableNames(ListBox lstbox)
        {
            connectToDB();

            if (myConn.State == ConnectionState.Open)
            {
                List<string> lstTableNames = Base.getListOfTableNamesInDatabase(Base.DBConnectionString);
                Base.DBTables = lstTableNames;
                lstbox.Items.Clear();
                switch (lstTableNames.Count)
                {
                    case 0:
                        lstbox.Items.Add("No tables in database");
                        break;
                    default:
                        foreach (string s in lstTableNames)
                        {
                            lstbox.Items.Add(s);
                        }
                        break;
                }
            }
        }

        private void extractConfiguration()
        {

            Configs = Base.SelectConfigs(Base.BaseConnectionString, BusinessLanguage.MiningType, BusinessLanguage.BonusType);

            grdConfigs.DataSource = Configs;

            foreach (DataRow dr in Configs.Rows)
            {
                //This extract the value identifying the first 3 digits that the gang must conform to.
                if (dr["PARAMETERNAME"].ToString().Trim() == "HOD"
                    && dr["PARM1"].ToString().Trim() == "WAGECODE")
                {
                    for (int i = 5; i <= 10; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strWagecodes = strWagecodes + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strWagecodes = "(" + strWagecodes.Trim().Substring(1) + ")";

                }


                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING"
                    && dr["PARM1"].ToString().Trim() == "ACTIVITY")
                {
                    strActivity = string.Empty;

                    for (int i = 5; i <= 10; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strActivity = strActivity + ",'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strActivity = "(" + strActivity.Trim().Substring(1) + ")";
                }
            }
        }

        private void extractKPFs()
        {
            //Check if KPF exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "KPF");

            if (intCount > 0)
            {
                //YES

                KPF = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "KPF ", "");

            }

            cboDeptParametersKPF.Items.Clear();

            #region loadSubsection
            lstNames = TB.loadDistinctValuesFromColumn(KPF, "KPF");

            if (lstNames.Count > 1)
            {

                foreach (string s in lstNames)
                {
                    if (cboDeptParametersKPF.Items.Contains(s))
                    { }
                    else
                    {
                        cboDeptParametersKPF.Items.Add(s.Trim());

                    }
                }

                cboDeptParametersKPF.Text = cboDeptParametersKPF.Items[0].ToString();
            }

            #endregion

        }

        private void loadInfo()
        {
            strWherePeriod = "  where period = '" + BusinessLanguage.Period + "'";
            //Check if records in calendar exists with the selected period
            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "CALENDAR", strWherePeriod);

            if (Calendar.Rows.Count > 0)
            {

                //Run the extraction of the primary keys on its own threat.
                Shared.extractPrimaryKeys(Base);

                //Run the extraction of the views.
                Shared.createViews(Base);

                //Check if formulas exist.  If not, copy
                Shared.copyFormulas(Base);

                if (myConn.State == ConnectionState.Open)
                {
                    //evaluateAll();
                    evaluateCalendar();

                }
                else
                {
                    connectToDB();
                    //evaluateAll();
                    evaluateCalendar();
                    //Create the tab names
                    foreach (TabPage tp in tabInfo.TabPages)
                    {
                        tp.Text = tp.Tag.ToString();
                    }

                    listBox2.SelectedIndex = 0;
                    // listBox2_SelectedIndexChanged("Method", null);
                }

            }
            else
            {
                //NO....
                //1. Get Previous months info  ==> MAG NIE MEER HIERIN GAAN NIE!!!!!!!!!!!!!!!!!!!!!!!

                getHistory();

                //2. Check if PREVIOUS months DB exists
                //if (BusinessLanguage.checkIfFileExists(Base.Directory + "\\" + prevDatabaseName + Base.DBExtention))
                //{
                //3. If exist - Create this selected DB and copy Formulas, Rates and Factors to the new database.
                DialogResult result = MessageBox.Show("Do you want to start a new Bonus Period: " + BusinessLanguage.Period + "?",
                                       "Information", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:
                        this.Cursor = Cursors.WaitCursor;
                        backupAndRestoreDB();
                        copyFormulas();
                        extractDBTableNames(listBox1);


                        //Run the extraction of the primary keys on its own threat.
                         
                        Shared.extractPrimaryKeys(Base);
                        //evaluateAll();
                        evaluateCalendar();
                        //Create the tab names
                        foreach (TabPage tp in tabInfo.TabPages)
                        {
                            tp.Text = tp.Tag.ToString();
                        }

                         
                        listBox2.SelectedIndex = 0;
                        // listBox2_SelectedIndexChanged("Method", null);
                         


                        this.Cursor = Cursors.Arrow;
                        break;

                    case DialogResult.No:
                        btnSelect_Click("METHOD", null);
                        break;
                }
            }
        }

        private void evaluateAll()
        {
            evaluateCalendar();
            //evaluateClockedShifts();
            //evaluateSubsectionDept();
            //evaluateDeptParameters();
            //evaluateKPFCostLevel();
            //evaluateParticipation();
            //evaluateMineParameters();
            //evaluateHOD();
            //evaluateArtisans();
            ////evaluateOfficials();
            //evaluateLabour();
            //evaluateEmployeePenalties();
            //evaluateRates();
            extractDBTableNames(listBox1);

        }

        private void evaluateMineParameters()
        {
            // Display die HOD info
            MineParameters.Rows.Clear();

            loadMineParameters();

            hideColumnsOfGrid("grdMineParameters");
        }

        private void evaluateHOD()
        {
            // Display die HOD info
            HOD.Rows.Clear();

            loadHOD();

            hideColumnsOfGrid("grdHOD");
        }

        private void evaluateArtisans()
        {
            // Display die Artisan info
            Artisans.Rows.Clear();

            loadArtisans();

            hideColumnsOfGrid("grdArtisan");
        }

        private void confirmCopyandCreate()
        {
            listBox2.Items.Add("No sections found");

            this.Cursor = Cursors.WaitCursor;

            #region Create the new DB
            //Create the new database
            Base.createDatabase(Base.DBName, Base.ServerName);

            myConn = Base.DBConnection;
            myConn.Open();

            TB.createEmployeePenalties(Base.DBConnectionString);
            TB.createCalendarTable(Base.DBConnectionString);
            TB.createOffday(Base.DBConnectionString);
            TB.createEmployeePenalties(Base.DBConnectionString);

            //Extract Calendar again and insert into 
            DataTable calendar = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Calendar");
            grdCalendar.DataSource = calendar;

            listBox2.Items.Clear();
            listBox2.Items.Add("No sections exist yet");

            panel2.Enabled = false;
            panel3.Enabled = false;
            panel4.Enabled = false;

            this.Cursor = Cursors.Arrow;

            #endregion

        }

        private void copyFormulas()
        {
            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable dtBaseFormulas = Analysis.SelectAllFormulasPerDatabaseName(Base.DBCopyName, Base.AnalysisConnectionString);
            if (dtBaseFormulas.Rows.Count > 0)
            {
                foreach (DataRow row in dtBaseFormulas.Rows)
                {
                    //Check if the receiving table already contains this formula.
                    object intCount = Analysis.countcalcbyname(Base.DBName + BusinessLanguage.Period.Trim(), row["TABLENAME"].ToString(),
                                      row["CALC_NAME"].ToString(), Base.AnalysisConnectionString);

                    if ((int)intCount > 0)
                    {
                        //rename the formula name to be inserted to NEW

                    }
                    else
                    {
                        //insert the formula.
                        Base.CopyFormulas(Base.DBName + strprevPeriod.Trim(),
                                          Base.DBName + BusinessLanguage.Period.Trim(),
                                          Analysis.AnalysisConnectionString);
                        break;
                    }
                }
            }
            else
            {
                MessageBox.Show("No formulas exist on " + "\n" + "database: " + Base.DBCopyName + "\n" + "tablename: " + TB.TBCopyName +
                                "\n" + "therefor" + "\n" + "nothing will be copied", "Information", MessageBoxButtons.OK);
            }
        }

        private void getHistory()
        {
            #region Generate previous months db name
            //Calculate the previous months db name
            string Year = txtPeriod.Text.Trim().Substring(0, 4);
            strprevPeriod = txtPeriod.Text.Trim();

            if (txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2) == "01")
            {
                strprevPeriod = Convert.ToString(Convert.ToInt16(Year) - 1) + "12";
                prevDatabaseName = Base.DBName.Replace(txtPeriod.Text.Trim(), strprevPeriod);
            }
            else
            {
                string strMonth = Convert.ToString(Convert.ToInt16(txtPeriod.Text.Trim().Substring(txtPeriod.Text.Trim().Length - 2)) - 1);
                if (strMonth.Length == 1)
                {
                    strMonth = "0" + strMonth;
                }

                strprevPeriod = Year + strMonth;
                prevDatabaseName = Base.DBName.Replace(txtPeriod.Text.Trim(), strprevPeriod);
            }

            Base.DBCopyName = prevDatabaseName;

            #endregion

        }

        private void createAndCopyCalendar()
        {

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar");

            foreach (DataRow rr in Calendar.Rows)
            {
                rr["FSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(1)).ToString("yyyy-MM-dd");
                rr["LSH"] = (Convert.ToDateTime(rr["LSH"].ToString().Trim()).AddDays(31)).ToString("yyyy-MM-dd");
            }

            TB.saveCalculations2(Calendar, Base.DBConnectionString, "", "CALENDAR");
            this.Cursor = Cursors.Arrow;
        }

        private void createAndCopyStatus()
        {
            getHistory();

            TB.createStatusTable(Base.DBConnectionString);
            myConn.Close();

            //create the Status datatable from the previous periods'table.
            Base.DBName = Base.DBCopyName;
            connectToDB();

            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");

            #region signoff from previous months DB and signon to this new DB

            myConn.Close();

            Base.DBName = TB.DBName;

            //Connect to the database that you want to copy from and load the tables into the listbox2.  Afterwards, change the db.dbname to the main database name.
            connectToDB();

            #endregion

            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            foreach (DataRow rr in Status.Rows)
            {
                strSQL.Append("insert into Status values('" + rr["MININGTYPE"].ToString().Trim() +
                              "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + rr["SECTION"].ToString().Trim() +
                              "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                              "','" + rr["STATUS"].ToString().Trim() + "','" + rr["LOCKED"].ToString().Trim() + "');");

            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            Application.DoEvents();
            TB.InsertData(Base.DBConnectionString, "Update Status set status = 'N', locked = '0'");
            Status = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from Status");
            Application.DoEvents();
            this.Cursor = Cursors.Arrow;
        }

        private void backupAndRestoreDB()
        {
            //copy the data of the previous period to the current period.
            //xxxxxxxxxxxxxxxxxxxxx
            this.Cursor = Cursors.WaitCursor;
            Base.createNewPeriodsData(Base.DBConnectionString, BusinessLanguage.Period, strprevPeriod);
            this.Cursor = Cursors.Arrow;

        }

        private void evaluateInputProcessStatus()
        {

            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere + " and category = 'Input Process'");

            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

            }
            else
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = 'Input Process'" +
                                      " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where category = 'Header' and process = 'Input Process'" +
                                     " and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");

                btnLock.Text = "Lock";

            }

            evaluateStatus();

        }

        private void evaluateStatus()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "STATUS");

            if (intCount > 0)
            {
                //Status exists,  
                loadStatus();
            }
            else
            {
                createAndCopyStatus();
            }

        }

        private void evaluateClockedShifts()
        {
            //xxxxxxxxxxx
            Clocked = Base.CShifts;
            grdClocked.DataSource = Clocked;

        }

        private void statusChangeButtonColors()
        {
            foreach (DataRow rr in Status.Rows)
            {
                if (rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "Exit")
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        btnRefresh.Visible = false;
                        btnx.Visible = false;

                        pictBox.Visible = false;
                        pictBox2.Visible = false;
                        //calcTime.Enabled = false;
                    }
                }
                else
                {
                    if (rr["STATUS"].ToString().Trim() == "Y")
                    {
                        string strButtonName = rr["PROCESS"].ToString().Trim();
                        Control c = (Control)buttonCollection[strButtonName];
                        c.BackColor = Color.LightGreen;

                    }
                    else
                    {
                        if (rr["STATUS"].ToString().Trim() == "P")
                        {
                            string strButtonName = rr["PROCESS"].ToString().Trim();
                            Control c = (Control)buttonCollection[strButtonName];
                            c.BackColor = Color.Orange;
                        }
                        else
                        {
                            if (rr["STATUS"].ToString().Trim() == "N" &&
                                pictBox.Visible == true &&
                                rr["CATEGORY"].ToString().Trim().Substring(0, 4) == "CALC")
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.Orange;
                            }
                            else
                            {
                                string strButtonName = rr["PROCESS"].ToString().Trim();
                                Control c = (Control)buttonCollection[strButtonName];
                                c.BackColor = Color.PowderBlue;
                            }
                        }
                    }
                }

                Application.DoEvents();
            }
        }

        private void evaluateLabour()
        {
            //xxxxxxxxxxxxxxx
                      
            Labour = Base.Labour;
            Labour.TableName = "Labour";
            grdLabour.DataSource = Labour;

            lstNames = TB.loadDistinctValuesFromColumn(Labour, "EMPLOYEE_NO");


            foreach (string s in lstNames)
            {

                cboEmplPenEmployeeNo.Items.Add(s.Trim());
            }

            lstNames = TB.loadDistinctValuesFromColumn(Labour, "GANG");

            cboBonusShiftsGang.Items.Clear();


            foreach (string s in lstNames)
            {
                cboBonusShiftsGang.Items.Add(s.Trim());
            }    

            lstNames = TB.loadDistinctValuesFromColumn(Labour, "WAGECODE");  //amp
            if (lstNames.Contains("316E004"))
            {
            }
            else
            {
                lstNames.Add("316E004");
            }

            cboBonusShiftsWageCode.Items.Clear();
            foreach (string s in lstNames)
            {

                cboBonusShiftsWageCode.Items.Add(s.Trim());

            }     

            lstNames = TB.loadDistinctValuesFromColumn(Labour, "LINERESPCODE");  //amp
            cboBonusShiftsResponseCode.Items.Clear();
            foreach (string s in lstNames)
            {

                cboBonusShiftsResponseCode.Items.Add(s.Trim());

            }    
      
            hideColumnsOfGrid("grdLabour");
        }

        private void hideColumnsOfGrid(string gridname)
        {

            switch (gridname)
            {
                case "grdKPFCostLevel":
                    if (grdKPFCostLevel.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdKPFCostLevel.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdKPFCostLevel.Columns.Contains("MININGTYPE"))
                    {
                        this.grdKPFCostLevel.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdKPFCostLevel.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdKPFCostLevel.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;

                case "grdDeptParameters":
                    if (grdDeptParameters.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdDeptParameters.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdDeptParameters.Columns.Contains("MININGTYPE"))
                    {
                        this.grdDeptParameters.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdDeptParameters.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdDeptParameters.Columns["BONUSTYPE"].Visible = false;
                    }
                    if (grdDeptParameters.Columns.Contains("SECTION"))
                    {
                        this.grdDeptParameters.Columns["SECTION"].Visible = false;
                    }
                    return;

                case "grdMineParameters":
                    if (grdMineParameters.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdMineParameters.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdMineParameters.Columns.Contains("MININGTYPE"))
                    {
                        this.grdMineParameters.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdMineParameters.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdMineParameters.Columns["BONUSTYPE"].Visible = false;
                    }
                    return;

                case "grdSubsectionDept":
                    if (grdSubsectionDept.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdSubsectionDept.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdSubsectionDept.Columns.Contains("MININGTYPE"))
                    {
                        this.grdSubsectionDept.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdSubsectionDept.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdSubsectionDept.Columns["BONUSTYPE"].Visible = false;
                    }

                    return;


                case "grdLabour":

                    if (grdLabour.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdLabour.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("MININGTYPE"))
                    {
                        this.grdLabour.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdLabour.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdLabour.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdRates":
                    if (grdRates.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdRates.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("MININGTYPE"))
                    {
                        this.grdRates.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdRates.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdRates.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdCalendar":
                    if (grdCalendar.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdCalendar.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("MININGTYPE"))
                    {
                        this.grdCalendar.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdCalendar.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdCalendar.Columns["BONUSTYPE"].Visible = false;
                    }
                    break;

                case "grdHOD":
                    if (grdHOD.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdHOD.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdHOD.Columns.Contains("MININGTYPE"))
                    {
                        this.grdHOD.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdHOD.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdHOD.Columns["BONUSTYPE"].Visible = false;
                    }
                    if (grdHOD.Columns.Contains("PERIOD"))
                    {
                        this.grdHOD.Columns["PERIOD"].Visible = false;
                    }
                    break;

                //case "grdOfficials":
                //    if (grdOfficials.Columns.Contains("BUSSUNIT"))
                //    {
                //        this.grdOfficials.Columns["BUSSUNIT"].Visible = false;
                //    }
                //    if (grdOfficials.Columns.Contains("MININGTYPE"))
                //    {
                //        this.grdOfficials.Columns["MININGTYPE"].Visible = false;
                //    }
                //    if (grdOfficials.Columns.Contains("BONUSTYPE"))
                //    {
                //        this.grdOfficials.Columns["BONUSTYPE"].Visible = false;
                //    }
                //    if (grdOfficials.Columns.Contains("PERIOD"))
                //    {
                //        this.grdOfficials.Columns["PERIOD"].Visible = false;
                //    }
                //    break;

                case "grdArtisans":
                    if (grdArtisans.Columns.Contains("BUSSUNIT"))
                    {
                        this.grdArtisans.Columns["BUSSUNIT"].Visible = false;
                    }
                    if (grdArtisans.Columns.Contains("MININGTYPE"))
                    {
                        this.grdArtisans.Columns["MININGTYPE"].Visible = false;
                    }
                    if (grdArtisans.Columns.Contains("BONUSTYPE"))
                    {
                        this.grdArtisans.Columns["BONUSTYPE"].Visible = false;
                    }
                    if (grdArtisans.Columns.Contains("PERIOD"))
                    {
                        this.grdArtisans.Columns["PERIOD"].Visible = false;
                    }
                    break;
            }
        }

        private void evaluateCalendar()
        {
            panel3.Enabled = true;
            panel4.Enabled = true;
            listBox2.Enabled = true;
            listBox3.Enabled = true;

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "CALENDAR");

            if (intCount > 0)
            {
                //Calendar exists,
                loadCalendar();
                loadDatePickers(0);
                loadSectionsFromCalendar();
            }
            else
            {
                createAndCopyCalendar();
            }
        }

        private void loadCalendar()
        {
            // Display die calendar info

            Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar", strWhere);

            grdCalendar.DataSource = Calendar;


        }

        private void evaluateShifts()
        {

            System.Threading.Thread.Sleep(2500);

            evaluateClockedShifts();
            evaluateLabour();
        }

        private void loadStatus()
        {
            // Display die STATUS info
            //XXXXXXXXXXXXXXXXX
            Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);  //XXXXXXXXXXXXXXXXX
            if (Status.Rows.Count > 0)
            {
                statusChangeButtonColors();
            }
            else
            {
                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status");
                string tempSection = Status.Rows[0]["SECTION"].ToString().Trim();

                DataTable temp = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "STATUS",
                                 "Where section = '" + tempSection + "' and period = '" + BusinessLanguage.Period + "'");
                StringBuilder strSQL = new StringBuilder();
                strSQL.Append("BEGIN transaction; ");

                foreach (DataRow rr in temp.Rows)
                {
                    strSQL.Append("insert into Status values('" + rr["BUSSUNIT"].ToString().Trim() + "','" + rr["MININGTYPE"].ToString().Trim() +
                                    "','" + rr["BONUSTYPE"].ToString().Trim() + "','" + txtSelectedSection.Text +
                                  "','" + txtPeriod.Text.Trim() + "','" + rr["CATEGORY"].ToString().Trim() + "','" + rr["PROCESS"].ToString().Trim() +
                                  "','N','0');");

                }

                strSQL.Append("Commit Transaction;");
                TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
                Application.DoEvents();
                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);
            }
        }

        private void evaluateParticipation()
        {

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PARTICIPATION");

            if (intCount > 0)
            {

                Participation = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "PARTICIPATION", strWhere);

                grdParticipation.DataSource = Participation;
                hideColumnsOfGrid("grdParticipation");

            }

            cboParticipationColumns.Items.Clear();
            cboParticipationShow.Items.Clear();
            cboParticipationValues.Items.Clear();

            //Extract distinct columns and load into comboboxes
            List<string> lstColumnNames = General.getListOfColumnNames(Base.DBConnectionString, TB.TBName);

            foreach (string s in lstColumnNames)
            {
                cboParticipationColumns.Items.Add(s.Trim());
                cboParticipationShow.Items.Add(s.Trim());
            }
        }

        private void loadDatePickers(int Position)
        {
            //xxxxxxxxxxxxxxxx
            if (Calendar.Rows.Count > 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[Position]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[Position]["LSH"].ToString().Trim());
            }
            intNoOfDays = Base.calcNoOfDays(dateTimePicker2.Value, dateTimePicker1.Value);
        }

        private void loadSectionsFromCalendar()
        {
            lstNames = TB.loadDistinctValuesFromColumn(Calendar, "SECTION");

            if (lstNames.Count > 0)
            {


                txtSelectedSection.Text = Calendar.Rows[0]["Section"].ToString().Trim();
                label15.Text = Calendar.Rows[0]["Section"].ToString().Trim();
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() +
                           "' and period = '" + BusinessLanguage.Period + "'";

                strWhereSection = "where section = '" + Calendar.Rows[0]["Section"].ToString().Trim() + "'";
                listBox2.Items.Clear();

                if (lstNames.Count > 1)
                {
                    foreach (string s in lstNames)
                    {
                        if (s != "XXX")
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
                else
                {
                    if (lstNames.Count == 1)
                    {
                        foreach (string s in lstNames)
                        {
                            listBox2.Items.Add(s.Trim());
                        }
                    }
                }
            }
        }

        private void evaluateSubsectionDept()
        {
            // Display die SubsectionDept info
            SubsectionDept.Rows.Clear();

            loadSubsectionDept();

        }

        private void evaluateDeptParameters()
        {
            // Display die KPFCostLevel info
            DeptParameters.Rows.Clear();

            loadDeptParameters();

            hideColumnsOfGrid("grdDeptParameters");

        }

        private void evaluateKPFCostLevel()
        {
            // Display die KPFCostLevel info
            KPFCostLevel.Rows.Clear();

            loadKPFCostLevel();

            hideColumnsOfGrid("grdKPFCostLevel");

        }

        private void loadKPFCostLevel()
        {
            //Check if KPFs exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "KPFCOSTLEVEL");

            if (intCount > 0)
            {
                //YES
                KPFCostLevel = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "KPFCOSTLEVEL", strWhere, 8);

            }

            grdKPFCostLevel.DataSource = KPFCostLevel;
            grdKPFCostLevel.Refresh();
        }


        private void loadMonitor()
        {
            //Check if Monitor exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "MONITOR");

            if (intCount > 0)
            {
                //YES

                Monitor = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Monitor");

                if (Monitor.Rows.Count == 0)
                {
                    extractMonitorData();
                }
                else
                {

                }
            }
            else
            {
                TB.createMonitor(Base.DBConnectionString);
                TB.TBName = "MONITOR";
                Monitor = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Monitor");
                extractMonitorData();

            }

        }

        private void extractMonitorData()
        {
            string strSQL = "select PROCESS, PROCESS_CATEGORY, PROCESS_IND, PROCESS_VALUE from MONITOR  ";

            DataTable tempDataTable = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            foreach (DataRow _row in tempDataTable.Rows)
            {
                if (string.IsNullOrEmpty(_row[0].ToString()))
                {
                }
                else
                {
                    Monitor.Rows.Add(_row.ItemArray);
                }
            }

            saveXXXMonitor();

        }

        private void saveXXXMonitor()
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

        }

        //private void evaluateGangLinking()
        //{
        //     Display die Ganglink info
        //    GangLink.Rows.Clear();

        //    loadGangLinking();

        //    lstNames = TB.loadDistinctValuesFromColumn(GangLink, "GANG");

        //    if (lstNames.Count > 1)
        //    {
        //        foreach (string s in lstNames)
        //        {
        //            cboOffDaysGang.Items.Add(s.Trim());
        //        }
        //    }

        //    cboGangLinkGangType.Text = "DEVELOPMENT";

        //}

        private void loadMO()
        {
            strMO = "";
            foreach (DataRow dr in Configs.Rows)
            {
                if (dr["PARAMETERNAME"].ToString().Trim() == "GANGLINKING" && dr["PARM1"].ToString().Trim() == "MO" && dr["PARM2"].ToString().Trim() == txtSelectedSection.Text)
                {
                    for (int i = 3; i <= 5; i++)
                    {
                        if (dr[i].ToString().Trim() != "Q")
                        {
                            strMO = "'" + dr[i].ToString().Trim() + "'";
                        }
                    }

                    strMO = "(" + strMO.Trim() + ")";
                }
            }
        }

        //private void loadGangLinking()
        //{
        //    //Check if ganglinking exists
        //    Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "GANGLINK");

        //    if (intCount > 0)
        //    {
        //        //YES
        //        GangLink = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "GangLink", strWhere);
        //        cboGangLinkGang.Items.Clear();
        //        List<string> lstGangs = TB.loadDistinctValuesFromColumn(Labour, "GANG");
        //        for (int i = 0; i <= lstGangs.Count - 1; i++)
        //        {
        //            cboGangLinkGang.Items.Add(lstGangs[i].ToString().Trim());
        //        }

        //    }
        //    else
        //    {
        //        //NO the ganglink table does not exist. 
        //        //Create the ganglink table
        //        //Check if BonusShifts Exists

        //        intCount = TB.checkTableExist(Base.DBConnectionString, "CLOCKEDSHIFTS");

        //        if (intCount > 0)
        //        {

        //            loadMonitor();
        //            TB.createGangLink(Base.DBConnectionString);
        //            TB.TBName = "GANGLINK";

        //            DialogResult result = MessageBox.Show("GangLink table does not exist or is corrupted. Do you want to recreate the table?", "Information", MessageBoxButtons.YesNo);

        //            //switch (result)
        //            //{
        //            //    case DialogResult.Yes:

        //            //        TB.createGangLink(Base.DBConnectionString);
        //            //        return;

        //            //    case DialogResult.No:
        //            //        return;
        //            //}

        //            ////extractGangLinkData();
        //            //saveXXXGangLink();

        //        }
        //        else
        //        {
        //        }

        //    }

        //    grdGangLink.DataSource = GangLink;

        //    grdGangLink.Refresh();
        //}

        private void evaluateRates()
        {
            // Display die Abnormal info
            Rates.Rows.Clear();

            loadRates();

        }

        private void evaluateEmployeePenalties()
        {
            // Display die EmployeePenalties info
            EmplPen.Rows.Clear();

            loadEmployeePenalties();

        }

        private void loadEmployeePenalties()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            if (intCount > 0)
            {
                //YES

                EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES");

            }
            else
            {
                //NO
                //Check if Bonusshifts Exists

                intCount = TB.checkTableExist(Base.DBConnectionString, "BONUSSHIFTS");

                if (intCount > 0)
                {
                    TB.createEmployeePenalties(Base.DBConnectionString);
                    TB.TBName = "EMPLOYEEPENALTIES";
                    EmplPen = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "EMPLOYEEPENALTIES ", strWhere);

                }
                else
                {
                }

            }

            grdEmplPen.DataSource = EmplPen;

            grdEmplPen.Refresh();

        }


        private void connectToDB()
        {

            if (myConn.State == ConnectionState.Closed)
            {
                try
                {
                    myConn.Open();
                }
                catch (SystemException eee)
                {
                    MessageBox.Show(eee.ToString());
                }
            }
        }

        private void loadRates()
        {
            //Check if ABNORMAL exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Rates");

            if (intCount > 0)
            {
                //YES

                Rates = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Rates");

            }
            else
            {
                //NO - Rates DOES NOT EXIST 
            }

            grdRates.DataSource = Rates;

            grdRates.Refresh();

        }

        private void loadHOD()
        {
            //Check if HOF exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "HOD");

            if (intCount > 0)
            {
                //YES

                HOD = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "HOD ", strWhere);

            }

            //Load the mine parameters into the mine parameters textboxes
            txtMineCostA.Text = MineParameters.Rows[0]["COST_ACTUAL"].ToString().Trim();
            txtMineCostP.Text = MineParameters.Rows[0]["COST_PLANNED"].ToString().Trim();
            txtMineSafetyA.Text = MineParameters.Rows[0]["SAFETY_ACTUAL"].ToString().Trim();
            txtMineTECA.Text = MineParameters.Rows[0]["TONSPERTEC_ACTUAL"].ToString().Trim();
            txtMineTECP.Text = MineParameters.Rows[0]["TONSPERTEC_PLANNED"].ToString().Trim();
            grdHOD.DataSource = HOD;
            grdHOD.Refresh();

            hideColumnsOfGrid("grdHOD");

        }

        private void loadArtisans()
        {
            //Check if HOF exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Artisans");

            if (intCount > 0)
            {
                //YES

                Artisans = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Artisans", strWhere);

            }

            //Change the column names of ITEM1-ITEM5
            grdArtisans.DataSource = Artisans;
            grdArtisans.Refresh();

            hideColumnsOfGrid("grdArtisans");

        }

        //private void loadOfficials()
        //{
        //    //Check if HOF exists
        //    Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "Officials");

        //    if (intCount > 0)
        //    {
        //        //YES

        //        Officials = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Officials", strWhere);

        //    }

        //    //Change the column names of ITEM1-ITEM5
        //    grdOfficials.DataSource = Officials;
        //    grdOfficials.Refresh();

        //}

        private void loadMineParameters()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "MineParameters");

            if (intCount > 0)
            {
                //YES

                MineParameters = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "MineParameters ", strWhere);
            }

            grdMineParameters.DataSource = MineParameters;
            grdMineParameters.Refresh();

            hideColumnsOfGrid("grdMineParameters");
        }

        private void loadSubsectionDept()
        {
            //Check if miners exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "SUBSECTIONDEPT");

            if (intCount > 0)
            {
                //YES

                SubsectionDept = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "SubsectionDept ", strWhere);

            }

            cboSubsection.Items.Clear();
            cboArtisanSubsection.Items.Clear();
            //cboOfficialsSubsection.Items.Clear();
            cboDeptParametersSubsection.Items.Clear();
            cboDeptParametersDepartment.Items.Clear();
            cboDepartment.Items.Clear();
            cboHODModel.Items.Clear();

            #region loadSubsection
            lstNames = TB.loadDistinctValuesFromColumn(SubsectionDept, "Subsection");

            if (lstNames.Count > 1)
            {

                foreach (string s in lstNames)
                {


                    if (cboSubsection.Items.Contains(s))
                    { }
                    else
                    {
                        cboDeptParametersSubsection.Items.Add(s.Trim());
                        cboSubsection.Items.Add(s.Trim());
                        cboHODSubsection.Items.Add(s.Trim());
                        cboArtisanSubsection.Items.Add(s.Trim());
                        //cboOfficialsSubsection.Items.Add(s.Trim());
                    }
                }

                cboSubsection.Text = cboSubsection.Items[0].ToString();
            }

            #endregion

            #region loadDepartment
            cboArtisanDepartment.Items.Clear();
            //cboOfficialsDepartment.Items.Clear();
            lstNames = TB.loadDistinctValuesFromColumn(SubsectionDept, "Department");
            if (lstNames.Count > 1)
            {

                foreach (string s in lstNames)
                {
                    if (cboDepartment.Items.Contains(s))
                    { }
                    else
                    {
                        cboArtisanDepartment.Items.Add(s.Trim());
                        //cboOfficialsDepartment.Items.Add(s.Trim());
                        cboDepartment.Items.Add(s.Trim());
                        cboHODDepartment.Items.Add(s.Trim());
                    }
                }
                cboDepartment.Text = cboDepartment.Items[0].ToString();
            }

            #endregion

            #region loadHODModel
            lstNames = TB.loadDistinctValuesFromColumn(SubsectionDept, "HODModel");

            if (lstNames.Count > 1)
            {

                foreach (string s in lstNames)
                {
                    if (cboHODModel.Items.Contains(s))
                    { }
                    else
                    {
                        cboHODModel.Items.Add(s.Trim());
                    }
                }
                cboHODModel.Text = cboHODModel.Items[0].ToString();
            }

            #endregion

            grdSubsectionDept.DataSource = SubsectionDept;
            grdSubsectionDept.Refresh();

            hideColumnsOfGrid("grdSubsectionDept");
        }

        private void loadDeptParameters()
        {
            //Check if KPFs exists
            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "DeptParameters");

            if (intCount > 0)
            {
                //YES
                DeptParameters = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "DeptParameters", strWhere, 8);
                ParameterNames.Clear();

                foreach (DataRow r in DeptParameters.Rows)
                {
                    if (string.IsNullOrEmpty(r["KPFPARAMETERDESC"].ToString()) || r["KPFPARAMETERDESC"].ToString().Trim() == "")
                    {

                        r["KPFPARAMETERDESC"] = r["DeptParameters"];

                    }

                    //ParameterNames.Add(r["DeptParameters"].ToString().Trim(), r["KPFPARAMETERDESC"].ToString().Trim());
                }

                DeptParameters.AcceptChanges();
                TB.saveCalculations2(DeptParameters, Base.DBConnectionString, "", "DeptParameters");
                Application.DoEvents();
            }

            cboDeptParametersColumns.Items.Clear();
            cboDeptParametersValues.Items.Clear();
            cboDeptParametersShow.Items.Clear();

            //Extract distinct columns and load into comboboxes
            List<string> lstColumnNames = General.getListOfColumnNames(Base.DBConnectionString, "DeptParameters");

            foreach (string s in lstColumnNames)
            {
                cboDeptParametersColumns.Items.Add(s.Trim());
                cboDeptParametersShow.Items.Add(s.Trim());
            }

            grdDeptParameters.DataSource = DeptParameters;
            grdDeptParameters.Refresh();

        }

        private void importTheSheet(string importFilename)
        {
            string path = BusinessLanguage.InputDirectory + Base.DBName;

            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                string filename = BusinessLanguage.InputDirectory + Base.DBName + importFilename;
                bool fileCheck = BusinessLanguage.checkIfFileExists(filename);

                if (fileCheck)
                {
                    FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                    spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                    fs.Close();
                    //If the file was SURVEY, all sections production data will be on this datatable.
                    //Only the selected section's data must be saved.

                    saveTheSpreadSheetToTheDatabase();
                }
                else
                {
                    MessageBox.Show("File " + filename + " - does not exist", "Check", MessageBoxButtons.OK);
                }

                //Check if file exists
                //If not  = Message
                //If exists ==>  Import
            }
            catch
            {
                MessageBox.Show("File " + importFilename + " - is inuse by another package?", "Check", MessageBoxButtons.OK);
            }
        }

        private void saveTheSpreadSheetToTheDatabase()
        {
            foreach (DataTable dt in spreadsheet.WorkbookData.Tables)
            {
                if (dt.TableName == "SURVEY" || dt.TableName == "Survey")
                {
                    for (int i = 1; i <= dt.Rows.Count - 1; i++)
                    {
                        if (dt.Rows[i][3].ToString().Trim() == txtSelectedSection.Text.Trim())
                        {
                        }
                        else
                        {
                            dt.Rows[i].Delete();

                        }
                    }

                }

                dt.AcceptChanges();
                //checker = true;

                TB.TBName = dt.TableName.ToString().ToUpper();
                TB.recreateDataTable();

                //Extract column names
                string strColumnHeadings = TB.getFirstRowValues(dt, Base.AnalysisConnectionString);

                switch (strColumnHeadings)
                {
                    case null:
                        break;

                    case "":
                        break;

                    default:


                        if (myConn.State == ConnectionState.Closed)
                        {
                            try
                            {
                                myConn = Base.DBConnection;
                                myConn.Open();

                                //create a table
                                bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);

                                if (tableCreate)
                                {
                                    MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);
                                }
                                else
                                {
                                    MessageBox.Show("Try again after correction of spreadsheet - input data.", "Information", MessageBoxButtons.OK);
                                }

                                //checker = false;
                            }
                            catch (System.Exception ex)
                            {
                                System.Windows.Forms.MessageBox.Show(ex.GetHashCode() + " " + ex.ToString(), "MyProgram", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                        }
                        else
                        {
                            //create a table
                            bool tableCreate = TB.createDatabaseTable(Base.DBConnectionString, strColumnHeadings);

                            if (tableCreate)
                            {
                                tableCreate = TB.copySpreadsheetToDatabaseTable(Base.DBConnectionString, dt);
                                MessageBox.Show("Data successfully imported", "Information", MessageBoxButtons.OK);

                            }
                            else
                            {
                                MessageBox.Show("Data was not imported.", "Information", MessageBoxButtons.OK);
                            }
                        }

                        break;
                }
            }
        }

        private String[] GetExcelSheetNames(string excelFile)
        {
            OleDbConnection objConn = null;
            System.Data.DataTable dt = null;

            try
            {
                // Connection String. Change the excel file to the file you
                // will search.
                String connString = "Provider=Microsoft.ACE.OLEDB.12.0;" +
                    "Data Source=" + excelFile + ";Extended Properties=Excel 12.0;";
                // Create connection object by using the preceding connection string.
                objConn = new OleDbConnection(connString);
                // Open connection with the database.
                objConn.Open();
                // Get the data table containg the schema guid.
                dt = objConn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                String[] excelSheets = new String[dt.Rows.Count];
                int i = 0;

                // Add the sheet name to the string array.
                foreach (DataRow row in dt.Rows)
                {
                    excelSheets[i] = row["TABLE_NAME"].ToString();
                    i++;
                }

                // Loop through all of the sheets if you want too...
                for (int j = 0; j < excelSheets.Length; j++)
                {
                    // Query each excel sheet.
                }

                return excelSheets;
            }
            catch
            {
                return null;
            }
            finally
            {
                // Clean up.
                if (objConn != null)
                {
                    objConn.Close();
                    objConn.Dispose();
                }
                if (dt != null)
                {
                    dt.Dispose();
                }
            }
        }

        private void btnImportADTeam_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;

            DataTable temp = new DataTable();
            if (Labour.Rows.Count > 0)
            {
                IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                              where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                              where locks.Field<string>("PROCESS").TrimEnd() == "tabLabour"
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }
            else
            {
                evaluateStatus();
                IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                              where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                              where locks.Field<string>("PROCESS").TrimEnd() == "tabLabour"
                                              select locks;


                temp = query1.CopyToDataTable<DataRow>();
            }
            if (temp.Rows[0]["STATUS"].ToString().Trim() == "N")
            {
                refreshLabour();

            }
            else
            {
                MessageBox.Show("BonusShifts is locked. Please unlock before refresh.  You WILL loose all previous updates.", "Information", MessageBoxButtons.OK);
            }

            this.Cursor = Cursors.Arrow;
        }

        private void refreshLabour()
        {

            #region extract the sheet name and FSH and LSH of the extract
            ATPMain.VkExcel excel = new ATPMain.VkExcel(false);


            bool XLSX_exists = File.Exists("C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\master" + BusinessLanguage.MiningType + BusinessLanguage.Period.Trim() + ".xlsx");
            bool XLS_exists = File.Exists("C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\master" + BusinessLanguage.MiningType + BusinessLanguage.Period.Trim() + ".xls");

            if (XLSX_exists.Equals(true))
            {
                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\master" + BusinessLanguage.MiningType + BusinessLanguage.Period.Trim() + ".xlsx", "BONTS2011");
                excel.SaveFile(BusinessLanguage.Period.Trim(), strServerPath);
                excel.CloseFile();
            }

            if (XLS_exists.Equals(true))
            {

                string status = excel.OpenFile("C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\master" + BusinessLanguage.MiningType + BusinessLanguage.Period.Trim() + ".xls", "BONTS2011");

                excel.SaveFile(BusinessLanguage.Period.Trim(),strServerPath);
                excel.CloseFile();
            }

            excel.stopExcel();

            string FilePath = "";

            string FilePath_XLSX = "C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\Engineering_" + BusinessLanguage.Period.Trim() + ".xlsx";

            string FilePath_XLS = "C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\Engineering_" + BusinessLanguage.Period.Trim() + ".xls";

            XLSX_exists = File.Exists(FilePath_XLSX);
            XLS_exists = File.Exists(FilePath_XLS);

            if (XLS_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\Engineering_" + BusinessLanguage.Period.Trim() + ".xls";
            }

            if (XLSX_exists.Equals(true))
            {
                FilePath = "C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\Engineering_" + BusinessLanguage.Period.Trim() + ".xlsx";
            }
            //excel.GetExcelSheets();
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Ace.OLEDB.12.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);
            IEnumerable<DataRow> query1 = from locks in dt.AsEnumerable()
                                          where locks.Field<string>("GANG NAME").Substring(2, 1) == "L" ||
                                                locks.Field<string>("GANG NAME").Substring(0, 6) == "JJPPRE"
                                          select locks;

            //Temp will contain a list of the gangs for the section
            DataTable Tempdt = query1.CopyToDataTable<DataRow>();

            dt = Tempdt.Copy();
            #region remove invalid records

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;
            DateTime SheetFSH;
            DateTime SheetLSH;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }
            else
            {
                strSheetFSH = strSheetFSHx.ToString();
            }

            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }
            else
            {
                strSheetLSH = strSheetLSHx.ToString();
            }

            SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());



            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            int intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);

            if (intNoOfDays <= 40)
            {
                for (int j = intNoOfDays + 1; j <= 40; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }
            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;

            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;
                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }
                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }
                }
            }

            //Add the extra columns
            dt.Columns.Add("BUSSUNIT");
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.Columns.Add("EMPLOYEETYPE");
            dt.Columns.Add("PERIOD");      //xxxxxxxx
            dt.AcceptChanges();

            MessageBox.Show("no of rows imported: " + dt.Rows.Count, "Information", MessageBoxButtons.OK);
            foreach (DataRow row in dt.Rows)
            {
                row["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "ENGINEERING";
                row["PERIOD"] = BusinessLanguage.Period;   //xxx
                if (row["GANG"].ToString().Length > 0)
                {
                    row["SECTION"] = "ENG";
                }
                else
                {
                    row["SECTION"] = "XXX";
                }
                if (row["WAGECODE"].ToString().Trim() == "")
                {
                    row["WAGECODE"] = "00000";
                }
                else
                {
                }

                row["EMPLOYEETYPE"] = Base.extractEmployeeType(Configs, row["WAGECODE"].ToString());

                //Replace all the null columns with a "-"
                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }

            //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
            DataColumn dcBussunit = new DataColumn();
            dcBussunit.ColumnName = "BUSSUNIT";
            dt.Columns.Remove("BUSSUNIT");
            dt.AcceptChanges();
            InsertAfter(dt.Columns, dt.Columns["BONUSTYPE"], dcBussunit);

            foreach (DataRow dr in dt.Rows)
            {
                dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
            }


            #endregion
            //exportToExcel("c:\\", dt);
            //Write to the database
            TB.saveCalculations2(dt, Base.DBConnectionString, "", "CLOCKEDSHIFTS");

            TB.InsertData(Base.DBConnectionString, "update clockedshifts set employeetype = '5 - ARMS' " +
                          "where employee_no in (select employee_no from employeelist where tablename = 'ARMS')");
            //==================================================================================
            Application.DoEvents();

            //TEMP HARD CODING

            grdClocked.DataSource = dt;
            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            string strSQL = "Select *,'SUBSECTION' as SUBSECTION,'DEPARTMENT' as DEPARTMENT,'HODMODEL' as HODMODEL,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS," +
                            "'0' as Q_SHIFTS " +
                            " from Clockedshifts where section = '" +
                            txtSelectedSection.Text.Trim() + "' order by employee_no";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";

            //jvdw laai die hele clockedshift table
            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);
            //==============================================================================
            BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            //exportToExcel("c:\\", BonusShifts);
            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;//jvdw
            sheetlhs = SheetLSH;//jvdw
            int intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            int intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            int intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {
            }

            if (intEndDay < 0 && intEndDay < -40)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 40;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the shifts
                Shared.evaluateDataTable(Base, "CLOCKEDSHIFTS");

                DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                switch (result)
                {
                    case DialogResult.OK:
                        Shared.evaluateDataTable(Base, "BONUSSHIFTS");
                        extractAndCalcShifts(intStartDay, intStopDay);
                        MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);
                        evaluateShifts();
                        break;

                    case DialogResult.Cancel:
                        break;

                }

                #endregion

            #endregion


                this.Cursor = Cursors.Arrow;

                MessageBox.Show("Shifts were imported successfully", "Information", MessageBoxButtons.OK);
                //}
            }

        }

        public void InsertAfter(DataColumnCollection columns, DataColumn currentColumn, DataColumn newColumn)
        {
            if (columns.Contains(currentColumn.ColumnName))
            {
                columns.Add(newColumn);
                //add the new column after the current one 
                columns[newColumn.ColumnName].SetOrdinal(currentColumn.Ordinal + 1);
            }
            else
            {

            }
        }


        private void extractAndCalcShifts(int DayStart, int DayEnd)
        {
            int intSubstringLength = 0;
            int intShiftsWorked = 0;
            int intAwopShifts = 0;
            int intQShifts = 0;

            foreach (DataRow row in BonusShifts.Rows)
            {
                foreach (DataColumn column in BonusShifts.Columns)
                {
                    //Import the subsection and department name
                    if ((column.ColumnName == "SUBSECTION"))
                    {
                        DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, row["GANG"].ToString().Trim());
                        if (sub.Rows.Count > 0)
                        {
                            row["SUBSECTION"] = sub.Rows[0]["SUBSECTION"].ToString().Trim();
                            row["DEPARTMENT"] = sub.Rows[0]["DEPARTMENT"].ToString().Trim();
                            row["HODMODEL"] = sub.Rows[0]["HODMODEL"].ToString().Trim();
                        }
                        else
                        {
                            row["SUBSECTION"] = "UNKNOWN";
                            row["DEPARTMENT"] = "UNKNOWN";
                            row["HODMODEL"] = "UNKNOWN";

                        }
                    }

                    if ((column.ColumnName == "SUD"))
                    {
                        row["SUD"] = row["GANG"].ToString().Substring(6, 1);

                    }

                    if ((column.ColumnName.Substring(0, 3) == "DAY"))
                    {
                        if (column.ColumnName.ToString().Length == 4)
                        {
                            intSubstringLength = 1;
                        }
                        else
                        {
                            intSubstringLength = 2;
                        }

                        if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                           Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                        {
                            if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u" || 
                                row[column].ToString().Trim() == "W" || row[column].ToString().Trim() == "w" || 
                                row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "q")
                            {
                                intShiftsWorked = intShiftsWorked + 1;
                            }
                            else
                            {
                                if (row[column].ToString().Trim() == "A")
                                {
                                    intAwopShifts = intAwopShifts + 1;
                                }
                                else
                                {
                                    if (row[column].ToString().Trim() == "Q" || row[column].ToString().Trim() == "q")
                                    {
                                        intQShifts = intQShifts + 1;
                                    }
                                }

                            }
                        }
                        else
                        {
                            row[column] = "*";
                        }
                    }
                    else
                    {
                        if (column.ColumnName == "BONUSTYPE")
                        {
                            row["BONUSTYPE"] = "SERVICES";
                        }
                    }
                }//foreach datacolumn

                row["SHIFTS_WORKED"] = intShiftsWorked;
                row["AWOP_SHIFTS"] = intAwopShifts;
                row["Q_SHIFTS"] = intQShifts;
                intShiftsWorked = 0;
                intAwopShifts = 0;
                intQShifts = 0;
            }

            //On BonusShifts the column PERIOD is part of the primary key.  Therefore must be moved xxxxxxxxx
            DataColumn dcPeriod = new DataColumn();
            dcPeriod.ColumnName = "PERIOD";
            BonusShifts.Columns.Remove("PERIOD");
            BonusShifts.AcceptChanges();
            InsertAfter(BonusShifts.Columns, BonusShifts.Columns["BONUSTYPE"], dcPeriod);

            foreach (DataRow dr in BonusShifts.Rows)
            {
                dr["PERIOD"] = BusinessLanguage.Period;
            }

            string strDelete = " where section = '" + txtSelectedSection.Text.Trim() +
                               "' and period = '" + BusinessLanguage.Period.Trim() + "'";

            TB.saveCalculations2(BonusShifts, Base.DBConnectionString, strDelete, "BONUSSHIFTS");
 

            ////printHTML(BonusShifts, "BONUSSHIFTS");

            //if (importdone == 0)//jvdw
            //{

            //    fillFixTable(fixShifts, sheetfhs, sheetlhs, noOFDay, DayStart, DayEnd);//Calls the method to load the fix clockedshiftstable
            //    importdone = 1;

            //}

            Application.DoEvents();
        }

        private string getSubsection(string gang)
        {
            DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, gang);
            if (sub.Rows.Count > 0)
            {
                return sub.Rows[0]["SUBSECTION"].ToString().Trim();

            }
            else
            {
                return "UNKNOWN";
            }
        }

        private string getDepartment(string gang)
        {
            DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, gang);
            if (sub.Rows.Count > 0)
            {

                return sub.Rows[0]["DEPARTMENT"].ToString().Trim();
            }
            else
            {

                return "UNKNOWN";

            }
        }

        private void btnLock_Click(object sender, EventArgs e)
        {
            string strProcess = tabInfo.SelectedTab.Name;

            if (btnLock.Text == "Lock")
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'Y' where process = '" + strProcess +
                                      "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                btnLock.Text = "Unlock";

            }

            else
            {

                TB.InsertData(Base.DBConnectionString, "Update STATUS set status = 'N' where process = '" + strProcess +
                                      "' and period = '" + txtPeriod.Text.Trim() + "' and section = '" + txtSelectedSection.Text.Trim() + "'");
                btnLock.Text = "Lock";

            }

            evaluateInputProcessStatus();
            openTab(tabProcess);

            Application.DoEvents();

        }

        private void btnInsertRow_Click(object sender, EventArgs e)
        {
            string strSQL = string.Empty;
            string strName = string.Empty;
            string strDesignation = string.Empty;
            string strDesignationDesc = string.Empty;

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabSubSectionDept":
                    #region tabSubsectionDept

                    if (cboSubsection.Text.Trim().Length != 0 && cboDepartment.Text.Trim().Length != 0 &&
                        cboHODModel.Text.Trim().Length != 0 && cboCostLevel.Text.Trim().Length != 0)
                    {
                        string strCostLevel = string.Empty;

                        if (cboCostLevel.Text.Contains("-"))
                        {
                            strCostLevel = cboCostLevel.Text.Substring(0, cboCostLevel.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strCostLevel = cboCostLevel.Text.Trim();
                        }

                        DataTable temp = new DataTable();
                        temp = SubsectionDept.Copy();

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();

                        }

                        temp.AcceptChanges();
                        DataRow dr = temp.NewRow();

                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["SUBSECTION"] = cboSubsection.Text.Trim();
                        dr["DEPARTMENT"] = cboDepartment.Text.Trim();
                        dr["HODMODEL"] = cboHODModel.Text.Trim();
                        dr["COSTLEVEL"] = strCostLevel;

                        temp.Rows.Add(dr);
                        //Create a total invalid delete.
                        string strDelete = " where Bussunit = '999'";
                        int rowindex = grdSubsectionDept.CurrentCell.RowIndex;
                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "SUBSECTIONDEPT");

                        evaluateSubsectionDept();
                        hideColumnsOfGrid("grdSubsectionDept");
                        grdSubsectionDept.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties
                    if (cboEmplPenEmployeeNo.Text.Trim().Length > 0 &&
                        txtPenaltyValue.Text.Trim().Length > 0 && cboPenaltyInd.Text.Trim().Length > 0)
                    {
                        DataRow dr;
                        dr = EmplPen.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["EMPLOYEE_NO"] = cboEmplPenEmployeeNo.Text.Trim();
                        dr["PENALTYVALUE"] = txtPenaltyValue.Text.Trim();
                        dr["PENALTYIND"] = cboPenaltyInd.Text.Trim();

                        EmplPen.Rows.Add(dr);

                        strSQL = "Insert into EmployeePenalties values ('" + BusinessLanguage.BussUnit +
                                 "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtSelectedSection.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + cboEmplPenEmployeeNo.Text.Trim() + "', '" + txtPenaltyValue.Text.Trim() +
                                 "', '" + cboPenaltyInd.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {
                        DataRow dr;
                        dr = Rates.NewRow();
                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["PERIOD"] = txtPeriod.Text.Trim();
                        dr["RATE_TYPE"] = txtRateType.Text.Trim();
                        dr["LOW_VALUE"] = txtLowValue.Text.Trim();
                        dr["HIGH_VALUE"] = txtHighValue.Text.Trim();
                        dr["RATE"] = txtRate.Text.Trim();

                        int rowindex = grdRates.CurrentCell.RowIndex;
                        strSQL = "Insert into Rates values ('" + BusinessLanguage.BussUnit +
                                 "', '" + BusinessLanguage.MiningType + "', '" + BusinessLanguage.BonusType +
                                 "', '" + txtRateType.Text.Trim() + "', '" + txtPeriod.Text.Trim() +
                                 "', '" + txtLowValue.Text.Trim() + "', '" + txtHighValue.Text.Trim() +
                                 "', '" + txtRate.Text.Trim() + "')";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateRates();
                        grdRates.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabDeptParameters":
                    #region tabDeptParameters

                    if (cboDeptParametersSubsection.Text.Trim().Length != 0 && cboDeptParametersKPF.Text.Trim().Length != 0 &&
                        cboDeptParametersKPFParameters.Text.Trim().Length != 0 && txtDeptParametersDesc.Text.Trim().Length != 0)
                    {

                        DataTable temp = new DataTable();
                        temp = DeptParameters.Copy();

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();

                        }

                        temp.AcceptChanges();
                        DataRow dr = temp.NewRow();

                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim();
                        dr["SUBSECTION"] = cboDeptParametersSubsection.Text.Trim();
                        dr["DEPARTMENT"] = cboDeptParametersDepartment.Text.Trim();
                        dr["KPF"] = cboDeptParametersKPF.Text.Trim();
                        dr["KPFPARAMETER"] = cboDeptParametersKPFParameters.Text.Trim();
                        dr["KPFPARAMETERDESC"] = txtDeptParametersDesc.Text.Trim();

                        temp.Rows.Add(dr);
                        //Create a total invalid delete.
                        string strDelete = " where Bussunit = '999'";
                        int rowindex = grdDeptParameters.CurrentCell.RowIndex;
                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "DeptParameters");

                        evaluateDeptParameters();
                        hideColumnsOfGrid("grdDeptParameters");
                        grdDeptParameters.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabHOD":
                    #region tabHOD

                    if (cboDeptParametersSubsection.Text.Trim().Length != 0 && cboDeptParametersKPF.Text.Trim().Length != 0 &&
                        cboDeptParametersKPFParameters.Text.Trim().Length != 0 && txtDeptParametersDesc.Text.Trim().Length != 0)
                    {

                        DataTable temp = new DataTable();
                        temp = DeptParameters.Copy();

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            temp.Rows[i].Delete();

                        }

                        temp.AcceptChanges();
                        DataRow dr = temp.NewRow();

                        dr["BUSSUNIT"] = BusinessLanguage.BussUnit.Trim();
                        dr["MININGTYPE"] = BusinessLanguage.MiningType.Trim();
                        dr["BONUSTYPE"] = BusinessLanguage.BonusType.Trim();
                        dr["SECTION"] = txtSelectedSection.Text.Trim().Trim();
                        dr["SUBSECTION"] = cboDeptParametersSubsection.Text.Trim();
                        dr["DEPARTMENT"] = cboDeptParametersDepartment.Text.Trim();
                        dr["KPF"] = cboDeptParametersKPF.Text.Trim();
                        dr["KPFPARAMETER"] = cboDeptParametersKPFParameters.Text.Trim();
                        dr["KPFPARAMETERDESC"] = txtDeptParametersDesc.Text.Trim();

                        temp.Rows.Add(dr);
                        //Create a total invalid delete.
                        string strDelete = " where Bussunit = '999'";
                        int rowindex = grdDeptParameters.CurrentCell.RowIndex;
                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "DeptParameters");

                        evaluateDeptParameters();
                        hideColumnsOfGrid("grdDeptParameters");
                        grdDeptParameters.FirstDisplayedScrollingRowIndex = rowindex;
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

            }
        }

        private string checkSQL(int intCounter, string strSQL)
        {
            if (intCounter > 0)
            {
                for (int i = 0; i <= intCounter - 1; i++)
                {
                    strSQL = strSQL.Trim() + ",'0'";
                }
                strSQL = strSQL.Trim() + ")";
            }
            else
            {
                strSQL = strSQL.Trim() + "')";
            }

            return strSQL;
        }

        private void UpdateClockedShifts()
        {
            clsMain.clsMain Base = new clsMain.clsMain();

            #region Extract dates
            //Load the section's first and last shift date
            DateTime dteFSH = dateTimePicker1.Value;
            DateTime dteLSH = dateTimePicker2.Value;

            //Load the clocked shifts' from and end date
            object tempObj = Base.ScalarQuery(Base.DBConnectionString, "select distinct datefrom from clockedshifts");
            string tempdte = (string)tempObj;
            string convdte = tempdte.Substring(0, 4) + "-" + tempdte.Substring(4, 2) + "-" + tempdte.Substring(6, 2);
            DateTime dteDateFrom = Convert.ToDateTime(convdte.Trim());


            tempObj = Base.ScalarQuery(Base.DBConnectionString, "select distinct dateend from clockedshifts");
            tempdte = (string)tempObj;
            convdte = tempdte.Substring(0, 4) + "-" + tempdte.Substring(4, 2) + "-" + tempdte.Substring(6, 2);
            DateTime dteDateEnd = Convert.ToDateTime(convdte.Trim());

            int intstart = dteDateFrom.Subtract(dteFSH).Days + 1;
            int intend = dteLSH.Subtract(dteDateFrom).Days + 2;

            #endregion

        }

        private void grdKPFCostLevel_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            DataTable temp = new DataTable();

            if (e.RowIndex < 0)
            {

            }
            else
            {

                txtKPFCostLevelCapLow.Text = grdKPFCostLevel["Cap_low", e.RowIndex].Value.ToString().Trim();
                txtKPFCostLevelCapHigh.Text = grdKPFCostLevel["Cap_high", e.RowIndex].Value.ToString().Trim();
                txtKPFCostLevelKPF.Text = grdKPFCostLevel["KPF", e.RowIndex].Value.ToString().Trim();
                txtKPFCostLevelKPFValue.Text = grdKPFCostLevel["KPFValue", e.RowIndex].Value.ToString().Trim();

                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = false;
                btnInsertRow.Enabled = false;

            }
        }

        private void extractPrimaryKey(DataTable p, string tablename)
        {
            //List Names contains the primary key columns of the selected table
            lstPrimaryKeyColumns.Clear();
            switch (tablename)
            {
                case "CALENDAR":
                    lstPrimaryKeyColumns = Base.listCalendarPrimaryKey;
                    break;

                case "BONUSSHIFTS":
                    lstPrimaryKeyColumns = Base.listBonusShiftsPrimaryKey;
                    break;

                case "PARTICIPANTS":
                    lstPrimaryKeyColumns = Base.listParticipantsPrimaryKey;
                    break;

                case "RATES":
                    lstPrimaryKeyColumns = Base.listRatesPrimaryKey;
                    break;

                case "FACTORS":
                    lstPrimaryKeyColumns = Base.listFactorsPrimaryKey;
                    break;

                case "CONFIGURATION":
                    lstPrimaryKeyColumns = Base.listConfigurationPrimaryKey;
                    break;


            }

            ////lstTableColumns contains all the column names of the table excluding "BUSSUNIT","MININGTYPE","BONUSTYPE","PERIOD")
            ////Do this extract on the table in memory, because much quicker.
            lstTableColumns.Clear();
            DataTable temp = p.Copy();
            deleteAllCalcColumnsFromTempTable(tablename, temp);

            if (temp.Columns.Count > 0)
            {
                foreach (DataColumn col in temp.Columns)
                {
                    if (col.ColumnName == "BUSSUNIT" || col.ColumnName == "MININGTYPE" || col.ColumnName == "BONUSTYPE" || col.ColumnName == "PERIOD")
                    {
                    }
                    else
                    {
                        lstTableColumns.Add(col.ColumnName.ToString().Trim());
                    }
                }
            }
        }


        private void tabInfo_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (txtSelectedSection.Text == "***")
            {
                MessageBox.Show("Please select a section.", "Information", MessageBoxButtons.OK);
            }
            else
            {
                btnInsertRow.Enabled = true;
                btnUpdate.Enabled = true;

                btnDeleteRow.Enabled = false;
                listBox1.Enabled = false;                                
                btnLoad.Enabled = false;
                dateTimePicker1.Enabled = false;                      
                dateTimePicker2.Enabled = false;                       
                btnPrint.Enabled = false;
                btnLock.Enabled = false;
                panelLock.BackColor = Color.Lavender;

                int intCount = checkLock(tabInfo.SelectedTab.Name);
                if (intCount > 0)
                {
                    btnLock.Text = "Unlock";
                }
                else
                {
                    btnLock.Text = "Lock";
                }

                switch (tabInfo.SelectedTab.Name)
                {
                    #region tabCalendar
                    case "tabCalendar":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnLoad.Enabled = true;
                        dateTimePicker1.Enabled = true;                 //HJ
                        dateTimePicker2.Enabled = true;                 //HJ
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        lstPrimaryKeyColumns.Clear();
                        extractPrimaryKey(Calendar, "CALENDAR");
                        break;
                    #endregion

                    #region tabDeptParameters
                    case "tabDeptParameters":

                        btnInsertRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        break;

                    #endregion

                    #region tabClockShifts
                    case "tabClockShifts":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnPrint.Enabled = true;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;
                    #endregion

                    #region tabLabour
                    case "tabLabour":

                        btnInsertRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        extractPrimaryKey(Labour, "BONUSSHIFTS");
                        break;
                    #endregion

                    #region tabMineParameters
                    case "tabMineParameters":

                        btnDeleteRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;

                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        evaluateMineParameters();
                        extractPrimaryKey(MineParameters, "MINEPARAMETERS");
                        break;
                    #endregion

                    #region tabParticipation
                    case "tabParticipation":

                        btnDeleteRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        btnUpdate.Enabled = true;
                        evaluateParticipation();

                        extractPrimaryKey(Participation, "PARTICIPANTION");
                        break;

                    #endregion

                    #region tabKPFCostLevel
                    case "tabKPFCostLevel":

                        btnDeleteRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        btnUpdate.Enabled = true;
                        evaluateKPFCostLevel();
                        break;

                    #endregion

                    #region tabSubsectionDept
                    case "tabSubSectionDept":

                        btnDeleteRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        btnUpdate.Enabled = true;
                        evaluateSubsectionDept();
                        break;

                    #endregion

                    #region tabHOD
                    case "tabHOD":


                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnDeleteRow.Enabled = true;
                        btnPrint.Enabled = true;
                        btnUpdate.Enabled = true;
                        evaluateHOD();
                        extractPrimaryKey(HOD, "HOD"); 
                        break;

                    #endregion

                    #region tabArtisans
                    case "tabArtisans":

                        btnDeleteRow.Enabled = false;
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;
                        btnUpdate.Enabled = true;
                        evaluateArtisans();

                        extractPrimaryKey(Artisans, "ARTISANS"); 
                        break;

                    #endregion

                    #region tabConfig
                    case "tabConfig":

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;

                        extractPrimaryKey(Configs, "CONFIGURATION");
                        break;

                    #endregion

                    #region tabEmplPen
                    case "tabEmplPen":

                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnLock.Enabled = true;
                        btnPrint.Enabled = true;

                        extractPrimaryKey(EmplPen, "EMPLOYEEPENALTY");
                        break;

                    #endregion

                    #region tabOffday
                    case "tabOffday":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        break;

                    #endregion

                    #region tabSelected
                    case "tabSelected":

                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        listBox1.Enabled = true;                            //HJ
                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabStatus

                    case "tabProcess":

                        evaluateStatus();
                        btnInsertRow.Enabled = false;
                        btnUpdate.Enabled = false;
                        btnDeleteRow.Enabled = false;
                        btnLoad.Enabled = false;
                        btnPrint.Enabled = false;
                        btnLock.Enabled = false;

                        panelInsert.BackColor = Color.Cornsilk;
                        panelUpdate.BackColor = Color.Cornsilk;
                        panelDelete.BackColor = Color.Cornsilk;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        break;

                    #endregion

                    #region tabRates
                    case "tabRates":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;

                        extractPrimaryKey(Rates, "RATES");

                        break;

                    #endregion

                    #region tabMonitor
                    case "tabMonitor":

                        btnDeleteRow.Enabled = true;
                        panelInsert.BackColor = Color.Lavender;
                        panelUpdate.BackColor = Color.Lavender;
                        panelDelete.BackColor = Color.Lavender;
                        panelPreCalcReport.BackColor = Color.Cornsilk;
                        btnPrint.Enabled = true;
                        btnLock.Enabled = true;
                        break;

                    #endregion

                }
            }

        }

        private int checkLock(string processToBeChecked)
        {
            //Lynx....LINQ
            DataTable contactTable = TB.getDataTable(TB.TBName);

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "Y"
                                          where locks.Field<string>("PROCESS").TrimEnd() == processToBeChecked
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          select locks;


            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();
            int intcount = query1.Count<DataRow>();

            return intcount;

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }

        private int checkLockInputProcesses()
        {

            IEnumerable<DataRow> query1 = from locks in Status.AsEnumerable()
                                          where locks.Field<string>("STATUS").TrimEnd() == "N"
                                          where locks.Field<string>("CATEGORY").TrimEnd() == "Input Process"
                                          select locks;

            int intcount = query1.Count<DataRow>();

            return intcount;

            //DataTable contacts1 = query1.CopyToDataTable<DataRow>();

        }


        private void grdEmplPen_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboEmplPenEmployeeNo.Text = grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtPenaltyValue.Text = grdEmplPen["PENALTYVALUE", e.RowIndex].Value.ToString().Trim();
                cboPenaltyInd.Text = grdEmplPen["PENALTYIND", e.RowIndex].Value.ToString().Trim();
                if (grdEmplPen["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim() == "XXXXXXXXXXXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                }
            }
            Cursor.Current = Cursors.Arrow;

        }

        private void grdLabour_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboBonusShiftsGang.Text = grdLabour["GANG", e.RowIndex].Value.ToString().Trim();
                txtEmployeeNo.Text = grdLabour["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtEmployeeName.Text = grdLabour["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsWageCode.Text = grdLabour["WAGECODE", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsResponseCode.Text = grdLabour["LINERESPCODE", e.RowIndex].Value.ToString().Trim();
                txtShifts.Text = grdLabour["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
                txtAwop.Text = grdLabour["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                txtStrikeShifts.Text = grdLabour["Q_SHIFTS", e.RowIndex].Value.ToString().Trim();


            }

            Cursor.Current = Cursors.Arrow;

        }

        private void grdConfigs_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                cboParameterName.Text = grdConfigs["PARAMETERNAME", e.RowIndex].Value.ToString().Trim();
                cboParm1.Text = grdConfigs["PARM1", e.RowIndex].Value.ToString().Trim();
                cboParm2.Text = grdConfigs["PARM2", e.RowIndex].Value.ToString().Trim();
                cboParm3.Text = grdConfigs["PARM3", e.RowIndex].Value.ToString().Trim();
                cboParm4.Text = grdConfigs["PARM4", e.RowIndex].Value.ToString().Trim();
                cboParm5.Text = grdConfigs["PARM5", e.RowIndex].Value.ToString().Trim();
                cboParm6.Text = grdConfigs["PARM6", e.RowIndex].Value.ToString().Trim();
                cboParm7.Text = grdConfigs["PARM7", e.RowIndex].Value.ToString().Trim();
            }
            Cursor.Current = Cursors.Arrow;
        }

        private void grdParticipation_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                txtParticipationSubsection.Text = grdParticipation["Section", e.RowIndex].Value.ToString().Trim();
                txtParticipationDepartment.Text = grdParticipation["Department", e.RowIndex].Value.ToString().Trim();
               // txtParticipGangType.Text = grdParticipation["GANGTYPE", e.RowIndex].Value.ToString().Trim();
               // cboParticipGangPerc.Text = grdParticipation["GANGTypeParticipation", e.RowIndex].Value.ToString().Trim();
                txtParticipEmplType.Text = grdParticipation["EMPLOYEETYPE", e.RowIndex].Value.ToString().Trim();
                cboParticipEmplPerc.Text = grdParticipation["EMPLOYEETypeParticipation", e.RowIndex].Value.ToString().Trim();
                txtParticipWage.Text = grdParticipation["WAGECODE", e.RowIndex].Value.ToString().Trim();
                cboParticipWagePerc.Text = grdParticipation["WAGECODEParticipation", e.RowIndex].Value.ToString().Trim();

                btnInsertRow.Enabled = false;
                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = false;
            }
        }

        #region AutoSize

        private void autoSizeGrid(DataGridView DG)
        {
            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader.ToString())
            {
                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            else
            {
                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                {
                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                }
                else
                {
                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.ColumnHeader.ToString())
                    {
                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                    }
                    else
                    {
                        if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                        {
                            DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                        }
                        else
                        {
                            if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.AllCells.ToString())
                            {
                                DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCellsExceptHeader;
                            }
                            else
                            {
                                if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.Fill.ToString())
                                {
                                    DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                                }
                                else
                                {
                                    if (DG.AutoSizeColumnsMode.ToString() == DataGridViewAutoSizeColumnsMode.DisplayedCells.ToString())
                                    {
                                        DG.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.ColumnHeader;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void grdActiveSheet_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdActiveSheet);
            }
        }

        private void grdCalendar_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdCalendar);
            }
        }

        private void grdClocked_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdClocked);
            }
        }

        private void grdLabour_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdLabour);
            }
        }

        private void grdHOD_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdHOD);
            }
        }

        private void grdKPF_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdDeptParameters);
            }
        }

        private void grdHOD_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdLabour);
            }
        }

        private void grdKPFCostLevel_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdKPFCostLevel);
            }
        }

        private void grdConfigs_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdConfigs);
            }
        }

        private void DoDataExtract()
        {
            connectToDB();
            TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);

        }
        #endregion

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {//xxxxxxxxxxxxxxxxxxxx
            string FormulaTableName = string.Empty;

            TB.TBName = (string)listBox1.SelectedItem;

            if (TB.TBName.Trim().ToUpper().Contains("EARN") && TB.TBName.Trim().ToUpper().Contains("20"))
            {
                FormulaTableName = TB.TBName.Trim().Substring(0, TB.TBName.Trim().ToUpper().IndexOf("20"));   //xxxxxxxxxxxxxxxxxx
            }
            else
            {
                FormulaTableName = TB.TBName;
            }

            TB.DBName = Base.DBName;

            connectToDB();
            cboColumnValues.Items.Clear();
            cboColumnNames.Items.Clear();
            cboColumnNames.Text = string.Empty;
            cboColumnValues.Text = string.Empty;

            List<string> lstColumnNames = General.getListOfColumnNames(Base.DBConnectionString, TB.TBName);

            foreach (string s in lstColumnNames)
            {
                cboColumnNames.Items.Add(s.Trim());
            }

            TB.ListOfSelectedTableColumns = lstColumnNames;

            DoDataExtract(strWhere);
            newDataTable = TB.getDataTable(TB.TBName);
            if (newDataTable == null)
            {
                DoDataExtract("");
                newDataTable = TB.getDataTable(TB.TBName);

            }
            //if newdatatable is still null
            if (newDataTable == null)
            {
                DoDataExtract("");
                newDataTable = TB.getDataTable(TB.TBName);

            }

            grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName + BusinessLanguage.Period.Trim(), FormulaTableName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }

            hideColumnsOfGrid("grdActiveSheet");
        }

        private void DoDataExtract(string Where)
        {
            connectToDB();
            if (Where.Trim().Length == 0)
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName);

            }
            else
            {
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, TB.TBName, Where);

            }

        }

        private void exportToExcel(string path, DataTable dt)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = path + "\\" + TB.TBName + ".xls";
                try
                {
                    StreamWriter SW = new StreamWriter(OPath);
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {
                            grid.RenderControl(HTMLWriter);
                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();
                    MessageBox.Show("Your spreadsheet was created at: " + OPath, "Information", MessageBoxButtons.OK);
                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void TBExport_Click(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void saveTheSpreadSheet()
        {
            string path = @"c:\" + TB.DBName + "\\" + TB.TBName;
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                DoDataExtract();
                DataTable outputTable = TB.getDataTable(TB.TBName);
                exportToExcel(path, outputTable);
                MessageBox.Show("Successfully Downloaded.", "Information", MessageBoxButtons.OK);

            }
            catch (Exception ee)
            {
                Console.WriteLine("The process failed: {0}", ee.ToString());
            }

            finally { }
        }

        private void grdActiveSheet_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            object intCount = Analysis.countcalcbyname(TB.DBName, TB.TBName, TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);
            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetails(TB.DBName, TB.TBName, TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS");
                DataTable dtFactors = TB.getDataTable("FACTORS");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        MessageBox.Show("SQL extract", "Not available to be tested", MessageBoxButtons.OK);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {
                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void loadDict(DataTable _datatable)
        {
            foreach (DataRow _row in _datatable.Rows)
            {
                string str = _row[0].ToString().Trim();
                if (dict.ContainsKey(str))
                {
                    dict.Remove(str);
                    dict.Add(str, _row[1].ToString().Trim());
                }
                else
                {
                    dict.Add(str, _row[1].ToString().Trim());
                }
            }
            dict.Remove("X");
            dict.Add("X", "0");

        }

        private void buildDisplaySQL(string strwhere, decimal decValue)
        {
            string strSQL = "";

            strSQL = "Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                         TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + TBFormulas.A + TBFormulas.B + TBFormulas.C + TBFormulas.D + TBFormulas.E + TBFormulas.F + TBFormulas.G + TBFormulas.H + " " + strwhere;
            strSQL = strSQL.Replace("#", "").Replace(":and:", "and").Replace(" from ", "\n from ").Replace(" and ", "\n and ").Replace(" where ", "\n where ");

            General.textTestSQL = strSQL;
            scrQuerySQL testsql = new scrQuerySQL();
            testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
            testsql.ShowDialog();

        }

        private void userProfile_Click(object sender, EventArgs e)
        {
            scrProfile userProfile = new scrProfile();
            userProfile.FormLoad(BusinessLanguage, BaseConn);
            userProfile.Show();
        }

        private void grantAccessToolStripMenuItem_Click(object sender, EventArgs e)
        {
            scrSecurity useraccess = new scrSecurity();
            useraccess.userAccessLoad(myConn, Base, TB, BusinessLanguage.Userid, strServerPath.ToString().ToUpper());
            useraccess.Show();
        }

        private void btn0_Click(object sender, EventArgs e)
        {

            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "0";

        }

        private void btn1_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "1";
        }

        private void btn2_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "2";
        }

        private void btn3_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "3";
        }

        private void btn4_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "4";
        }

        private void btn5_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "5";
        }

        private void btn6_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "6";
        }

        private void btn7_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "7";
        }

        private void btn8_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "8";
        }

        private void btn9_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = txtSearchEmpl.Text.Trim() + "9";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            txtSearchEmpl.Text = "";
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            DataTable searchEmpl = TB.createDataTableWithAdapter(Base.DBConnectionString, "Select * from ClockedShifts where employee_no like '%" + txtSearchEmpl.Text.Trim() + "%'");

            if (searchEmpl.Rows.Count > 0)
            {
                //amp
                string strLSH = Clocked.Rows[0]["LSH"].ToString().Trim();
                DateTime LSH = Convert.ToDateTime(strLSH);
                string Mnth = string.Empty;
                string Day = string.Empty;
                foreach (DataColumn dc in searchEmpl.Columns)
                {
                    if (dc.Caption.Substring(0, 3) == "DAY")
                    {
                        double d = Convert.ToDouble(dc.Caption.Substring(3).Trim());
                        string strTemp = Clocked.Rows[0]["FSH"].ToString().Trim();
                        DateTime temp = Convert.ToDateTime(strTemp);
                        temp = temp.AddDays(d);
                        if (temp > LSH)  //remember the days start at 0
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            searchEmpl.Columns[dc.Caption].ColumnName = "x" + Day + '-' + Mnth;
                        }
                        else
                        {
                            if (Convert.ToString(temp.Day).Length < 2)
                            {
                                Day = "0" + Convert.ToString(temp.Day);
                            }
                            else
                            {
                                Day = Convert.ToString(temp.Day);
                            }
                            if (Convert.ToString(temp.Month).Length < 2)
                            {
                                Mnth = "0" + Convert.ToString(temp.Month);
                            }
                            else
                            {
                                Mnth = Convert.ToString(temp.Month);
                            }
                            searchEmpl.Columns[dc.Caption].ColumnName = "d" + Day + '-' + Mnth;
                        }
                    }
                }
            }
            grdClocked.DataSource = searchEmpl;
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            grdClocked.DataSource = Clocked;
        }

        private void grdActiveSheet_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            //Get calc name
            this.Cursor = Cursors.WaitCursor;
            int columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
            int rownr = grdActiveSheet.CurrentCell.RowIndex;
            TBFormulas.CalcName = grdActiveSheet.Columns[columnnr].HeaderText;

            //Check if it is a calculated column
            object intCount = Analysis.countcalcbyname(TB.DBName, TB.TBName, TBFormulas.CalcName.Trim(), Base.AnalysisConnectionString);

            if ((int)intCount > 0)
            {
                //It is a calculated column.
                DataTable dtFormula = Analysis.GetCalcDetails(TB.DBName, TB.TBName, TBFormulas.CalcName, Base.AnalysisConnectionString);
                //Extract the formula details:
                decimal decValue = 0;
                try
                {
                    decValue = Convert.ToDecimal(grdActiveSheet.CurrentCell.Value);
                }
                catch
                {
                    decValue = 0;
                }

                //Extract Factors
                TB.extractDBTableIntoDataTable(Base.DBConnectionString, "FACTORS");
                DataTable dtFactors = TB.getDataTable("FACTORS");
                dict.Clear();
                loadDict(dtFactors);

                if (dtFormula.Rows.Count > 0)
                {
                    TBFormulas.A = dtFormula.Rows[0]["A"].ToString().Trim();
                    TBFormulas.B = dtFormula.Rows[0]["B"].ToString().Trim();
                    TBFormulas.C = dtFormula.Rows[0]["C"].ToString().Trim();
                    TBFormulas.D = dtFormula.Rows[0]["D"].ToString().Trim();
                    TBFormulas.E = dtFormula.Rows[0]["E"].ToString().Trim();
                    TBFormulas.F = dtFormula.Rows[0]["F"].ToString().Trim();
                    TBFormulas.G = dtFormula.Rows[0]["G"].ToString().Trim();
                    TBFormulas.H = dtFormula.Rows[0]["H"].ToString().Trim();
                    TBFormulas.I = dtFormula.Rows[0]["I"].ToString().Trim();
                    TBFormulas.J = dtFormula.Rows[0]["J"].ToString().Trim();
                    TBFormulas.TableFormulaCall = dtFormula.Rows[0]["FORMULA_CALL"].ToString().Trim();
                    decimal decA = 0;
                    decimal decB = 0;
                    decimal decC = 0;
                    decimal decD = 0;
                    decimal decE = 0;
                    decimal decF = 0;
                    decimal decG = 0;
                    decimal decH = 0;
                    decimal decI = 0;
                    decimal decJ = 0;

                    if (TBFormulas.TableFormulaCall.Contains("SQL"))
                    {
                        string strWhere = " ";
                        for (int i = 0; i < grdActiveSheet.Columns.Count - 1; i++)
                        {

                            strWhere = strWhere.Trim() + " and t1." + grdActiveSheet.Columns[i].HeaderText.Trim() + " = '" + (string)(grdActiveSheet[i, e.RowIndex].Value).ToString().Trim() + "'";

                        }

                        buildDisplaySQL(strWhere, decValue);
                    }
                    else
                    {
                        if (TBFormulas.CalcName.Contains("xx") || TBFormulas.TableFormulaCall.Contains("Concat"))
                        {
                        }
                        else
                        {
                            if (grdActiveSheet.Columns.Contains(TBFormulas.A))
                            {
                                decA = Convert.ToDecimal(grdActiveSheet[TBFormulas.A, rownr].Value);
                            }
                            else
                                if (dict.ContainsKey(TBFormulas.A))
                                {
                                    decA = Convert.ToDecimal(dict[TBFormulas.A]);
                                }
                                else
                                {
                                    decA = 9999;
                                }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.B))
                            {
                                decB = Convert.ToDecimal(grdActiveSheet[TBFormulas.B, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.B))
                                {
                                    decB = Convert.ToDecimal(dict[TBFormulas.B]);
                                }
                                else
                                {
                                    decB = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.C))
                            {
                                decC = Convert.ToDecimal(grdActiveSheet[TBFormulas.C, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.C))
                                {
                                    decC = Convert.ToDecimal(dict[TBFormulas.C]);
                                }
                                else
                                {
                                    decC = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.D))
                            {
                                decD = Convert.ToDecimal(grdActiveSheet[TBFormulas.D, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.D))
                                {
                                    decD = Convert.ToDecimal(dict[TBFormulas.D]);
                                }
                                else
                                {
                                    decD = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.E))
                            {
                                decE = Convert.ToDecimal(grdActiveSheet[TBFormulas.E, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.E))
                                {
                                    decE = Convert.ToDecimal(dict[TBFormulas.E]);
                                }
                                else
                                {
                                    decE = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.F))
                            {
                                decF = Convert.ToDecimal(grdActiveSheet[TBFormulas.F, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.F))
                                {
                                    decF = Convert.ToDecimal(dict[TBFormulas.F]);
                                }
                                else
                                {
                                    decF = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.G))
                            {
                                decG = Convert.ToDecimal(grdActiveSheet[TBFormulas.G, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.G))
                                {
                                    decG = Convert.ToDecimal(dict[TBFormulas.G]);
                                }
                                else
                                {
                                    decG = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.H))
                            {
                                decH = Convert.ToDecimal(grdActiveSheet[TBFormulas.H, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.H))
                                {
                                    decH = Convert.ToDecimal(dict[TBFormulas.H]);
                                }
                                else
                                {
                                    decH = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.I))
                            {
                                decI = Convert.ToDecimal(grdActiveSheet[TBFormulas.I, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.I))
                                {
                                    decI = Convert.ToDecimal(dict[TBFormulas.I]);
                                }
                                else
                                {
                                    decI = 9999;
                                }
                            }

                            if (grdActiveSheet.Columns.Contains(TBFormulas.J))
                            {
                                decJ = Convert.ToDecimal(grdActiveSheet[TBFormulas.J, rownr].Value);
                            }
                            else
                            {
                                if (dict.ContainsKey(TBFormulas.J))
                                {
                                    decJ = Convert.ToDecimal(dict[TBFormulas.J]);
                                }
                                else
                                {
                                    decJ = 9999;
                                }
                            }

                            MessageBox.Show("Database Name:     " + TB.DBName + '\n' + "Table Name:           " + TB.TBName + '\n' + "Calculation Name:   " +
                            TBFormulas.CalcName + "        Formula Name:   " + TBFormulas.TableFormulaCall + "   =   " + decValue + '\n' + '\n' + '\n' + "A =             " +
                            TBFormulas.A + "   =   " + decA + '\n' + "B =             " + TBFormulas.B + "   =   " + decB + '\n' + "C =             " +
                            TBFormulas.C + "   =   " + decC + '\n' + "D =             " +
                            TBFormulas.D + "   =   " + decD + '\n' + "E =             " +
                            TBFormulas.E + "   =   " + decE + '\n' + "F =             " +
                            TBFormulas.F + "   =   " + decF + '\n' + "G =             " +
                            TBFormulas.G + "   =   " + decG + '\n' + "H =             " +
                            TBFormulas.H + "   =   " + decH + '\n' + "I  =            " +
                            TBFormulas.I + "   =    " + decI + '\n' + "J  =            " +
                            TBFormulas.J + "   =    " + decJ, "FORMULA DETAILS - of selected value: ---------------------------------------------------->        ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }

                else
                {
                    this.Cursor = Cursors.Arrow;
                    MessageBox.Show("Calculation does not exist anymore. Delete the column.", "ERROR", MessageBoxButtons.OK);
                }
            }
            this.Cursor = Cursors.Arrow;
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            if (listBox2.SelectedIndex >= 0)
            {
                this.Cursor = Cursors.WaitCursor;
                txtSelectedSection.Text = listBox2.SelectedItem.ToString().Trim();
                Base.Section = txtSelectedSection.Text.Trim();    //xxxxxxxxxxxxxxxxxx

                //Start Threads
                Shared.evaluateDataTable(Base, "BONUSSHIFTS");
                Shared.evaluateDataTable(Base, "CLOCKEDSHIFTS");
                Shared.extractListOfTableNames(Base);

                int intRowPosition = 0;
                for (int i = 0; i <= Calendar.Rows.Count - 1; i++)
                {
                    if (Calendar.Rows[i]["SECTION"].ToString().Trim() == txtSelectedSection.Text.Trim() &&
                        Calendar.Rows[i]["PERIOD"].ToString().Trim() == BusinessLanguage.Period.Trim())
                    {
                        intRowPosition = i;
                    }
                }


                label15.Text = listBox2.SelectedItem.ToString().Trim();
                label30.Text = BusinessLanguage.Period;
                strWhere = "where section = '" + listBox2.SelectedItem.ToString().Trim() +
                           "' and period = '" + BusinessLanguage.Period + "'";     //xxxxxxxxxxxxxxxxxx

                evaluateStatus();

                evaluateDeptParameters();
                evaluateParticipation();
                evaluateEmployeePenalties();
                evaluateSubsectionDept();
                evaluateKPFCostLevel();               
                evaluateMineParameters();
                evaluateHOD();
                evaluateArtisans();          
                evaluateRates();
                extractMeasuringDates();
                extractDBTableNames(listBox1);
            }
            this.Cursor = Cursors.Arrow;


        }

        private void extractMeasuringDates()
        {

            IEnumerable<DataRow> query1 = from locks in Calendar.AsEnumerable()
                                          where locks.Field<string>("SECTION").TrimEnd() == txtSelectedSection.Text.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();
            dateTimePicker1.Value = Convert.ToDateTime(temp.Rows[0]["FSH"].ToString().Trim());
            dateTimePicker2.Value = Convert.ToDateTime(temp.Rows[0]["LSH"].ToString().Trim());

        }

        private void btnEmployeeCalc_Click(object sender, EventArgs e)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        private void dataSort_Click(object sender, EventArgs e)
        {

        }

        private void DataPrintCrewPrint_Click(object sender, EventArgs e)
        {

        }

        private void btnUpdate_Click_1(object sender, EventArgs e)
        {
            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabSubSectionDept":
                    #region tabSubsectionDept


                    if (cboSubsection.Text.Trim().Length != 0 && cboDepartment.Text.Trim().Length != 0 &&
                        cboHODModel.Text.Trim().Length != 0 && cboCostLevel.Text.Trim().Length != 0)
                    {
                        intRow = grdSubsectionDept.CurrentCell.RowIndex;
                        intColumn = grdSubsectionDept.CurrentCell.ColumnIndex;

                        string strCostLevel = string.Empty;

                        if (cboCostLevel.Text.Contains("-"))
                        {
                            strCostLevel = cboCostLevel.Text.Substring(0, cboCostLevel.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strCostLevel = cboCostLevel.Text.Trim();
                        }


                        strSQL = "BEGIN transaction; Update SubsectionDept set " +
                                 " Subsection = '" + cboSubsection.Text.Trim() + "', Department = '" + cboDepartment.Text.Trim() +
                                 "', HODModel = '" + cboHODModel.Text.Trim() +
                                 "', Costlevel = '" + strCostLevel +
                                 "' Where Subsection = '" + grdSubsectionDept["Subsection", intRow].Value.ToString().Trim() +
                                 "' and HODModel = '" + grdSubsectionDept["HODModel", intRow].Value.ToString().Trim() +
                                 "' and Costlevel = '" + grdSubsectionDept["CostLevel", intRow].Value.ToString().Trim() +
                                 "' and Department = '" + grdSubsectionDept["Department", intRow].Value.ToString() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdSubsectionDept["SUBSECTION", intRow].Value = cboSubsection.Text.Trim();
                        grdSubsectionDept["DEPARTMENT", intRow].Value = cboDepartment.Text.Trim();
                        grdSubsectionDept["HODMODEL", intRow].Value = cboHODModel.Text.Trim();
                        grdSubsectionDept["COSTLEVEL", intRow].Value = strCostLevel;

                        for (int i = 0; i <= grdSubsectionDept.Columns.Count - 1; i++)
                        {
                            grdSubsectionDept[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        grdSubsectionDept.FirstDisplayedScrollingRowIndex = intRow;
                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdSubsectionDept[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("KPFCOSTLEVEL", "U - Update", s, dictGridValues[s], grdSubsectionDept[s, intRow].Value.ToString().Trim());

                            }

                        }


                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabMineParameters":
                    #region tabMineParameters


                    if (txtCost_Actual.Text.Trim().Length != 0 && txtCost_Planned.Text.Trim().Length != 0 &&
                        txtTonsPerTEC_Planned.Text.Trim().Length != 0 && txtTonsPerTEC_Actual.Text.Trim().Length != 0)
                    {
                        intRow = grdMineParameters.CurrentCell.RowIndex;
                        intColumn = grdMineParameters.CurrentCell.ColumnIndex;

                        strSQL = "BEGIN transaction; Update MineParameters set " +
                                 " Cost_Actual = '" + txtCost_Actual.Text.Trim() + "', Cost_Planned = '" + txtCost_Planned.Text.Trim() +
                                 "', TonsPerTEC_Actual = '" + txtTonsPerTEC_Actual.Text.Trim() +
                                 "', TonsPerTEC_Planned = '" + txtTonsPerTEC_Planned.Text.Trim() +
                                 "', Safety_Actual = '" + txtMineSafety_Actual.Text.Trim() +
                                 "' Where section = '" + grdMineParameters["Section", intRow].Value.ToString().Trim() +
                                 "' and TonsPerTEC_Actual = '" + grdMineParameters["TonsPerTEC_Actual", intRow].Value.ToString().Trim() +
                                 "' and TonsPerTEC_Planned = '" + grdMineParameters["TonsPerTEC_Planned", intRow].Value.ToString().Trim() +
                                 "' and Safety_Actual = '" + grdMineParameters["Safety_Actual", intRow].Value.ToString().Trim() +
                                 "' and Cost_Planned = '" + grdMineParameters["Cost_Planned", intRow].Value.ToString().Trim() +
                                 "' and Cost_Actual = '" + grdMineParameters["Cost_Actual", intRow].Value.ToString() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdMineParameters["Cost_Actual", intRow].Value = txtCost_Actual.Text.Trim();
                        //grdMineParameters["Cost_Planned", intRow].Value = txtCost_Planned.Text.Trim();
                        //grdMineParameters["Production_Actual", intRow].Value = txtProdActual.Text.Trim();
                        //grdMineParameters["Production_Planned", intRow].Value = txtProdPlanned.Text.Trim();
                        //grdMineParameters["Gold_Actual", intRow].Value = txtGoldActual.Text.Trim();
                        //grdMineParameters["Gold_Planned", intRow].Value = txtGoldPlanned.Text.Trim();
                        grdMineParameters["TonsPerTEC_Actual", intRow].Value = txtTonsPerTEC_Actual.Text.Trim();
                        grdMineParameters["TonsPerTEC_Planned", intRow].Value = txtTonsPerTEC_Planned.Text.Trim();
                        grdMineParameters["Safety_Actual", intRow].Value = txtMineSafety_Actual.Text.Trim();

                        for (int i = 0; i <= grdMineParameters.Columns.Count - 1; i++)
                        {
                            grdMineParameters[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        grdMineParameters.FirstDisplayedScrollingRowIndex = intRow;

                        
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    evaluateMineParameters();

                    break;
                    #endregion

                case "tabKPFCostLevel":
                    #region tabKPFCostLevel

                    if (txtKPFCostLevelCapLow.Text.Trim().Length != 0 && txtKPFCostLevelCapHigh.Text.Trim().Length != 0 &&
                        txtKPFCostLevelKPF.Text.Trim().Length != 0 && txtKPFCostLevelKPFValue.Text.Trim().Length != 0)
                    {
                        intRow = grdKPFCostLevel.CurrentCell.RowIndex;
                        intColumn = grdKPFCostLevel.CurrentCell.ColumnIndex;

                        string strCostLevel = string.Empty;

                        strSQL = "BEGIN transaction; Update KPFCostLevel set " +
                                 " Cap_Low = '" + txtKPFCostLevelCapLow.Text.Trim() + "', Cap_High = '" + txtKPFCostLevelCapHigh.Text.Trim() +
                                 "', KPFValue = '" + txtKPFCostLevelKPFValue.Text.Trim() +
                                 "' Where KPFValue = '" + grdKPFCostLevel["KPFValue", intRow].Value.ToString().Trim() +
                                 "' and KPF = '" + grdKPFCostLevel["KPF", intRow].Value.ToString().Trim() +
                                 "' and Cap_Low = '" + grdKPFCostLevel["Cap_Low", intRow].Value.ToString().Trim() +
                                 "' and Cap_High = '" + grdKPFCostLevel["Cap_High", intRow].Value.ToString().Trim() +
                                 "' and CostLevel = '" + grdKPFCostLevel["CostLevel", intRow].Value.ToString() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdKPFCostLevel["KPFValue", intRow].Value = txtKPFCostLevelKPFValue.Text.Trim();
                        grdKPFCostLevel["CAP_LOW", intRow].Value = txtKPFCostLevelCapLow.Text.Trim();
                        grdKPFCostLevel["CAP_HIGH", intRow].Value = txtKPFCostLevelCapHigh.Text.Trim();

                        for (int i = 0; i <= grdSubsectionDept.Columns.Count - 1; i++)
                        {
                            grdKPFCostLevel[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        grdKPFCostLevel.FirstDisplayedScrollingRowIndex = intRow;

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdKPFCostLevel[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("KPFCOSTLEVEL", "U - Update", s, dictGridValues[s], grdKPFCostLevel[s, intRow].Value.ToString().Trim());

                            }

                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabLabour":
                    #region tabLabour


                    if (txtEmployeeNo.Text.Trim().Length > 0
                        && txtEmployeeName.Text.Trim().Length > 0
                        && cboBonusShiftsGang.Text.Trim().Length > 0
                        && cboBonusShiftsWageCode.Text.Trim().Length > 0
                        && cboBonusShiftsResponseCode.Text.Trim().Length > 0
                        && txtShifts.Text.Trim().Length > 0
                        && txtAwop.Text.Trim().Length > 0
                        && txtStrikeShifts.Text.Trim().Length > 0)
                    {

                        intRow = grdLabour.CurrentCell.RowIndex;
                        string strWagecode = Convert.ToString(grdLabour["WAGECODE", intRow].Value);
                        string strEmployeeName = Convert.ToString(grdLabour["EMPLOYEE_NAME", intRow].Value);
                        string strGang = Convert.ToString(grdLabour["GANG", intRow].Value);
                        string strResponseCo = Convert.ToString(grdLabour["LINERESPCODE", intRow].Value);
                        string strShiftsWorked = Convert.ToString(grdLabour["SHIFTS_WORKED", intRow].Value);
                        string strAwops = Convert.ToString(grdLabour["AWOP_SHIFTS", intRow].Value);
                        string strStrikes = Convert.ToString(grdLabour["Q_SHIFTS", intRow].Value);
                        string strSubsection = string.Empty;
                        string strDepartment = string.Empty;
                        string strHODModel = string.Empty;

                        //Extract the updated employee's department and subsection
                        DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, cboBonusShiftsGang.Text.Trim());
                        if (sub.Rows.Count > 0)
                        {
                            strSubsection = sub.Rows[0]["SUBSECTION"].ToString().Trim();
                            strDepartment = sub.Rows[0]["DEPARTMENT"].ToString().Trim();
                            strHODModel = sub.Rows[0]["HODMODEL"].ToString().Trim();
                        }
                        else
                        {
                            strSubsection = "UNKNOWN";
                            strDepartment = "UNKNOWN";
                            strHODModel = "UNKNOWN";

                        }

                        //===================================================================================
                        //Update the HOD table aswell.

                        strSQL = "Update bonusshifts set wagecode = '" + cboBonusShiftsWageCode.Text.Trim() +
                                 "' , Gang = '" + cboBonusShiftsGang.Text.Trim() +
                                 "' , Linerespcode = '" + cboBonusShiftsResponseCode.Text.Trim() +
                                 "' , Shifts_Worked = '" + txtShifts.Text.Trim() +
                                 "' , Awop_Shifts = '" + txtAwop.Text.Trim() +
                                 "' , Q_Shifts = '" + txtStrikeShifts.Text.Trim() +
                                 "' , Subsection = '" + strSubsection +
                                 "' , Department = '" + strDepartment +
                                 "' , HODModel = '" + strHODModel +
                                 "' where employee_no = '" + grdLabour["Employee_No", intRow].Value +
                                 "' and Linerespcode = '" + grdLabour["Linerespcode", intRow].Value +
                                 "' and Employee_name = '" + grdLabour["Employee_Name", intRow].Value +
                                 "' and WageCode = '" + grdLabour["WageCode", intRow].Value +
                                 "' and Gang = '" + grdLabour["Gang", intRow].Value + "';" +
                                 "Update HOD set Gang = '" + cboBonusShiftsGang.Text.Trim() +
                                 "' , Shifts_Worked = '" + txtShifts.Text.Trim() +
                                 "' , Awop_Shifts = '" + txtAwop.Text.Trim() +
                                 "' , Subsection = '" + strSubsection +
                                 "' , Department = '" + strDepartment +
                                 "' , HODModel = '" + strHODModel +
                                 "' where employee_no = '" + grdLabour["Employee_No", intRow].Value +
                                 "' and Employee_Name = '" + grdLabour["Employee_Name", intRow].Value +
                                 "' and Gang = '" + grdLabour["Gang", intRow].Value + "';";

                        TB.InsertData(Base.DBConnectionString, strSQL);

                        grdLabour["WAGECODE", intRow].Value = cboBonusShiftsWageCode.Text.Trim();
                        grdLabour["GANG", intRow].Value = cboBonusShiftsGang.Text.Trim();
                        grdLabour["LINERESPCODE", intRow].Value = cboBonusShiftsResponseCode.Text.Trim();
                        grdLabour["SHIFTS_WORKED", intRow].Value = txtShifts.Text.Trim();
                        grdLabour["AWOP_SHIFTS", intRow].Value = txtAwop.Text.Trim();
                        grdLabour["Q_SHIFTS", intRow].Value = txtStrikeShifts.Text.Trim();
                        grdLabour["SUBSECTION", intRow].Value = strSubsection;
                        grdLabour["DEPARTMENT", intRow].Value = strDepartment;
                        grdLabour["HODMODEL", intRow].Value = strHODModel;

                        for (int i = 0; i <= grdLabour.Columns.Count - 1; i++)
                        {
                            grdLabour[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdLabour[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("BONUSSHIFTS", "U - Update", s, dictGridValues[s], grdLabour[s, intRow].Value.ToString().Trim());

                            }

                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabDeptParameters":
                    #region tabDeptParameters

                    if (txtDeptParametersDesc.Text.Trim().Length != 0)
                    {

                        intRow = grdDeptParameters.CurrentCell.RowIndex;
                        intColumn = grdDeptParameters.CurrentCell.ColumnIndex;

                        strSQL = "BEGIN transaction; Update DeptParameters set KPFParameterDesc = '" + txtDeptParametersDesc.Text.Trim().ToUpper() +
                                 "', KPFParameter = '" + cboDeptParametersKPFParameters.Text.Trim() +
                                 "', KPF = '" + cboDeptParametersKPF.Text.Trim() + "'" +
                                 " Where KPFParameter = '" + grdDeptParameters["KPFParameter", intRow].Value.ToString().Trim() +
                                 "' and KPF = '" + grdDeptParameters["KPF", intRow].Value.ToString().Trim() +
                                 "' and subsection = '" + grdDeptParameters["subsection", intRow].Value.ToString().Trim() +
                                 "' and department = '" + grdDeptParameters["department", intRow].Value.ToString().Trim() +
                                 "';Commit Transaction;";

                        grdDeptParameters["KPFParameterDesc", intRow].Value = txtDeptParametersDesc.Text.Trim().ToUpper();
                        grdDeptParameters["KPFParameterDesc", intRow].Style.BackColor = Color.Lavender;
                        grdDeptParameters["KPFParameter", intRow].Value = txtDeptParametersDesc.Text.Trim();
                        grdDeptParameters["KPFParameter", intRow].Style.BackColor = Color.Lavender;
                        //grdDeptParameters["KPFParameterDesc", intRow].Value = txtDeptParametersDesc.Text.Trim();
                        //grdDeptParameters["KPFParameterDesc", intRow].Style.BackColor = Color.Lavender;

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        if (cboDeptParametersColumns.Text.Trim().Length > 0)
                        {
                            evaluateDeptParameters();
                            cboDeptParametersValues_SelectedIndexChanged("Method", null);
                        }
                        else
                        {
                            evaluateDeptParameters();
                        }

                        grdDeptParameters.FirstDisplayedScrollingRowIndex = intRow;

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdDeptParameters[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("DEPTPARAMETERS", "U - Update", s, dictGridValues[s], grdDeptParameters[s, intRow].Value.ToString().Trim());

                            }

                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabHOD":
                    #region tabHOD

                    if (txtTotalPlanned.Text != "100")
                    {
                        MessageBox.Show("Sum of TOTALS PLANNED must be 100.", "Information", MessageBoxButtons.OK);
                        txtTotalPlanned.BackColor = Color.Red;
                    }
                    else
                    {
                        txtTotalPlanned.BackColor = Color.AliceBlue;

                        if (txtItem1Actual.Text.Trim().Length != 0 && txtItem2Actual.Text.Trim().Length != 0 && txtItem3Actual.Text.Trim().Length != 0 &&
                                txtItem4Actual.Text.Trim().Length != 0 && txtItem5Actual.Text.Trim().Length != 0 && txtHODHODModel.Text.Trim().Length != 0 &&
                                txtItem1Planned.Text.Trim().Length != 0 && txtItem2Planned.Text.Trim().Length != 0 &&
                                txtItem3Planned.Text.Trim().Length != 0 && txtItem4Planned.Text.Trim().Length != 0 && txtItem5Planned.Text.Trim().Length != 0)
                        {
                            intRow = grdHOD.CurrentCell.RowIndex;
                            intColumn = grdHOD.CurrentCell.ColumnIndex;

                            if (txtTonsPerTEC_Actual.Text == "ToBeCalculated")
                            {
                                txtTonsPerTEC_Actual.Text = "0";
                            }

                            Int16 intHODLength = Convert.ToInt16(Convert.ToDouble(txtHODHODModel.Text.Trim().Length));
                            Int16 intInsertLength = Convert.ToInt16(Convert.ToDouble("10") - intHODLength);

                            strSQL = "BEGIN transaction; Update HOD set Item1_Actual = '" + txtItem1Actual.Text.Trim() +
                                     "', Item2_Actual = '" + txtItem2Actual.Text.Trim() +
                                     "', Item3_Actual = '" + txtItem3Actual.Text.Trim() +
                                     "', Item4_Actual = '" + txtItem4Actual.Text.Trim() +
                                     "', Item5_Actual = '" + txtItem5Actual.Text.Trim() +
                                     "', Item1_Planned = '" + txtItem1Planned.Text.Trim() +
                                     "', Item2_Planned = '" + txtItem2Planned.Text.Trim() +
                                     "', Item3_Planned = '" + txtItem3Planned.Text.Trim() +
                                     "', Item4_Planned = '" + txtItem4Planned.Text.Trim() +
                                     "', Item5_Planned = '" + txtItem5Planned.Text.Trim() +
                                     "', Subsection = '" + cboHODSubsection.Text.Trim() +
                                     "', Department = '" + cboHODDepartment.Text.Trim() +
                                     "', Safety_Actual = '" + txtSafetyA.Text.Trim() +
                                     "', Cost_Actual = '" + txtCostA.Text.Trim() +
                                     "', Cost_Planned = '" + txtCostP.Text.Trim() +
                                     "', Tons_Actual = '" + txtTonsPerTECA.Text.Trim() +
                                     "', Tons_Planned = '" + txtTonsPerTECP.Text.Trim() +
                                     "', Shifts_Worked = '" + txtHODShiftsWorked.Text.Trim() +
                                     "', Awop_Shifts = '" + txtHODAwops.Text.Trim() +
                                     "', HODModel = '" + txtHODHODModel.Text.Trim() +
                                     "', Designation = '" + cboHODDesignation.Text.Trim() +
                                      "', Employee_no = '" + txtHODEmployeeNo.Text.Trim() +
                                       "', Employee_name = '" + txtHODEmployeeName.Text.Trim() +
                                     "', Designation_Desc = '" + txtHODDesignationDesc.Text.Trim() +
                                     "', Gang = '" + txtHODHODModel.Text.Trim() + grdHOD["Gang", intRow].Value.ToString().Substring(intHODLength, intInsertLength) +
                                     "' Where Subsection = '" + grdHOD["Subsection", intRow].Value.ToString().Trim() +
                                     "' and Department = '" + grdHOD["Department", intRow].Value.ToString() +
                                     "' and Gang = '" + grdHOD["Gang", intRow].Value.ToString() +
                                     "' and Employee_no = '" + grdHOD["Employee_No", intRow].Value.ToString() +
                                     "' and Designation = '" + grdHOD["Designation", intRow].Value.ToString() +
                                     "' and HODModel = '" + grdHOD["HODModel", intRow].Value.ToString() +
                                     "';Commit Transaction;";




                            TB.InsertData(Base.DBConnectionString, strSQL);
                            grdHOD["Employee_name", intRow].Value = txtHODEmployeeName.Text.Trim();
                            grdHOD["Employee_no", intRow].Value = txtHODEmployeeNo.Text.Trim();
                            grdHOD["Item1_Actual", intRow].Value = txtItem1Actual.Text.Trim();
                            grdHOD["Item2_Actual", intRow].Value = txtItem2Actual.Text.Trim();
                            grdHOD["Item3_Actual", intRow].Value = txtItem3Actual.Text.Trim();
                            grdHOD["Item4_Actual", intRow].Value = txtItem4Actual.Text.Trim();
                            grdHOD["Item5_Actual", intRow].Value = txtItem5Actual.Text.Trim();
                            grdHOD["Item1_Planned", intRow].Value = txtItem1Planned.Text.Trim();
                            grdHOD["Item2_Planned", intRow].Value = txtItem2Planned.Text.Trim();
                            grdHOD["Item3_Planned", intRow].Value = txtItem3Planned.Text.Trim();
                            grdHOD["Item4_Planned", intRow].Value = txtItem4Planned.Text.Trim();
                            grdHOD["Item5_Planned", intRow].Value = txtItem5Planned.Text.Trim();
                            grdHOD["Subsection", intRow].Value = cboHODSubsection.Text.Trim();
                            grdHOD["Department", intRow].Value = cboHODDepartment.Text.Trim();
                            grdHOD["Safety_Actual", intRow].Value = txtSafetyA.Text.Trim();
                            grdHOD["Tons_Planned", intRow].Value = txtTonsPerTECP.Text.Trim();
                            grdHOD["Tons_Actual", intRow].Value = txtTonsPerTECA.Text.Trim();
                            grdHOD["Cost_Planned", intRow].Value = txtCostP.Text.Trim();
                            grdHOD["Cost_Actual", intRow].Value = txtCostA.Text.Trim();
                            grdHOD["Shifts_Worked", intRow].Value = txtHODShiftsWorked.Text.Trim();
                            grdHOD["Awop_Shifts", intRow].Value = txtHODAwops.Text.Trim();
                            grdHOD["HODModel", intRow].Value = txtHODHODModel.Text.Trim();
                            grdHOD["Gang", intRow].Value = txtHODHODModel.Text.Trim() + grdHOD["Gang", intRow].Value.ToString().Substring(intHODLength, intInsertLength);
                            grdHOD["Designation", intRow].Value = cboHODDesignation.Text.Trim();
                            grdHOD["Designation_desc", intRow].Value = txtHODDesignationDesc.Text.Trim();


                            for (int i = 0; i <= grdHOD.Columns.Count - 1; i++)
                            {
                                grdHOD[i, intRow].Style.BackColor = Color.Lavender;
                            }

                            //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                            foreach (string s in lstTableColumns)
                            {
                                if (dictGridValues[s] == grdHOD[s, intRow].Value.ToString().Trim())
                                {

                                }
                                else
                                {
                                    //Write out to audit log
                                    writeAudit("HOD", "U - Update", s, dictGridValues[s], grdHOD[s, intRow].Value.ToString().Trim());

                                }

                            }

                        }
                        else
                        {
                            MessageBox.Show("Invalid data.  Please check all input boxes.", "Error", MessageBoxButtons.OK);
                        }
                    }
                    break;
                    #endregion

                case "tabArtisan":
                    #region tabArtisan"

                    if (cboArtisanSubsection.Text.Trim().Length != 0 && cboArtisanDepartment.Text.Trim().Length != 0 &&
                        txtArtisanHODModel.Text.Trim().Length != 0 && txtArtisanShiftsWorked.Text.Trim().Length != 0 &&
                        txtArtisanAwops.Text.Trim().Length != 0 && txtArtisanHODModel.Text.Trim().Length != 0 &&
                        txtArtisanPPActual.Text.Trim().Length != 0)
                    {
                        intRow = grdArtisans.CurrentCell.RowIndex;
                        intColumn = grdArtisans.CurrentCell.ColumnIndex;

                        strSQL = "BEGIN transaction; Update Artisans set " +
                                 " Subsection = '" + cboArtisanSubsection.Text.Trim() +
                                 "', Department = '" + cboArtisanDepartment.Text.Trim() +
                                 "', HODModel = '" + txtArtisanHODModel.Text.Trim() +
                                 "', Shifts_Worked = '" + txtArtisanShiftsWorked.Text.Trim() +
                                 "', Awop_Shifts = '" + txtArtisanAwops.Text.Trim() +
                                 "', PERSONALPERFORMANCE_ACTUAL = '" + txtArtisanPPActual.Text.Trim() +
                                 "' Where Subsection = '" + grdArtisans["Subsection", intRow].Value.ToString().Trim() +
                                 "' and Department = '" + grdArtisans["Department", intRow].Value.ToString().Trim() +
                                 "' and Gang = '" + grdArtisans["Gang", intRow].Value.ToString().Trim() +
                                 "' and Employee_no = '" + grdArtisans["Employee_No", intRow].Value.ToString().Trim() +
                                 "' and HODModel = '" + grdArtisans["HODModel", intRow].Value.ToString().Trim() +
                                 "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);


                        grdArtisans["Subsection", intRow].Value = cboArtisanSubsection.Text.Trim();
                        grdArtisans["Department", intRow].Value = cboArtisanDepartment.Text.Trim();
                        grdArtisans["PERSONALPERFORMANCE_ACTUAL", intRow].Value = txtArtisanPPActual.Text.Trim();
                        grdArtisans["HODModel", intRow].Value = txtArtisanHODModel.Text.Trim();
                        grdArtisans["Awop_Shifts", intRow].Value = txtArtisanAwops.Text.Trim();
                        grdArtisans["Shifts_Worked", intRow].Value = txtArtisanShiftsWorked.Text.Trim();

                        for (int i = 0; i <= grdArtisans.Columns.Count - 1; i++)
                        {
                            grdArtisans[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdArtisans[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("ARTISANS", "U - Update", s, dictGridValues[s], grdArtisans[s, intRow].Value.ToString().Trim());

                            }

                        }



                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    //HJ
                    if (cboEmplPenEmployeeNo.Text.Trim().Length != 0 &&
                        txtPenaltyValue.Text.Trim().Length != 0 && cboPenaltyInd.Text.Trim().Length != 0)
                    {

                        intRow = grdEmplPen.CurrentCell.RowIndex;
                        intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                        if (cboEmplPenEmployeeNo.Text.Contains("-"))
                        {
                            strName = cboEmplPenEmployeeNo.Text.Substring(0, cboEmplPenEmployeeNo.Text.IndexOf("-")).Trim();
                        }
                        else
                        {
                            strName = cboEmplPenEmployeeNo.Text.Trim();
                        }

                        strSQL = "BEGIN transaction; Update EmployeePenalties set Period = '" + txtPeriod.Text.Trim() +
                                             "', Employee_No = '" + strName + "', PenaltyValue = '" + txtPenaltyValue.Text.Trim() +
                                             "', PenaltyInd = '" + cboPenaltyInd.Text.Trim() + "'" +
                                             " Where Section = '" + grdEmplPen["SECTION", intRow].Value.ToString().Trim() +
                                             "' and Period = '" + grdEmplPen["PERIOD", intRow].Value.ToString().Trim() +
                                             "' and Employee_No = '" + grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() +
                                             "' and PenaltyValue = '" + grdEmplPen["PENALTYVALUE", intRow].Value.ToString().Trim() +
                                             "' and PenaltyInd = '" + grdEmplPen["PENALTYIND", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXXXXXXXXXXX")
                        {
                            grdEmplPen["Section", intRow].Value = txtSelectedSection.Text.Trim();
                            grdEmplPen["Section", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["Period", intRow].Value = txtPeriod.Text.Trim();
                            grdEmplPen["Period", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["Employee_No", intRow].Value = cboEmplPenEmployeeNo.Text.Trim();
                            grdEmplPen["Employee_No", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["PenaltyValue", intRow].Value = txtPenaltyValue.Text.Trim();
                            grdEmplPen["PenaltyValue", intRow].Style.BackColor = Color.Lavender;
                            grdEmplPen["PenaltyInd", intRow].Value = cboPenaltyInd.Text.Trim();
                            grdEmplPen["PenaltyInd", intRow].Style.BackColor = Color.Lavender;

                            TB.InsertData(Base.DBConnectionString, strSQL);
                            //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                            foreach (string s in lstTableColumns)
                            {
                                if (dictGridValues[s] == grdEmplPen[s, intRow].Value.ToString().Trim())
                                {

                                }
                                else
                                {
                                    //Write out to audit log
                                    writeAudit("EmplPen", "U - Update", s, dictGridValues[s], grdEmplPen[s, intRow].Value.ToString().Trim());

                                }

                            }


                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabConfig":
                    #region tabConfiguration

                    //HJ
                    if (grdConfigs[0, intRow].Value.ToString().Trim() != "XXX")
                    {
                        if (cboParameterName.Text.Trim().Length != 0 && cboParm1.Text.Trim().Length != 0 &&
                            cboParm2.Text.Trim().Length != 0 && cboParm3.Text.Trim().Length != 0 &&
                            cboParm4.Text.Trim().Length != 0 && cboParm5.Text.Trim().Length != 0 &&
                            cboParm6.Text.Trim().Length != 0 && cboParm7.Text.Trim().Length != 0)
                        {

                            intRow = grdConfigs.CurrentCell.RowIndex;
                            intColumn = grdConfigs.CurrentCell.ColumnIndex;

                            InputBoxResult intresult = InputBox.Show("Password: ");

                            if (intresult.ReturnCode == DialogResult.OK)
                            {
                                if (intresult.Text.Trim() == "Moses")
                                {

                                    General.updateConfigsRecord(Base.BaseConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType, BusinessLanguage.BonusType,
                                    cboParameterName.Text.Trim(), cboParm1.Text.Trim(), cboParm2.Text.Trim(), cboParm3.Text.Trim(), cboParm4.Text.Trim(),
                                    cboParm5.Text.Trim(), cboParm6.Text.Trim(), cboParm7.Text.Trim(), grdConfigs["ParameterName", intRow].Value.ToString().Trim(),
                                    grdConfigs["Parm1", intRow].Value.ToString().Trim(), grdConfigs["Parm2", intRow].Value.ToString().Trim(),
                                    grdConfigs["Parm3", intRow].Value.ToString().Trim(), grdConfigs["Parm4", intRow].Value.ToString().Trim());
                                }
                                else
                                {
                                    MessageBox.Show("Invalid password", "Error", MessageBoxButtons.OK);
                                }
                            }

                            grdConfigs["ParameterName", intRow].Value = cboParameterName.Text.Trim();
                            grdConfigs["ParameterName", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm1", intRow].Value = cboParm1.Text.Trim();
                            grdConfigs["Parm1", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm2", intRow].Value = cboParm2.Text.Trim();
                            grdConfigs["Parm2", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm3", intRow].Value = cboParm3.Text.Trim();
                            grdConfigs["Parm3", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm4", intRow].Value = cboParm4.Text.Trim();
                            grdConfigs["Parm4", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm5", intRow].Value = cboParm5.Text.Trim();
                            grdConfigs["Parm5", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm6", intRow].Value = cboParm6.Text.Trim();
                            grdConfigs["Parm6", intRow].Style.BackColor = Color.Lavender;
                            grdConfigs["Parm7", intRow].Value = cboParm7.Text.Trim();
                            grdConfigs["Parm7", intRow].Style.BackColor = Color.Lavender;

                            //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                            foreach (string s in lstTableColumns)
                            {
                                if (dictGridValues[s] == grdConfigs[s, intRow].Value.ToString().Trim())
                                {

                                }
                                else
                                {
                                    //Write out to audit log
                                    writeAudit("CONFIGURATION", "U - Update", s, dictGridValues[s], grdConfigs[s, intRow].Value.ToString().Trim());

                                }

                            }

                        }
                        else
                        {
                            MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data.", "Error", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    //HJ
                    if (txtLowValue.Text.Trim().Length != 0 &&
                        txtHighValue.Text.Trim().Length != 0 && txtRate.Text.Trim().Length != 0)
                    {

                        InputBoxResult result = InputBox.Show("Password: ", "Rates Inputs are Password Protected!", "*", "0");

                        if (result.ReturnCode == DialogResult.OK)
                        {
                            if (result.Text.Trim() == "Moses")
                            {
                                intRow = grdRates.CurrentCell.RowIndex;
                                intColumn = grdRates.CurrentCell.ColumnIndex;

                                General.updateRatesRecord(myConn.ConnectionString, BusinessLanguage.BussUnit, txtMiningType.Text.Trim(),
                                                                 txtBonusType.Text.Trim(),
                                                                 txtPeriod.Text.ToString().Trim(), txtRateType.Text.Trim(), txtLowValue.Text.Trim(),
                                                                 txtHighValue.Text.Trim(), txtRate.Text.Trim(),
                                                                 grdRates["Low_Value", intRow].Value.ToString().Trim(), grdRates["High_Value", intRow].Value.ToString().Trim(),
                                                                 grdRates["Rate", intRow].Value.ToString().Trim());
                                Application.DoEvents();

                                MessageBox.Show("All calculations will becleared.  Recalculations have to be done.", "Information", MessageBoxButtons.OK);


                                grdRates["Low_Value", intRow].Value = txtLowValue.Text.Trim();
                                grdRates["Low_Value", intRow].Style.BackColor = Color.Lavender;
                                grdRates["High_Value", intRow].Value = txtHighValue.Text.Trim();
                                grdRates["High_Value", intRow].Style.BackColor = Color.Lavender;
                                grdRates["Rate", intRow].Value = txtRate.Text.Trim();
                                grdRates["Rate", intRow].Style.BackColor = Color.Lavender;

                                //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                                foreach (string s in lstTableColumns)
                                {
                                    if (dictGridValues[s] == grdRates[s, intRow].Value.ToString().Trim())
                                    {

                                    }
                                    else
                                    {
                                        //Write out to audit log
                                        writeAudit("RATES", "U - Update", s, dictGridValues[s], grdRates[s, intRow].Value.ToString().Trim());

                                    }

                                }

                            }
                            else
                            {
                                MessageBox.Show("Invalid Password.", "Information", MessageBoxButtons.OK);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion

                case "tabParticipation":
                    #region tabParticipation

                    if (cboParticipWagePerc.Text.Trim().Length != 0)
                    {
                        intRow = grdParticipation.CurrentCell.RowIndex;
                        intColumn = grdParticipation.CurrentCell.ColumnIndex;

                        TB.updateParticipation(Base.DBConnectionString, grdParticipation.Rows[intRow],
                            txtSelectedSection.Text.Trim(),
                            txtParticipEmplType.Text.Trim(),
                            cboParticipEmplPerc.Text.Trim(),
                            txtParticipGangType.Text.Trim(),
                            cboParticipGangPerc.Text.Trim(),
                            txtParticipWage.Text.Trim(),
                            cboParticipWagePerc.Text.Trim());

                        grdParticipation["WageCodeParticipation", intRow].Value = cboParticipWagePerc.Text.Trim();
                        grdParticipation["GangTypeParticipation", intRow].Value = cboParticipGangPerc.Text.Trim();
                        grdParticipation["EmployeeTypeParticipation", intRow].Value = cboParticipEmplPerc.Text.Trim();


                        for (int i = 0; i <= grdParticipation.Columns.Count - 1; i++)
                        {
                            grdParticipation[i, intRow].Style.BackColor = Color.Lavender;
                        }

                        //move updated values to the dictionary.  Compare updated values with the old values and write trigger.
                        foreach (string s in lstTableColumns)
                        {
                            if (dictGridValues[s] == grdParticipation[s, intRow].Value.ToString().Trim())
                            {

                            }
                            else
                            {
                                //Write out to audit log
                                writeAudit("PARTICIPATION", "U - Update", s, dictGridValues[s], grdParticipation[s, intRow].Value.ToString().Trim());

                            }

                        }


                    }

                    else
                    {
                        MessageBox.Show("Invalid data", "Error", MessageBoxButtons.OK);
                    }


                    break;
                    #endregion
            }
        }

        private void writeAudit(string tablename, string function, string fieldname, string oldValue, string newValue)
        {
            string PK = string.Empty;
            foreach (string key in dictPrimaryKeyValues.Keys)
            {
                PK = PK + "<" + key.Trim() + "=" + dictPrimaryKeyValues[key] + ">";
            }

            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "AUDIT");
            audit.Clear();

            DataRow dr = audit.NewRow();
            dr["Type"] = function.Substring(0, 1);
            dr["TableName"] = tablename;
            dr["PK"] = PK;
            dr["FieldName"] = fieldname;
            dr["OldValue"] = oldValue;
            dr["NewValue"] = newValue;
            dr["UpdateDate"] = DateTime.Today.ToLongDateString();
            dr["UserName"] = BusinessLanguage.Userid;

            audit.Rows.Add(dr);
            audit.AcceptChanges();

            TB.saveCalculations2(audit, Base.DBConnectionString, " where type = 'x'", "AUDIT");
        }


        private void deleteAllColumns(string Tablename)
        {
          //xxxxxxxxxxxxxxxxxxx  
            //Create the earnings table
            createTheFile(Tablename);

            //Add the calculation columns.
            createEarningsColumns(Tablename);

            List<string> lstColumnNames = new List<string>();

            //extract the latest data from the base file e.g. Ganglink, Bonusshifts and replace data in the earningsfile.

            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                           " where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

            //Give the tempory file a name
            tb.TableName = Tablename + "EARN" + BusinessLanguage.Period.Trim();

            if (Tablename.ToUpper() == "BONUSSHIFTS")
            {
                #region Remove columns starting with DAY from BONUSSHIFTS
                //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                        lstColumnNames.Add(dc.ColumnName.Trim());
                    }
                    else
                    {

                    }
                }

                foreach (string s in lstColumnNames)
                {
                    tb.Columns.Remove(s);
                    tb.AcceptChanges();
                }

                lstColumnNames.Clear();
                #endregion
            }

            //Save the data to be processed to the earnings table.
            TB.saveCalculations2(tb, Base.DBConnectionString, " where section = '" + txtSelectedSection.Text.Trim() + "'",
                                 tb.TableName.Trim());

            Application.DoEvents();
            //}
        }

        private void createTheFile(string Tablename)
        {
            //Check if earningstable exist - e.g. GangLinkEarn201108....if not...CREATE the table
            List<string> lstColumnNames = new List<string>();

            Int16 intCount = TB.checkTableExist(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period.Trim());

            if (intCount > 0)
            {
            }
            else
            {
                //CREATE the earnings table:  GanglinkEarn201108
                //Extract the table into a temp file from the datafile e.g. GANGLINK, BONUSSHIFTS, DRILLERS etc.

                DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename,
                               "where section = '" + txtSelectedSection.Text.Trim() + "' and period = '" + BusinessLanguage.Period + "'");

                //Give the tempory file a name
                tb.TableName = Tablename + "Earn" + BusinessLanguage.Period.Trim();

                if (Tablename.ToUpper() == "BONUSSHIFTS")
                {
                    #region Remove columns starting with DAY from BONUSSHIFTS
                    //Remove all the columns starting with "day" from temporary file, because BONUSSHIFTSEARN does not carry the DAY columns
                    foreach (DataColumn dc in tb.Columns)
                    {
                        if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                        {
                            lstColumnNames.Add(dc.ColumnName.Trim());
                        }
                        else
                        {

                        }
                    }

                    foreach (string s in lstColumnNames)
                    {
                        tb.Columns.Remove(s);
                        tb.AcceptChanges();
                    }

                    lstColumnNames.Clear();
                    #endregion
                }

                strSqlAlter.Remove(0, strSqlAlter.Length);

                //First create the base table.  Why, because all these columns should be NOT NULL.  
                //The Formulas SHOULD be NULL when created
                foreach (DataColumn dc in tb.Columns)
                {
                    if (dc.ColumnName.Substring(0, 3) == "DAY" && dc.ColumnName.Trim() != "DAYGANG")
                    {
                    }
                    else
                    {
                        lstColumnNames.Add(dc.ColumnName);
                    }
                }

                //Create the earningstable e.g. BONUSSHIFTSEARN201108T

                TB.createEarningsTable(Base.DBConnectionString, tb.TableName, Tablename, lstColumnNames);

            }
        }

        private void createEarningsColumns(string Tablename)
        {
            DataTable tb = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, Tablename + "EARN" + BusinessLanguage.Period);

            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(Base.DBName + BusinessLanguage.Period,
                                      Tablename + "EARN", Base.AnalysisConnectionString);

            foreach (DataRow row in tableformulas.Rows)
            {
                if (tb.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                }
                else
                {
                    strSqlAlter = strSqlAlter.Append(" ; Alter table " + Tablename + "EARN" + BusinessLanguage.Period + " add " +
                                                     row["CALC_NAME"].ToString().Trim() + " varchar(50) NULL");
                }
            }

            if (strSqlAlter.ToString().Trim().Length > 0)
            {
                StringBuilder bld = new StringBuilder();
                bld.Append("BEGIN transaction;" + strSqlAlter.ToString().Substring(1).Trim() + ";COMMIT transaction;");
                TB.InsertData(Base.DBConnectionString, bld.ToString().Trim());
                Application.DoEvents();
            }
            else
            {
            }
        }

        private void deleteAllCalcColumns(string Tablename)
        {
            strSqlAlter.Remove(0, strSqlAlter.Length);
            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
            }
        }

        private void deleteAllCalcColumns(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    TB.removeColumn(Base.DBConnectionString, Tablename, row["CALC_NAME"].ToString());
                }

            }
        }

        private void deleteAllCalcColumnsFromTempTable(string Tablename, DataTable Table)
        {
            //remove the column from the database.
            strSqlAlter.Remove(0, strSqlAlter.Length);

            DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, Tablename, Base.AnalysisConnectionString);
            foreach (DataRow row in tableformulas.Rows)
            {
                if (Table.Columns.Contains(row["CALC_NAME"].ToString().Trim()))
                {
                    Table.Columns.Remove(row["CALC_NAME"].ToString().Trim());
                }
            }

            Table.AcceptChanges();
        }

        private void Calcs(string tablename, string phasename, string Delete)
        {
            if (Delete == "Y")
            {
                deleteAllColumns(tablename);
            }

            TB.insertProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period, tablename + "EARN", phasename, txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "N", "N", (string)DateTime.Now.ToLongTimeString(), Convert.ToString(++intProcessCounter));

        }

        private void openTab(TabPage tp)
        {
            this.tabInfo.SelectedTab = tp;

            Application.DoEvents();

        }

        private void calcCrewsandGangs()
        {
            string strTableName = "";

            for (int i = 1; i <= 4; i++)
            {
                strTableName = "GangLink" + Convert.ToString(i).Trim();
                switch (i)
                {
                    case 1:
                        //btnPhase1.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 1", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "GangLink1", "Y");
                        break;

                    case 2:
                        //btnPhase1.BackColor = Color.LightGreen;
                        //btnPhase2.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 2", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "GangLink2", "Y");
                        break;

                    case 3:
                        //btnPhase2.BackColor = Color.LightGreen;
                        //btnPhase3.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 3", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Application.DoEvents();
                        Calcs("GangLink", "GangLink3", "Y");
                        break;

                    case 4:
                        //btnPhase3.BackColor = Color.LightGreen;
                        //btnPhase4.BackColor = Color.Orange;
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Base Calc Process", "Base Calc Process - Phase 4", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());
                        //Base.UpdateStatus(Base.DBConnectionString, "Y", "Header", "Base Calc Process", txtPeriod.Text.Trim(), txtSelectedSection.Text.Trim());

                        //Application.DoEvents();
                        Calcs("GangLink", "GangLink4", "Y");
                        break;
                }

                //executeFormulas(strTableName);
            }


            //btnPhase4.BackColor = Color.LightGreen;
            //Application.DoEvents();

        }

        private void calcCrewsandGangs(int counter)
        {
            string strTableName = "";

            for (int i = counter; i <= counter; i++)
            {
                strTableName = "GangLink" + Convert.ToString(i).Trim();
                switch (i)
                {
                    case 1:
                        Calcs("GangLink", "GangLink1", "Y");
                        break;

                    case 2:
                        Calcs("GangLink", "GangLink2", "N");
                        break;

                    case 3:
                        Calcs("GangLink", "GangLink3", "N");
                        break;

                    case 4:
                        Calcs("GangLink", "GangLink4", "N");
                        break;
                }


            }



            Application.DoEvents();

        }

        private void executeCostSheetFormulas(string TableName)
        {

            string strSQL = "BEGIN transaction; Delete from monitor ; commit transaction;";
            TB.InsertData(Base.DBConnectionString, strSQL);
            string strprevPeriod = TableName;
            strSQL = "BEGIN transaction; insert into monitor values('" + Base.DBName + "','" + strprevPeriod + "','N','0','" + txtSelectedSection.Text.Trim() + "','0','0'); commit transaction; ";
            TB.InsertData(Base.DBConnectionString, strSQL);

        }

        #region Open Tabs

        private void btnLockCalendar_Click(object sender, EventArgs e)
        {
            openTab(tabCalendar);
        }

        private void btnLockMineParameters_Click(object sender, EventArgs e)
        {
            openTab(tabMineParameters);
        }

        private void btnLockSurvey_Click(object sender, EventArgs e)
        {
            openTab(tabSubSectionDept);
        }

        private void btnLockBonusShifts_Click(object sender, EventArgs e)
        {
            openTab(tabLabour);
        }

        private void btnLockHOD_Click(object sender, EventArgs e)
        {
            openTab(tabHOD);
        }

        private void btnLockSubsectionDept_Click(object sender, EventArgs e)
        {
            openTab(tabSubSectionDept);
        }

        private void btnLockEmplPen_Click(object sender, EventArgs e)
        {
            openTab(tabEmplPen);
        }
        #endregion

        private void btnCrewLevel_Click(object sender, EventArgs e)
        {
            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {
                calcCrewsandGangs();

                evaluateStatus();
            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }
        }

        private void btnEmplTeamCalcHeader_Click(object sender, EventArgs e)
        {
            evaluateStatus();
        }

        private void saveXXXTeamShifts(DataTable TeamShifts)
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region TeamPrint
            foreach (DataRow rr in TeamShifts.Rows)
            {

                strSQL.Append("insert into TeamShifts values('" + rr["SECTION"].ToString().Trim() +
                              "','" + rr["CONTRACT"].ToString().Trim() + "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                              rr["GANG"].ToString().Trim() + "','" + rr["WAGECODE"].ToString().Trim() + "','" + rr["LINERESPCODE"].ToString().Trim() + "','" +
                              rr["EMPLOYEE_NO"].ToString().Trim() + "','" + rr["INITIALS"].ToString().Trim() + "','" +
                              rr["SURNAME"].ToString().Trim() + "','" + rr["REGISTER"].ToString().Trim() + "','" +
                              rr["DATEFROM"].ToString().Trim() + "','" + rr["EMPLOYEEPRODUCTIONBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEEDRESSINGBONUS"].ToString().Trim() + "','" + rr["EMPLOYEEAWOPPENALTYBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEEAWOPDRESSNGPENALTYBONUS"].ToString().Trim() + "','" + rr["EMPLOYEEHYDROBONUS"].ToString().Trim() + "','" +
                              rr["EMPLOYEESTOPEPROCESSBONUS"].ToString().Trim() + "')");



            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void saveXXXTeamProd(DataTable Teamprod)
        {
            StringBuilder strSQL = new StringBuilder();
            strSQL.Append("BEGIN transaction; ");

            #region TeamPrint
            foreach (DataRow rr in Teamprod.Rows)
            {
                //"CREATE TABLE TEAMPROD (SECTION char(50), CONTRACT Char(50), WORKPLACE Char(50), " +
                //    "GANG Char(50),WPNAME Char(50),WPSHIFTS Char(50),WPSHIFTSTOTAL Char(50), WPSQM Char(50), " +
                //    "WPFOOTWALL Char(50),WPSTOPEWIDTH Char(50),WPSTOPEWIDTHRATE Char(50), WPSTOPEWIDTHBONUS Char(50), " +
                //    "WPCONTRACTORBONUS Char(50),WPTOTALBONUS Char(50))";

                strSQL.Append("insert into TeamProd values('" + rr["SECTION"].ToString().Trim() + "','" + rr["CONTRACT"].ToString().Trim() +
                              "','" + rr["WORKPLACE"].ToString().Trim() + "','" +
                              rr["GANG"].ToString().Trim() + "','" + rr["CREWNO"].ToString().Trim() + "','" + rr["WPNAME"].ToString().Trim() + "','" + rr["WPSHIFTS"].ToString().Trim() +
                              "','" + rr["WPSHIFTSTOTAL"].ToString().Trim() + "','" + rr["WPSQM"].ToString().Trim() + "','" +
                              rr["WPFOOTWALL"].ToString().Trim() + "','" + rr["WPSTOPEWIDTH"].ToString().Trim() +
                              "','" + rr["WPSTOPEWIDTHRATE"].ToString().Trim() + "','" + rr["WPSTOPEWIDTHBONUS"].ToString().Trim() +
                              "','" + rr["WPCONTRACTORBONUS"].ToString().Trim() + "','" + rr["WPTOTALBONUS"].ToString().Trim() + "');");
            }

            strSQL.Append("Commit Transaction;");
            TB.InsertData(Base.DBConnectionString, Convert.ToString(strSQL));
            #endregion

        }

        private void WriteCSV(DataTable dt)
        {
            StreamWriter sw;
            string filePath = strServerPath + ":\\Crew.csv";


            sw = File.CreateText(filePath);

            try
            {
                // write the data in each row & column
                int intcounter = 0;
                foreach (DataRow row in dt.Rows)
                {
                    // recreate an empty Stringbuilder through each row iteration.
                    StringBuilder rowToWrite = new StringBuilder();

                    for (int counter = 0; counter <= dt.Columns.Count - 1; counter++)
                    {
                        if (intcounter == 0)
                        {
                            foreach (DataColumn column in dt.Columns)
                            {
                                //rowToWrite.Append("'" + column.ColumnName + "'");
                                rowToWrite.Append("'" + column.ColumnName + "'");
                            }
                            rowToWrite.Replace("''", "','");
                            rowToWrite.Replace("'", "");

                            rowToWrite.Append("\r\n");
                            sw.Write(rowToWrite);
                            rowToWrite.Remove(0, rowToWrite.Length);
                        }
                        intcounter = intcounter + 1;
                        rowToWrite.Append("'" + row[counter] + "'");
                    }

                    rowToWrite.Replace("''", "','");
                    rowToWrite.Replace("'", "");

                    rowToWrite.Append("\r\n");
                    sw.Write(rowToWrite);
                }
            }
            catch
            {
                //("An error occurred while attempting to build the CSV file. " + e.Message);
            }
            finally
            {
                sw.Close();
            }
        }

        private void btnDeleteRow_Click_1(object sender, EventArgs e)
        {

            int intRow = 0;
            int intColumn = 0;

            string strSQL = "";

            switch (tabInfo.SelectedTab.Name)
            {
                case "tabSubSectionDept":
                    #region tabSubsectionDept
                    intRow = grdSubsectionDept.CurrentCell.RowIndex;
                    intColumn = grdSubsectionDept.CurrentCell.ColumnIndex;

                    string strCostLevel = string.Empty;

                    if (cboCostLevel.Text.Contains("-"))
                    {
                        strCostLevel = cboCostLevel.Text.Substring(0, cboCostLevel.Text.IndexOf("-")).Trim();
                    }
                    else
                    {
                        strCostLevel = cboCostLevel.Text.Trim();
                    }

                    strSQL = "BEGIN transaction; Delete from SubsectionDept " +
                             " Where Section = '" + grdSubsectionDept["Section", intRow].Value.ToString().Trim() +
                             "' and Subsection = '" + grdSubsectionDept["Subsection", intRow].Value.ToString().Trim() +
                             "' and Department = '" + grdSubsectionDept["Department", intRow].Value.ToString().Trim() +
                             "' and HODModel = '" + grdSubsectionDept["HODModel", intRow].Value.ToString().Trim() +
                             "' and CostLevel = '" + strCostLevel + "';Commit Transaction;";

                    TB.InsertData(Base.DBConnectionString, strSQL);

                    evaluateSubsectionDept();

                    break;

                    #endregion

                //case "tabGangLinking":
                //    #region tabGangLink
                //    intRow = grdKPF.CurrentCell.RowIndex;
                //    intColumn = grdKPF.CurrentCell.ColumnIndex;

                //    if (grdKPF["GANG", intRow].Value.ToString().Trim() != "XXX")
                //    {

                //        strSQL = "BEGIN transaction; Delete from Ganglink " +
                //                 " Where Section = '" + grdKPF["Section", intRow].Value.ToString().Trim() +
                //                 "' and Period = '" + grdKPF["Period", intRow].Value.ToString().Trim() +
                //                 "' and Gang = '" + grdKPF["Gang", intRow].Value.ToString().Trim() +
                //                 "' and Workplace = '" + grdKPF["Workplace", intRow].Value.ToString().Trim() +
                //                 "' and SafetyInd = '" + grdKPF["SafetyInd", intRow].Value.ToString().Trim() +
                //                 "' and GangType = '" + grdKPF["GangType", intRow].Value.ToString().Trim() +
                //                 "';Commit Transaction;";

                //        TB.InsertData(Base.DBConnectionString, strSQL);
                //        evaluateGangLinking();
                //    }
                //    else
                //    {
                //        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                //    }
                //    break;

                //    #endregion

                case "tabAbnormal":
                //#region tabAbnormal
                //intRow = grdAbnormal.CurrentCell.RowIndex;
                //intColumn = grdAbnormal.CurrentCell.ColumnIndex;


                //if (grdAbnormal["ABNORMALVALUE", intRow].Value.ToString().Trim() != "XXX")
                //{

                //    strSQL = "BEGIN transaction; Delete from Abnormal " +
                //                 " Where Section = '" + grdAbnormal["Section", intRow].Value.ToString().Trim() +
                //                 "' and Period = '" + grdAbnormal["Period", intRow].Value.ToString().Trim() +
                //                 "' and Contract = '" + grdAbnormal["Contract", intRow].Value.ToString().Trim() +
                //                 "' and Workplace = '" + grdAbnormal["Workplace", intRow].Value.ToString().Trim() +
                //                 "' and AbnormalLevel = '" + grdAbnormal["AbnormalLevel", intRow].Value.ToString().Trim() +
                //                 "' and AbnormalType = '" + grdAbnormal["AbnormalType", intRow].Value.ToString().Trim() +
                //                 "' and AbnormalValue = '" + grdAbnormal["AbnormalValue", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                //    TB.InsertData(Base.DBConnectionString, strSQL);
                //    evaluateAbnormal();
                //}
                //else
                //{
                //    MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                //}
                //break;

                //#endregion

                case "tabEmplPen":
                    #region tabEmployeePenalty

                    intRow = grdEmplPen.CurrentCell.RowIndex;
                    intColumn = grdEmplPen.CurrentCell.ColumnIndex;

                    if (grdEmplPen["EMPLOYEE_NO", intRow].Value.ToString().Trim() != "XXX")
                    {

                        strSQL = "BEGIN transaction; Delete from EmployeePenalties " +
                                 " Where Section = '" + grdEmplPen["Section", intRow].Value.ToString().Trim() +
                                 "' and Period = '" + grdEmplPen["Period", intRow].Value.ToString().Trim() +
                                 "' and Employee_No = '" + grdEmplPen["Employee_no", intRow].Value.ToString().Trim() +
                                 "' and Workplace = '" + grdEmplPen["Workplace", intRow].Value.ToString().Trim() +
                                 "' and PenaltyInd = '" + grdEmplPen["PenaltyInd", intRow].Value.ToString().Trim() + "';Commit Transaction;";

                        TB.InsertData(Base.DBConnectionString, strSQL);
                        evaluateEmployeePenalties();
                    }
                    else
                    {
                        MessageBox.Show("This row cannot be deleted", "Information", MessageBoxButtons.OK);
                    }
                    break;

                    #endregion


                case "tabDeptParameters":
                    #region tabDeptParameters

                    intRow = grdDeptParameters.CurrentCell.RowIndex;
                    intColumn = grdDeptParameters.CurrentCell.ColumnIndex;

                    strSQL = "BEGIN transaction; Delete from DeptParameters " +
                             " Where Subsection = '" + grdDeptParameters["SubSection", intRow].Value.ToString().Trim() +
                             "' and department = '" + grdDeptParameters["Department", intRow].Value.ToString().Trim() +
                             "' and KPF = '" + grdDeptParameters["KPF", intRow].Value.ToString().Trim() +
                             "' and KPFParameter = '" + grdDeptParameters["KPFParameter", intRow].Value.ToString().Trim() +
                             "' and KPFParameterDesc = '" + grdDeptParameters["KPFParameterDesc", intRow].Value.ToString().Trim() +
                             "';Commit Transaction;";

                    TB.InsertData(Base.DBConnectionString, strSQL);
                    evaluateDeptParameters();

                    break;
                    #endregion
                case "tabHOD":
                    #region tabHOD

                    intRow = grdHOD.CurrentCell.RowIndex;
                    intColumn = grdHOD.CurrentCell.ColumnIndex;

                    strSQL = "BEGIN transaction; Delete from HOD " +
                               " Where Subsection = '" + grdHOD["Subsection", intRow].Value.ToString().Trim() +
                                "' and Department = '" + grdHOD["Department", intRow].Value.ToString() +
                                "' and Gang = '" + grdHOD["Gang", intRow].Value.ToString() +
                                     "' and Employee_no = '" + grdHOD["Employee_No", intRow].Value.ToString() +
                                     "' and Designation = '" + grdHOD["Designation", intRow].Value.ToString() +
                                     "' and HODModel = '" + grdHOD["HODModel", intRow].Value.ToString() +
                             "';Commit Transaction;";

                    TB.InsertData(Base.DBConnectionString, strSQL);
                    evaluateHOD();
                    break;

                    #endregion



            }
        }

        protected virtual void FrontDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteFullBeginTag("HTML");
            writer.WriteFullBeginTag("Head");
            writer.RenderBeginTag(System.Web.UI.HtmlTextWriterTag.Style);
            writer.Write("<!--");

            StreamReader sr = File.OpenText(strServerPath + ":\\koos.html");
            String input;
            while ((input = sr.ReadLine()) != null)
            {
                writer.WriteLine(input);
            }
            sr.Close();
            writer.Write("-->");
            writer.RenderEndTag();
            writer.WriteEndTag("Head");
            writer.WriteFullBeginTag("Body");
        }

        protected virtual void RearDecorator(System.Web.UI.HtmlTextWriter writer)
        {
            writer.WriteEndTag("Body");
            writer.WriteEndTag("HTML");
        }

        private void printHTML(DataTable dt, string TabName)
        {
            if (dt.Columns.Count > 0)
            {
                string OPath = "c:\\koos.html";

                try
                {

                    StreamWriter SW = new StreamWriter(OPath);
                    //StringWriter SW = new StringWriter();
                    System.Web.UI.HtmlTextWriter HTMLWriter = new System.Web.UI.HtmlTextWriter(SW);
                    System.Web.UI.WebControls.DataGrid grid = new System.Web.UI.WebControls.DataGrid();

                    grid.DataSource = dt;
                    grid.DataBind();

                    using (SW)
                    {
                        using (HTMLWriter)
                        {

                            HTMLWriter.WriteLine("HARMONY - Phakisa Mine - " + TabName);
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteLine("==============================");
                            HTMLWriter.WriteBreak();
                            HTMLWriter.WriteBreak();

                            grid.RenderControl(HTMLWriter);
                            //RearDecorator(HTMLWriter);

                        }
                    }

                    SW.Close();
                    HTMLWriter.Close();


                    System.Diagnostics.Process P = new System.Diagnostics.Process();
                    P.StartInfo.WorkingDirectory = strServerPath + ":\\Program Files\\Internet Explorer";
                    P.StartInfo.FileName = "IExplore.exe";
                    P.StartInfo.Arguments = "C:\\koos.html";
                    P.Start();
                    P.WaitForExit();


                }
                catch (Exception exx)
                {
                    MessageBox.Show("Could not create " + OPath.Trim() + ".  Create the directory first." + exx.Message, "Error", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Your spreadsheet could not be created.  No columns found in datatable.", "Error Message", MessageBoxButtons.OK);
            }

        }

        private void btnLoad_Click_1(object sender, EventArgs e)
        {
            if (listBox3.SelectedItems.Count == 0)
            {
                MessageBox.Show("Please select the number of measuring shifts", "Information", MessageBoxButtons.OK);
            }
            else
            {
                if (txtSelectedSection.Text.Trim().Length == 0)
                {
                    MessageBox.Show("Please select a section and the correct month measuring shifts for the section.", "Information", MessageBoxButtons.OK);
                }
                else
                {
                    string selectedSection = txtSelectedSection.Text.Trim();
                    string grdSection = grdCalendar["SECTION", intFiller].Value.ToString().Trim();
                    if (selectedSection == grdSection)
                    {
                        Base.updateCalendarRecord(Base.DBConnectionString, BusinessLanguage.BussUnit, txtMiningType.Text.Trim(),
                                                         txtBonusType.Text.Trim(), txtSelectedSection.Text.Trim(),
                                                         txtPeriod.Text.ToString().Trim(),
                                                         (Convert.ToDateTime(dateTimePicker1.Text)).ToString("yyyy-MM-dd"),
                                                         (Convert.ToDateTime(dateTimePicker2.Text)).ToString("yyyy-MM-dd"),
                                                         listBox3.SelectedItem.ToString().Trim());
                        Application.DoEvents();
                    }

                    else
                    {
                        MessageBox.Show("Selected section not the same as grid section.", "Informations", MessageBoxButtons.OK);
                    }

                    //Extract Calendar again and insert into 
                    Calendar = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Calendar");
                    grdCalendar.DataSource = Calendar;
                }
            }
        }


        private void label83_Click(object sender, EventArgs e)
        {
            extractDBTableNames(listBox1);
        }

        private void btnLockPaysend_Click(object sender, EventArgs e)
        {
            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();


        }


        private void btnPrint_Click_1(object sender, EventArgs e)
        {
            switch (tabInfo.SelectedTab.Name)
            {
                case "tabDeptParameters":
                    #region tabDeptParameters
                    DataTable dt = new DataTable();
                    dt = Base.extractPrintData(Base.DBConnectionString, "DeptParameters", strWhere);
                    deleteAllCalcColumns("DeptParameters", dt);
                    dt.AcceptChanges();
                    if (dt.Rows.Count > 0)
                    {

                        printHTML(dt, "DeptParameters");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabHOD":
                    #region tabDeptParameters
                    dt = new DataTable();
                    dt = Base.extractPrintData(Base.DBConnectionString, "HOD", strWhere);
                    deleteAllCalcColumns("HOD", dt);
                    dt.AcceptChanges();

                    if (dt.Columns.Contains("BUSSUNIT"))
                    {
                        dt.Columns.Remove("BUSSUNIT");
                    }

                    if (dt.Columns.Contains("MININGTYPE"))
                    {
                        dt.Columns.Remove("MININGTYPE");
                    }

                    if (dt.Columns.Contains("BONUSTYPE"))
                    {
                        dt.Columns.Remove("BONUSTYPE");
                    }

                    if (dt.Columns.Contains("DESIGNATION"))
                    {
                        dt.Columns.Remove("DESIGNATION");
                    }
                    if (dt.Columns.Contains("ITEM1_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM1_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM2_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM2_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM3_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM3_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM4_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM4_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM5_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM5_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM1_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM1_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM2_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM2_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM3_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM3_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM4_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM4_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM5_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM5_PLANNED");
                    }
                    if (dt.Rows.Count > 0)
                    {

                        printHTML(dt, "Head of Departments");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabParticipation":
                    #region tabDeptParameters
                    dt = new DataTable();
                    dt = Base.extractPrintData(Base.DBConnectionString, "PARTICIPATION", strWhere);
                    deleteAllCalcColumns("PARTICIPATION", dt);
                    dt.AcceptChanges();

                    if (dt.Columns.Contains("BUSSUNIT"))
                    {
                        dt.Columns.Remove("BUSSUNIT");
                    }

                    if (dt.Columns.Contains("MININGTYPE"))
                    {
                        dt.Columns.Remove("MININGTYPE");
                    }

                    if (dt.Columns.Contains("BONUSTYPE"))
                    {
                        dt.Columns.Remove("BONUSTYPE");
                    }

                    if (dt.Columns.Contains("DESIGNATION"))
                    {
                        dt.Columns.Remove("DESIGNATION");
                    }
                    if (dt.Columns.Contains("ITEM1_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM1_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM2_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM2_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM3_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM3_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM4_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM4_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM5_ACTUAL"))
                    {
                        dt.Columns.Remove("ITEM5_ACTUAL");
                    }
                    if (dt.Columns.Contains("ITEM1_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM1_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM2_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM2_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM3_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM3_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM4_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM4_PLANNED");
                    }
                    if (dt.Columns.Contains("ITEM5_PLANNED"))
                    {
                        dt.Columns.Remove("ITEM5_PLANNED");
                    }
                    if (dt.Rows.Count > 0)
                    {

                        printHTML(dt, "Head of Departments");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabAbnormal":
                    #region tabAbnormal

                    dt = Base.extractPrintData(Base.DBConnectionString, "Abnormal", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Abnormal");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabLabour":
                    #region tabLabour

                    dt = Base.extractPrintData(Base.DBConnectionString, "BonusShifts", strWhere);
                    deleteAllCalcColumns("BonusShifts", dt);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "BonusShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabSurvey":
                    #region tabSurvey

                    dt = Base.extractPrintData(Base.DBConnectionString, "Survey", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Survey");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabEmplPen":
                    #region tabEmployee Penalties

                    dt = Base.extractPrintData(Base.DBConnectionString, "EmployeePenalties", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "EmployeePenalties");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }
                    break;
                    #endregion

                case "tabOffday":
                    #region tabOffdays

                    dt = Base.extractPrintData(Base.DBConnectionString, "Offdays", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Offdays");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabCalendar":
                    #region tabCalendar

                    dt = Base.extractPrintData(Base.DBConnectionString, "Calendar", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Calendar");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabClockShifts":
                    #region tabClockShifts

                    dt = Base.extractPrintData(Base.DBConnectionString, "ClockedShifts", strWhere);
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "ClockedShifts");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabRates":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Rates", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Rates");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion

                case "tabMonitor":
                    #region tabRates

                    dt = Base.extractPrintData(Base.DBConnectionString, "Monitor", "");
                    if (dt.Rows.Count > 0)
                    {
                        printHTML(dt, "Monitor");
                    }
                    else
                    {
                        MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
                    }

                    break;
                    #endregion
            }
        }

        private void calcStopeData()
        {
            Base.Period = txtPeriod.Text.Trim();
            //Base.Period = "200909";

            SqlConnection stopeConn = Base.StopeConnection;
            stopeConn.Open();

            try
            {
                DataTable ContractTotals = TB.getContractCrewOfficialBonus(Base.StopeConnectionString, "STOPING", txtSelectedSection.Text.Trim());

                stopeConn.Close();

                TB.updateDSShiftbossCrewBonus(Base.DBConnectionString, ContractTotals);
            }
            catch { }
        }

        private void startCalcProcess()
        {
            this.Cursor = Cursors.WaitCursor;
            btnx.Visible = true;
            btnx.Enabled = true;
            btnx.Text = "Run";
            TB.deleteProcess(Base.AnalysisConnectionString, Base.DBName + BusinessLanguage.Period);
            //clear the monitor table
            TB.deleteAllExcept(Base.DBConnectionString, "Monitor");
            TB.deleteAllExcept(Base.DBConnectionString, "BONUSSHIFTS", " where wagecode in " + strWagecodes);
            Calcs("MINEPARAMETERS", "MINEPARAMETERSEarn5", "Y");
            Calcs("MINEPARAMETERS", "MINEPARAMETERSEarn10", "N");
            Calcs("MINEPARAMETERS", "MINEPARAMETERSEarn20", "N");
            Calcs("MINEPARAMETERS", "MINEPARAMETERSEarn30", "N");
            Calcs("HOD", "HODEarn5", "Y");
            Calcs("HOD", "HODEarn10", "N");
            Calcs("HOD", "HODEarn15", "N");
            Calcs("HOD", "HODEarn20", "N");
            Calcs("HOD", "HODEarn25", "N");
            Calcs("HOD", "HODEarn30", "N");
            Calcs("Artisans", "ArtisansEarn5", "Y");
            Calcs("Artisans", "ArtisansEarn10", "N");
            Calcs("Artisans", "ArtisansEarn15", "N");
            Calcs("Artisans", "ArtisansEarn20", "N");
            Calcs("BonusShifts", "BonusShiftsEarn5", "Y");
            Calcs("BonusShifts", "BonusShiftsEarn10", "N");
            Calcs("BonusShifts", "BonusShiftsEarn15", "N");
            Calcs("Exit", "Exit", "N");
            btnBaseCalcs.BackColor = Color.Orange; 
            btnHODCalcs.BackColor = Color.Orange;
            btnEmployeeCalcs.BackColor = Color.Orange;
            btnGangLevelCalcs.BackColor = Color.Orange;

            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "HODEarn5", txtSelectedSection.Text.Trim(), DateTime.Now.ToShortTimeString().ToString().Trim());
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "BonusShiftsEarn5", txtSelectedSection.Text.Trim(), DateTime.Now.ToShortTimeString().ToString().Trim());
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "BonusShiftsEarn15", txtSelectedSection.Text.Trim(), DateTime.Now.ToShortTimeString().ToString().Trim());
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "HODEarn30", txtSelectedSection.Text.Trim(), DateTime.Now.ToShortTimeString().ToString().Trim());
            TB.updateStatusFromArchive(Base.DBConnectionString, "N", "Exit", txtSelectedSection.Text.Trim(), BusinessLanguage.Period.Trim(), "");

            //Base.backupDatabase3(Base.DBConnectionString, Base.DBName, Base.BackupPath);
            this.Cursor = Cursors.Arrow;
        }

        private void btnBaseCalcsHeader_Click(object sender, EventArgs e)
        {
            int intCheckLocks = checkLockInputProcesses();

            if (intCheckLocks == 0)
            {
                //Check if the a calculator is currently running
                Int16 intCount1 = TB.checkTableExist(Base.DBConnectionString, "BonusShiftsEARN");
                Int16 intCount2 = TB.checkTableExist(Base.DBConnectionString, "ParticipantsEARN");
                Int16 intCount3 = TB.checkTableExist(Base.DBConnectionString, "SupportLinkEARN");
                Int16 intCount4 = TB.checkTableExist(Base.DBConnectionString, "DrillersEARN");
                Int16 intCount5 = TB.checkTableExist(Base.DBConnectionString, "MinersEARN");
                Int16 intCount6 = TB.checkTableExist(Base.DBConnectionString, "SectionEarningsEARN");

                if (intCount1 > 0 || intCount2 > 0 || intCount3 > 0 || intCount4 > 0 || intCount5 > 0 || intCount6 > 0)
                {
                    MessageBox.Show("A calculator is currently running for this bonus scheme: " + BusinessLanguage.MiningType +
                                    " " + BusinessLanguage.BonusType);
                }
                else
                {
                    startCalcProcess();

                }

            }
            else
            {
                MessageBox.Show("Finish all input processes first, before trying to process all.", "Informations", MessageBoxButtons.OK);
            }
        }


        private void grdActiveSheet_ColumnHeaderMouseClick_1(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int mouseY = MousePosition.Y;
                int mouseX = MousePosition.X;

                ctMenu.Show(this, new Point(mouseX, mouseY));

                columnnr = e.ColumnIndex;
                string columname = grdActiveSheet.Columns[columnnr].Name;
                DialogResult result = MessageBox.Show("Do you want to delete the column:  " + grdActiveSheet.Columns[columnnr].HeaderText + "?", "INFORMATION", MessageBoxButtons.YesNo);

                if (result == DialogResult.Yes)
                {
                    //columnnr = grdActiveSheet.CurrentCell.ColumnIndex;
                    //TB.removeColumn(Base.DBConnectionString, TB.TBName, grdActiveSheet.Columns[columnnr].HeaderText);
                    //DoDataExtract();
                    //grdActiveSheet.DataSource = TB.getDataTable(TB.TBName);
                    grdActiveSheet.Columns[columnnr].Visible = false;
                }
                else
                {
                    if (listBox1.SelectedItem.ToString().Trim() == "MONITOR")
                    {

                        string strSQL = "Begin transaction; Delete from monitor; commit transaction";
                        TB.InsertData(Base.DBConnectionString, strSQL);
                        Application.DoEvents();

                    }
                }
            }
        }

        private void grdCalendar_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                dateTimePicker1.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["FSH"].ToString().Trim());
                dateTimePicker2.Value = Convert.ToDateTime(Calendar.Rows[e.RowIndex]["LSH"].ToString().Trim());
                intFiller = e.RowIndex;
            }

        }

        private void grdRates_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;

            if (e.RowIndex < 0)
            {

            }
            else
            {
                if (grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim() == "XXX")
                {
                    btnUpdate.Enabled = false;
                    btnDeleteRow.Enabled = false;
                    btnInsertRow.Enabled = true;

                }
                else
                {
                    btnUpdate.Enabled = true;
                    btnDeleteRow.Enabled = true;
                    btnInsertRow.Enabled = true;
                }

                txtRateType.Text = grdRates["RATE_TYPE", e.RowIndex].Value.ToString().Trim();
                txtLowValue.Text = grdRates["LOW_VALUE", e.RowIndex].Value.ToString().Trim();
                txtHighValue.Text = grdRates["HIGH_VALUE", e.RowIndex].Value.ToString().Trim();
                txtRate.Text = grdRates["RATE", e.RowIndex].Value.ToString().Trim();
            }
        }



        private void payrollSend_Click(object sender, EventArgs e)
        {

            if (Base.DBTables.Contains("PAYROLL"))
            {
            }
            else
            {
                if (myConn.State == ConnectionState.Open)
                {
                }
                else
                {
                    myConn.Open();
                }

                //Create a table
                Int16 intCount = TB.checkTableExist(Base.DBConnectionString, "PAYROLL");
                if (intCount > 0)
                {
                }
                else
                {
                    TB.createPayrollTable(Base.DBConnectionString);
                }
            }

            scrPayroll paysend = new scrPayroll();
            string conn = myConn.ToString();
            string baseconn = BaseConn.ToString();
            string lang = BusinessLanguage.ToString();
            string tb = TB.ToString();
            string tbFormu = TBFormulas.ToString();
            paysend.PayrollSendLoad(myConn, BaseConn, BusinessLanguage, TB, TBFormulas, Base, txtSelectedSection.Text.Trim());
            paysend.Show();


        }

        private void emailInfo_Click(object sender, EventArgs e)
        {

        }

        private void basicGraph_Click(object sender, EventArgs e)
        {

        }

        private void drillDownGraph_Click(object sender, EventArgs e)
        {

        }

        private void dataFilter_Click(object sender, EventArgs e)
        {
            if (General.textTestSQL.ToString().Trim().Length > 0)
            {
                scrQuerySQL testsql = new scrQuerySQL();
                testsql.TestSQL(Base.DBConnection, General, Base.DBConnectionString);
                testsql.Show();
            }
            else
            {
                MessageBox.Show("No SQL to pass", "Information", MessageBoxButtons.OK);
            }
        }

        private void dataPrintTables_Click(object sender, EventArgs e)
        {

        }

        private void dataFormulasImportTable_Click(object sender, EventArgs e)
        {
            //Email error information to the standby person

            //OutlookIntegrationEx.MainForm ex = new OutlookIntegrationEx.MainForm();
            //ex.Show();

        }

        private void TBCreateSpreadsheet_Click(object sender, EventArgs e)
        {
            try
            {
                if (openDialog.ShowDialog() != DialogResult.OK) return;
                //grpData.Enabled = false;
                string filename = openDialog.FileName;
                FileStream fs = new FileStream(filename, FileMode.Open, FileAccess.Read, FileShare.Read);
                spreadsheet = new ExcelDataReader.ExcelDataReader(fs);
                fs.Close();

                if (spreadsheet.WorkbookData.Tables.Count > 0)
                {
                    switch (string.IsNullOrEmpty(Base.DBName))
                    {
                        case true:
                            MessageBox.Show("Create or select a database.", "DATABASE NEEDED!", MessageBoxButtons.OK);
                            break;

                        case false:
                            saveTheSpreadSheetToTheDatabase();
                            MessageBox.Show("Successfully Uploaded.", "Information", MessageBoxButtons.OK);
                            break;
                        default:

                            break;
                    }
                }

                //cboSheet.Items.Clear();
                //cboSheet.DisplayMember = "TableName";
                //foreach (DataTable dt in spreadsheet.WorkbookData.Tables)
                //    cboSheet.Items.Add(dt);

                //if (cboSheet.Items.Count == 0) return;

                //grpData.Enabled = true;
                //checker = true;
                //cboSheet.SelectedIndex = 0;
                //btnSave.Visible = true;
                //lblSheet.Visible = true;
                //cboSheet.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to read file: \n" + ex.Message);
            }
        }

        private void TBDeleteTable_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Delete table: " + TB.TBName + " ? ", "Confirm", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
                    extractDBTableNames(listBox1);
                    TB.deleteDataTableFromCollection(TB.DBName);
                    TB.TBName = "";
                    TBFormulas.Tablename = "";
                    loadInfo();
                    break;


                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteCalcColumns_Click(object sender, EventArgs e)
        {
            DialogResult result1 = MessageBox.Show("Confirm DELETE of calculated columns from table: " + TBFormulas.Tablename + "?", "", MessageBoxButtons.YesNo);

            switch (result1)
            {
                case DialogResult.Yes:

                    DataTable tableformulas = Analysis.selectTableFormulasToBeProcessed(TB.DBName, TB.TBName, Base.AnalysisConnectionString);
                    foreach (DataRow row in tableformulas.Rows)
                    {
                        TB.removeColumn(Base.DBConnectionString, TB.TBName, row["CALC_NAME"].ToString());

                    }
                    loadInfo();
                    break;

                case DialogResult.No:
                    break;
            }
        }

        private void TBDeleteAllTables_Click(object sender, EventArgs e)
        {
            foreach (string s in listBox1.Items)
            {
                TB.TBName = s.Trim();
                bool tableCreate = TB.dropDatabaseTable(Base.DBConnectionString);
            }
            extractDBTableNames(listBox1);
            loadInfo();
        }

        private void DBCreate_Click(object sender, EventArgs e)
        {

        }

        private void createNewDatabase(string Databasename)
        {

        }

        private void extractDatabaseFormulas()
        {

        }

        private void DBDeleteList_Click(object sender, EventArgs e)
        {

        }

        private void listDB()
        {

        }

        private void DBList_Click(object sender, EventArgs e)
        {

        }

        private void evaluateStatusButtons()
        {
            btnInsertRow.Enabled = false;
            btnUpdate.Enabled = false;
            btnDeleteRow.Enabled = false;
            btnLoad.Enabled = false;
            btnPrint.Enabled = false;
            btnLock.Enabled = false;

            panelInsert.BackColor = Color.Cornsilk;
            panelUpdate.BackColor = Color.Cornsilk;
            panelDelete.BackColor = Color.Cornsilk;
            panelPreCalcReport.BackColor = Color.Cornsilk;
        }

        private void btnx_Click_1(object sender, EventArgs e)
        {
            btnx.Text = "Running";
            btnx.Enabled = false;
            btnRefresh.Visible = true;

            execute();
            refreshExecution();

        }

        private void refreshExecution()
        {
            lstNames.Clear();

            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcesses(Environment.MachineName);

            foreach (System.Diagnostics.Process process in procs)
            {
                lstNames.Add(process.ProcessName.Trim());

            }

            if (lstNames.Contains("Archive2"))
            {


                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);
                if (Status.Rows.Count > 0)
                {
                    statusChangeButtonColors();
                }
            }
            else
            {

                btnRefresh.Visible = false;
                btnx.Visible = false;

                pictBox.Visible = false;//JVDW
                pictBox2.Visible = false;//JVDW
                calcTime.Enabled = false;//JVDW

                evaluateStatus();
                evaluateStatusButtons();
            }
        }

        private void execute()
        {

            System.Diagnostics.Process P = new System.Diagnostics.Process();
            P.StartInfo.WorkingDirectory = "C:\\OEM\\";
            P.StartInfo.FileName = "Archive2.exe";

            lblRun.Visible = true;
            pictBox.Visible = true;
            pictBox2.Visible = true;
            calcTime.Enabled = true;

            P.Start();
            P.Close();
        }

        private void btnRefresh_Click_1(object sender, EventArgs e)
        {
            lstNames.Clear();

            System.Diagnostics.Process[] procs = System.Diagnostics.Process.GetProcesses(Environment.MachineName);

            foreach (System.Diagnostics.Process process in procs)
            {
                lstNames.Add(process.ProcessName.Trim());

            }

            if (lstNames.Contains("Archive2"))
            {

                Status = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Status", strWhere);
                if (Status.Rows.Count > 0)
                {
                    statusChangeButtonColors();
                }
            }
            else
            {
                btnRefresh.Visible = false;
                btnx.Visible = false;

                pictBox.Visible = false;//JVDW
                pictBox2.Visible = false;//JVDW
                calcTime.Enabled = false;//JVDW

                evaluateStatus();
                evaluateStatusButtons();
            }
        }

        private void printAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.StartReport("STP_AUTHGSUM");
            MessageBox.Show("To be implemented", "Information", MessageBoxButtons.OK);
            this.Cursor = Cursors.Arrow;
        }

        private void btnCostsheetAuth_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            //MetaReportRuntime.App mm = new MetaReportRuntime.App();
            //mm.Init(strMetaReportCode);
            //mm.StartReport("STPTM_CAS");
            MessageBox.Show("To be implemented", "Information", MessageBoxButtons.OK);
            this.Cursor = Cursors.Arrow;
        }

        private void TBExport_Click_1(object sender, EventArgs e)
        {
            saveTheSpreadSheet();
        }

        private void cboNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Search for the coyno in the Labour datatable
            DataTable temp = new DataTable();
            //if (Clocked.Rows.Count > 0)
            //{
            //    IEnumerable<DataRow> query1 = from locks in Clocked.AsEnumerable()
            //                                  where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboNames.Text.Trim()
            //                                  select locks;


            //    temp = query1.CopyToDataTable<DataRow>();
            //}

            if (temp.Rows.Count > 0)
            {
                txtKPFCostLevelKPF.Text = temp.Rows[0]["Employee_Name"].ToString().Trim();
            }
            else
            {
                txtKPFCostLevelKPF.Text = "xxx";
            }

            //if (Labour.Rows.Count > 0)
            //{
            //    IEnumerable<DataRow> query2 = from locks in Labour.AsEnumerable()
            //                                  where locks.Field<string>("EMPLOYEE_NO").TrimEnd() == cboNames.Text.Trim()
            //                                  select locks;


            //    temp = query2.CopyToDataTable<DataRow>();
            //}

            if (temp.Rows.Count > 0)
            {
                txtKPFCostLevelKPFValue.Text = temp.Rows[0]["SHIFTS_WORKED"].ToString().Trim();
                txtKPFCostLevelCapHigh.Text = temp.Rows[0]["AWOP_SHIFTS"].ToString().Trim();
            }
            else
            {
                txtKPFCostLevelCapLow.Text = "0";
                txtKPFCostLevelCapHigh.Text = "0";
            }
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void label15_Click(object sender, EventArgs e)
        {

        }

        private void btnChangePeriod_Click(object sender, EventArgs e)
        {
            //Gets the name of all open forms in application
            foreach (Form form in Application.OpenForms)
            {
                if (form is scrLogon)
                {
                    form.Show(); //Show the form
                    break;
                }
            }
            exitValue = 2;//Change exit value

            this.Close(); //Close the current window

        }

        private void scrEngineering_FormClosing(object sender, FormClosingEventArgs e)//jvdw
        {

            if (exitValue == 0)
            {
                DialogResult result = MessageBox.Show("Have you saved your data? If not sure, please SAVE.", "REMINDER", MessageBoxButtons.YesNo);

                switch (result)
                {
                    case DialogResult.Yes:

                        myConn.Close();
                        AAConn.Close();
                        AConn.Close();
                        exitValue = 1;
                        Application.Exit();
                        break;

                    case DialogResult.No:
                        e.Cancel = true;
                        break;
                }
                if (exitValue == 2)
                {
                    exitValue = 1;
                    this.Close();
                }
            }
        }

        private void btnAttendance_Click(object sender, EventArgs e)
        {

        }

        private void btnAttendance_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            evaluateLabour();
            if (Labour.Rows.Count == 0)
            {
                MessageBox.Show("No Labour records to print for the section: " + txtSelectedSection.Text.Trim(), "Information", MessageBoxButtons.OK);
            }
            else
            {
                DataTable temp = Labour.Copy();
                deleteAllCalcColumnsFromTempTable("BonusShifts", temp);
                TB.createAttendanceTable(Base.DBConnectionString, temp);

                MetaReportRuntime.App mm = new MetaReportRuntime.App();
                mm.Init(strMetaReportCode);
                mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";

                mm.StartReport("ENGSERTA");
            }
            this.Cursor = Cursors.Arrow;
        }

        private void btnSearchEmployNr_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = true;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            txtSearchEmplyNr.Focus();
        }

        private void btnEmployName_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = false;
            txtSearchEmplName.Visible = true;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            txtSearchEmplName.Focus();
        }

        private void btnSearchGang_Click(object sender, EventArgs e)
        {
            txtSearchEmplyNr.Visible = false;
            txtSearchGang.Visible = true;
            txtSearchEmplName.Visible = false;
            txtSearchEmplName.Text = "";
            txtSearchEmplyNr.Text = "";
            txtSearchGang.Text = "";
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            txtSearchGang.Focus();
        }

        private void txtSearchEmplyNr_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NO"], ListSortDirection.Ascending);
            searchEmplNr = txtSearchEmplyNr.Text.ToString();
            searchEmplName = "";
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang); //Calls the metod

        }

        private void txtSearchEmplName_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["EMPLOYEE_NAME"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = txtSearchEmplName.Text.ToString();
            searchEmplGang = "";
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang); //Calls the metod

        }

        private void txtSearchGang_TextChanged(object sender, EventArgs e)
        {
            //Setting the names to be send to the method
            grdLabour.Sort(grdLabour.Columns["GANG"], ListSortDirection.Ascending);
            searchEmplNr = "";
            searchEmplName = "";
            searchEmplGang = txtSearchGang.Text.ToString();
            searchBonus(searchEmplNr, searchEmplName, searchEmplGang); //Calls the metod
        }

        public void searchBonus(string nr, string name, string gang)
        {
            //Sets the details passed to lower case
            nr = nr.ToLower();
            name = name.ToLower();
            gang = gang.ToLower();

            //Gets the length
            int nrLenght = nr.Length;
            int nameLenght = name.Length;
            int gangLenght = gang.Length;

            // Ensuring length are always 1 and not 0 as
            // "" can not be tested.
            if (nrLenght == 0)
            {
                nrLenght = 1;
            }
            if (nameLenght == 0)
            {
                nameLenght = 1;
            }
            if (gangLenght == 0)
            {
                gangLenght = 1;
            }

            //Iterate through all the rows in the grid
            for (int i = 0; i < grdLabour.Rows.Count - 1; i++)
            {
                //Gets the values of the grid in the different columns
                string nrColumn = grdLabour.Rows[i].Cells[0].Value.ToString();  //Cells from grid count from left starting at 0
                string nameColumn = grdLabour.Rows[i].Cells[1].Value.ToString();
                string gangColumn = grdLabour.Rows[i].Cells[2].Value.ToString();

                //Sets the values from grid to lowercase for testing
                nrColumn = nrColumn.ToLower();
                nameColumn = nameColumn.ToLower();
                gangColumn = gangColumn.ToLower();

                //Gets the same amount from the grid string as was entertered bty the user to 
                //ensure the string can be tested
                nrColumn = nrColumn.Substring(1, nrLenght);//Start at 1 to throw away the aphabetic nr
                nameColumn = nameColumn.Substring(0, nameLenght);
                gangColumn = gangColumn.Substring(0, gangLenght);

                //Compares the different strings
                if (nr == nrColumn) //Employee nr
                {
                    //Empty the string not used
                    nameColumn = "";
                    gangColumn = "";
                    grdLabour.ClearSelection(); // Clears all past selection
                    grdLabour.Rows[i].Selected = true; //Selects the current row
                    grdLabour.FirstDisplayedScrollingRowIndex = i; //Jumps automatically to the row
                    break; //breaks the loop
                }
                if (name == nameColumn) //Employee name
                {
                    nrColumn = "";
                    gangColumn = "";
                    grdLabour.ClearSelection();
                    grdLabour.Rows[i].Selected = true;
                    grdLabour.FirstDisplayedScrollingRowIndex = i;
                    break;
                }
                if (gang == gangColumn) //Gang
                {
                    nrColumn = "";
                    nameColumn = "";
                    grdLabour.ClearSelection();
                    grdLabour.Rows[i].Selected = true;
                    grdLabour.FirstDisplayedScrollingRowIndex = i;
                    break;
                }
            }
        }

        private void dataBonusShiftsFromClockedShifts_Click(object sender, EventArgs e)
        {
            InputBoxResult result = InputBox.Show("Import Shifts per Gang.  Gang Number: ", "Employees to import");

            if (result.ReturnCode == DialogResult.OK)
            {

                #region Calculate the shifts per employee en output to bonusshifts

                string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS " +
                                " from Clockedshifts where section = '" +
                                txtSelectedSection.Text.Trim() + "' and Gang = '" + result.Text.Trim() + "'";

                if (BusinessLanguage.MiningType == "STOPE")
                {
                    //strSQL = strSQL.Trim() + " and bonustype = 'Stoping' ";  \\amp
                }
                else
                {
                    if (BusinessLanguage.MiningType == "DEVELOPMENT")
                    {
                        strSQL = strSQL.Trim() + " and bonustype = 'Development' ";
                    }
                }

                BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
                if (BonusShifts.Rows.Count > 0)
                {
                    string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
                    string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

                    DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
                    DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

                    int intStartDay = Base.calcNoOfDays(CalendarFSH, Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()));
                    int intEndDay = Base.calcNoOfDays(CalendarLSH, Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()));
                    int intStopDay = 0;

                    //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
                    int intNoOfDays = Base.calcNoOfDays(Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()), Convert.ToDateTime(BonusShifts.Rows[0]["FSH"].ToString()));

                    if (intStartDay < 0)
                    {
                        //The calendarFSH falls outside the startdate of the sheet.
                        intStartDay = 0;
                    }
                    else
                    {

                    }

                    if (intEndDay < 0 && intEndDay < -40)
                    {
                        intStopDay = 0;
                    }
                    else
                    {
                        if (intEndDay < 0)
                        {
                            //the LSH of the measuring period falls within the spreadsheet
                            intStopDay = intNoOfDays + intEndDay;

                        }
                        else
                        {
                            //The LSH of the measuring period falls outside the spreadsheet
                            intStopDay = 40;
                        }


                        //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                        //were not imported.

                        #region count the shifts
                        //Count the the shifts

                        DialogResult result2 = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for gang " + result.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                        switch (result2)
                        {
                            case DialogResult.OK:
                                
                                strWhere = strWhere + " and gang = '" + result.Text.Trim() + "'";
                                extractAndCalcShifts(intStartDay, intStopDay);
                                
                                break;

                            case DialogResult.Cancel:
                                break;

                        }

                        #endregion

                #endregion

                    }
                }
            }
        }

        private void dataPrintFormulas_Click(object sender, EventArgs e)
        {
            DataTable dt = Base.dataPrintFormulasBonusShifts(Base.AnalysisConnectionString, Base.DBName, "HOD");
            if (dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row["FORMULA_NAME"].ToString().Trim().Length > 3 && row["FORMULA_NAME"].ToString().Trim().Substring(0, 3) == "SQL")
                    {
                    }
                    else
                    {
                        switch (row["INPUTORDER"].ToString().Trim())
                        {
                            case "0":
                                row["INPUTORDER"] = "A = ";
                                break;
                            case "1":
                                row["INPUTORDER"] = "B = ";
                                break;
                            case "2":
                                row["INPUTORDER"] = "C = ";
                                break;
                            case "3":
                                row["INPUTORDER"] = "D = ";
                                break;
                            case "4":
                                row["INPUTORDER"] = "E = ";
                                break;
                            case "5":
                                row["INPUTORDER"] = "F = ";
                                break;
                            case "6":
                                row["INPUTORDER"] = "G = ";
                                break;
                            case "7":
                                row["INPUTORDER"] = "H = ";
                                break;
                            case "8":
                                row["INPUTORDER"] = "I = ";
                                break;
                            case "9":
                                row["INPUTORDER"] = "J = ";
                                break;
                        }
                    }
                }
                printHTML(dt, "Head of Departments");
            }
            else
            {
                MessageBox.Show("No records available to print", "", MessageBoxButtons.OK);
            }
        }

        private void auditByTable_Click(object sender, EventArgs e)
        {
            DataTable audit = TB.createDataTableWithAdapterSelectAll(Base.DBConnectionString, "Audit", " where tablename = 'Ganglink'");
            string[] auditcolumns = new string[10];

            string test = audit.Rows[0]["PK"].ToString().Trim();
            int testlength = test.Length;

            for (int i = 0; i <= 9; i++)
            {
                int tstLength = test.IndexOf(">");
                if (tstLength != -1)
                {
                    auditcolumns[i] = test.Substring(0, tstLength).Replace("<", "").Trim();
                    test = test.Substring(test.IndexOf(">") + 1);
                }

            }

        }

        private void btnEmplyeAudit_Click(object sender, EventArgs e)
        {


            #region extract the sheet name and FSH and LSH of the extract
            string FilePath = "C:\\iCalc\\Harmony\\Phakisa\\Development\\Data\\ADTeam_201005.xls";
            string[] sheetNames = GetExcelSheetNames(FilePath);
            string sheetName = sheetNames[0];
            #endregion

            #region import Clockshifts
            this.Cursor = Cursors.WaitCursor;
            DataTable dt = new DataTable();

            OleDbConnection con = new OleDbConnection();
            OleDbDataAdapter da;
            con.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="
                    + FilePath + ";Extended Properties='Excel 8.0;'";

            /*"HDR=Yes;" indicates that the first row contains columnnames, not data.
            * "HDR=No;" indicates the opposite.
            * "IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. 
            * Note that this option might affect excel sheet write access negative.
            */

            da = new OleDbDataAdapter("select * from [" + sheetName + "]", con); //read first sheet named Sheet1
            da.Fill(dt);

            #region remove invalid records
            // Delete records that does not conform to configurations
            //foreach (DataRow row in dt.Rows)
            //{
            //    if ((row["GANG NAME"].ToString().Substring(5, 1) == "A" || row["GANG NAME"].ToString().Substring(5, 1) == "B" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "C" || row["GANG NAME"].ToString().Substring(5, 1) == "D" ||
            //        row["GANG NAME"].ToString().Substring(5, 1) == "E" || row["WAGE CODE"].ToString() == "245M003" ||
            //        row["WAGE CODE"].ToString() == "400M009" || row["WAGE CODE"].ToString() == "245M001" ||
            //        row["WAGE CODE"].ToString() == "246M004" || row["WAGE CODE"].ToString() == "400M009")
            //        && (row["GANG NAME"].ToString().Substring(0, 5) == txtSelectedSection.Text.Trim()))
            //    {
            //    }
            //    else
            //    {
            //        //row.Delete();
            //    }

            //}

            //dt.AcceptChanges();

            //extract the column names with length less than 3.  These columns must be deleted.
            string[] columnNames = new String[dt.Columns.Count];

            for (int i = 0; i <= dt.Columns.Count - 1; i++)
            {
                if (dt.Columns[i].ColumnName.Length <= 2)
                {
                    columnNames[i] = dt.Columns[i].ColumnName;
                }
            }

            for (Int16 i = 0; i <= columnNames.GetLength(0) - 1; i++)
            {
                if (string.IsNullOrEmpty(columnNames[i]))
                {

                }
                else
                {
                    dt.Columns.Remove(columnNames[i].ToString().Trim());
                    dt.AcceptChanges();
                }
            }

            dt.Columns.Remove("INDUSTRY NUMBER");
            dt.AcceptChanges();
            #endregion

            string strSheetFSH = string.Empty;
            string strSheetLSH = string.Empty;

            //Extract the dates from the spreadsheet - the name of the spreadsheet contains the the start and enddate of the extract
            string strSheetFSHx = sheetName.Substring(0, sheetName.IndexOf("_TO")).Replace("_", "-").Replace("'", "").Trim(); ;
            string strSheetLSHx = sheetName.Substring(sheetName.IndexOf("_TO") + 4).Replace("$", "").Replace("_", "-").Replace("'", "").Trim(); ;

            //Correct the dates and calculate the number of days extracted.
            if (strSheetFSHx.Substring(6, 1) == "-")
            {
                strSheetFSH = strSheetFSHx.Substring(0, 5) + "0" + strSheetFSHx.Substring(5);
            }

            if (strSheetLSHx.Substring(6, 1) == "-")
            {
                strSheetLSH = strSheetLSHx.Substring(0, 5) + "0" + strSheetLSHx.Substring(5);
            }

            DateTime SheetFSH = Convert.ToDateTime(strSheetFSH.ToString());
            DateTime SheetLSH = Convert.ToDateTime(strSheetLSH.ToString());

            //If the intNoOfDays < 40 then the days up to 40 must be filled with '-'
            intNoOfDays = Base.calcNoOfDays(SheetLSH, SheetFSH);
            noOFDay = intNoOfDays;

            if (intNoOfDays <= 40)
            {
                for (int j = intNoOfDays + 1; j <= 40; j++)
                {
                    dt.Columns.Add("DAY" + j);
                }
            }
            else
            {

            }

            #region Change the column names
            //Change the column names to the correct column names.
            Dictionary<string, string> dictNames = new Dictionary<string, string>();
            DataTable varNames = TB.createDataTableWithAdapter(Base.AnalysisConnectionString,
                                 "Select * from varnames");
            dictNames.Clear();

            dictNames = TB.loadDict(varNames, dictNames);
            int counter = 0;


            //If it is a column with a date as a name.
            foreach (DataColumn column in dt.Columns)
            {
                if (column.ColumnName.Substring(0, 1) == "2")
                {
                    if (counter == 0)
                    {
                        strSheetFSH = column.ColumnName.ToString().Replace("/", "-");
                        column.ColumnName = "DAY" + counter;
                        counter = counter + 1;

                    }
                    else
                    {
                        if (column.Ordinal == dt.Columns.Count - 1)
                        {

                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;

                        }
                        else
                        {
                            column.ColumnName = "DAY" + counter;
                            counter = counter + 1;
                        }
                    }


                }
                else
                {
                    if (dictNames.Keys.Contains<string>(column.ColumnName.Trim().ToUpper()))
                    {
                        column.ColumnName = dictNames[column.ColumnName.Trim().ToUpper()];
                    }

                }
            }

            //Add the extra columns
            dt.Columns.Add("FSH");
            dt.Columns.Add("LSH");
            dt.Columns.Add("SECTION");
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                row["FSH"] = strSheetFSH;
                row["LSH"] = strSheetLSH;
                row["MININGTYPE"] = "STOPE";
                row["SECTION"] = row["GANG"].ToString().Substring(0, 5);

                for (int i = 0; i <= dt.Columns.Count - 1; i++)
                {
                    if (string.IsNullOrEmpty(row[i].ToString()) || row[i].ToString() == "")
                    {
                        row[i] = "-";
                    }
                }
            }
            #endregion


            #endregion

            #region Calculate the shifts per employee en output to bonusshifts

            //string strSQL = "Select *,'0' as SHIFTS_WORKED,'0' as AWOP_SHIFTS, '0' as STRIKE_SHIFTS," +
            //                "'0' as DRILLERIND,'0' AS DRILLERSHIFTS from Clockedshifts where (section = '"
            //                + txtSelectedSection.Text.Trim() + "' or WAGE_DESCRIPTION = 'STOPER')";

            string strSQLFix = "Select *,'0' as SHIFTS_WORKED from Clockedshifts";//jvdw

            // BonusShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);
            fixShifts = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQLFix);//jvdw laai die hele clockedshift table

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            sheetfhs = SheetFSH;//jvdw
            sheetlhs = SheetLSH;//jvdw
            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -40)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 40;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                // DialogResult result = MessageBox.Show("Do you want to REPLACE the current BONUSSHIFTS for section " + txtSelectedSection.Text.Trim() + " ?", "QUESTION", MessageBoxButtons.OKCancel);

                //switch (result)
                //{
                //    case DialogResult.OK:
                //        extractAndCalcShifts(intStartDay, intStopDay);
                //        break;

                //    case DialogResult.Cancel:
                //        break;

                //}

                #endregion

            #endregion

                #region Extract the ganglinking of the current section


                #endregion

                fillFixTable(fixShifts, sheetfhs, sheetlhs, intNoOfDays, intStartDay, intStopDay);
                this.Cursor = Cursors.Arrow;
                //}
            }

        }

        public void fillFixTable(DataTable clockedTable, DateTime SheetFSH, DateTime SheetLSH, int intNoOfDays, int DayStart, int DayEnd)//jvdw
        {
            //Calculate the shifts in the clockedshifts table and insert all in a fixed
            //table that cannot be changed by the user!

            string SQLTable = "IF OBJECT_ID(N'emplshiftfix', N'U')IS NOT NULL DROP TABLE EMPLSHIFTFIX create table EMPLSHIFTFIX (employeeno char(20),shiftsfix char(20)) truncate table EMPLSHIFTFIX";
            Base.VoidQuery(Base.DBConnectionString, SQLTable);

            #region Calculate the shifts per employee en output to bonusshifts

            string strCalendarFSH = dateTimePicker1.Value.ToString("yyyy-MM-dd");
            string strCalendarLSH = dateTimePicker2.Value.ToString("yyyy-MM-dd");

            DateTime CalendarFSH = Convert.ToDateTime(strCalendarFSH.ToString());
            DateTime CalendarLSH = Convert.ToDateTime(strCalendarLSH.ToString());

            intStartDay = Base.calcNoOfDays(CalendarFSH, SheetFSH);
            intEndDay = Base.calcNoOfDays(CalendarLSH, SheetLSH);
            intStopDay = 0;

            if (intStartDay < 0)
            {
                //The calendarFSH falls outside the startdate of the sheet.
                intStartDay = 0;
            }
            else
            {

            }

            if (intEndDay < 0 && intEndDay < -40)
            {
                intStopDay = 0;
            }
            else
            {
                if (intEndDay < 0)
                {
                    //the LSH of the measuring period falls within the spreadsheet
                    intStopDay = intNoOfDays + intEndDay;

                }
                else
                {
                    //The LSH of the measuring period falls outside the spreadsheet
                    intStopDay = 40;
                }


                //If intStartDay < 0 then the SheetFSH is bigger than the calendarFSH.  Therefore some of the Calendar's shifts 
                //were not imported.

                #region count the shifts
                //Count the the shifts

                int intSubstringLength = 0;
                int intShiftsWorked = 0;
                int intAwopShifts = 0;
                int shiftsCheck = 0;
                StringBuilder sqlInsertFixShifts = new StringBuilder("BEGIN TRANSACTION; ");

                foreach (DataRow row in clockedTable.Rows)
                {
                    foreach (DataColumn column in clockedTable.Columns)
                    {
                        if ((column.ColumnName.Substring(0, 3) == "DAY"))
                        {

                            if (column.ColumnName.ToString().Length == 4)
                            {
                                intSubstringLength = 1;
                            }
                            else
                            {
                                intSubstringLength = 2;
                            }

                            if ((Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) >= DayStart &&
                               Convert.ToInt16(column.ColumnName.Substring(3, intSubstringLength)) <= (DayEnd)))
                            {

                                if (row[column].ToString().Trim() == "U" || row[column].ToString().Trim() == "u")
                                {
                                    intShiftsWorked = intShiftsWorked + 1;
                                    shiftsCheck = 1;
                                }
                                else
                                {
                                    if (row[column].ToString().Trim() == "A")
                                    {
                                        intAwopShifts = intAwopShifts + 1;
                                    }
                                    else { }

                                }
                            }
                            else
                            {
                                row[column] = "*";
                            }
                        }
                        else
                        {
                            if (column.ColumnName == "BONUSTYPE")
                            {
                                row["BONUSTYPE"] = "SERVICES";
                            }
                        }
                    }//foreach datacolumn

                    row["SHIFTS_WORKED"] = intShiftsWorked;

                    string emplNr = row["employee_no"].ToString();
                    workedShiftsFixedClockedShift = intShiftsWorked;
                    intShiftsWorked = 0;
                    intAwopShifts = 0;
                    if (shiftsCheck == 1)
                    {
                        sqlInsertFixShifts.Append("INSERT INTO EMPLSHIFTFIX VALUES ('" + emplNr.Trim() + "','" + workedShiftsFixedClockedShift.ToString().Trim() + "');");
                    }
                }

                sqlInsertFixShifts.Append(" COMMIT TRANSACTION");


                Base.VoidQuery(Base.DBConnectionString, sqlInsertFixShifts.ToString());



                #endregion

            #endregion

            }
        }

        private void lstBErrorLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedNr = lstBErrorLog.SelectedItem.ToString();
            if (selectedNr != "Employee Nr")
            {

                selectedNr = selectedNr.Remove(0, 1);
                int last = selectedNr.LastIndexOf("-");
                selectedNr = selectedNr.Remove(last - 1).Trim();
                txtSearchEmplyNr.Visible = true;
                txtSearchEmplyNr.Text = selectedNr;
            }
        }

        private void hideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            grdActiveSheet.Columns[columnnr].Visible = false;


        }

        private void calcTime_Tick(object sender, EventArgs e)
        {
            btnRefresh_Click_1("Method", null);
        }

        private void grdSubsectionDept_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboSubsection.Text = grdSubsectionDept["SUBSECTION", e.RowIndex].Value.ToString().Trim();
                cboDepartment.Text = grdSubsectionDept["DEPARTMENT", e.RowIndex].Value.ToString().Trim();
                cboHODModel.Text = grdSubsectionDept["HODMODEL", e.RowIndex].Value.ToString().Trim();
                cboCostLevel.Text = grdSubsectionDept["COSTLEVEL", e.RowIndex].Value.ToString().Trim();

                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = true;
                btnInsertRow.Enabled = true;

            }
            Cursor.Current = Cursors.Arrow;
        }

        private void grdDeptParameters_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                cboDeptParametersSubsection.Text = grdDeptParameters["Subsection", e.RowIndex].Value.ToString().Trim();
                cboDeptParametersDepartment.Text = grdDeptParameters["Department", e.RowIndex].Value.ToString().Trim();
                cboDeptParametersKPF.Text = grdDeptParameters["KPF", e.RowIndex].Value.ToString().Trim();
                cboDeptParametersKPFParameters.Text = grdDeptParameters["KPFParameter", e.RowIndex].Value.ToString().Trim();
                txtDeptParametersDesc.Text = grdDeptParameters["KPFParameterDesc", e.RowIndex].Value.ToString().Trim();

                btnUpdate.Enabled = true;
                btnDeleteRow.Enabled = true;
                btnInsertRow.Enabled = true;

            }
            Cursor.Current = Cursors.Arrow;
        }

        private void btnLockParticipation_Click(object sender, EventArgs e)
        {
            openTab(tabParticipation);
        }

        private void btnLockDeptParameters_Click(object sender, EventArgs e)
        {
            //Open the DeptParameters and allow the user to change the descriptions.

            openTab(tabDeptParameters);
        }

        private void btnParticipation_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            DataTable temp = TB.extractNewParticipation(Base.DBConnectionString, txtSelectedSection.Text.Trim());
            DataTable nodupsParticipation = new DataTable();

            //If participation is brandnew
            #region New Participation table
            if (Participation.Rows.Count == 0)
            {
                extractDictionaries();

                //Update the participation table with the correct employee and gangtype values

                foreach (DataRow dr in temp.Rows)
                {
                    if (string.IsNullOrEmpty((string)(dr["WAGE_DESCRIPTION"])))
                    {
                        dr["WAGE_DESCRIPTION"] = "UNKNOWN";
                    }
                    try
                    {
                        dr["GANGPARTICIPATION"] = GangTypes[dr["GANGTYPE"].ToString().Trim()].ToString().Trim();
                    }
                    catch
                    {
                        dr["GANGPARTICIPATION"] = "0";
                    }
                    try
                    {
                        dr["EMPLOYEETYPEPARTICIPATION"] = Employeetypes[dr["EMPLOYEETYPE"].ToString().Trim()].ToString().Trim();
                    }
                    catch
                    {
                        dr["EMPLOYEETYPEPARTICIPATION"] = "0";
                    }

                    try
                    {
                        dr["WAGECODEPARTICIPATION"] = Wagecodetypes[dr["WAGECODE"].ToString().Trim()].ToString().Trim();
                    }
                    catch
                    {
                        dr["WAGECODEPARTICIPATION"] = "0";
                    }
                }

                temp.AcceptChanges();

                //remove the duplicate records
                string duptest = string.Empty;

                for (int i = 0; i <= temp.Rows.Count - 1; i++)
                {

                    if (temp.Rows[i]["SUBSECTION"].ToString().Trim() +
                        temp.Rows[i]["DEPARTMENT"].ToString().Trim() +
                        temp.Rows[i]["EMPLOYEETYPE"].ToString().Trim() +
                        temp.Rows[i]["GANGTYPE"].ToString().Trim() +
                        temp.Rows[i]["WAGECODE"].ToString().Trim() == duptest)
                    {
                        temp.Rows[i]["WAGECODE"] = "XXXX";
                    }
                    else
                    {
                        duptest = temp.Rows[i]["SUBSECTION"].ToString().Trim() +
                        temp.Rows[i]["DEPARTMENT"].ToString().Trim() +
                        temp.Rows[i]["EMPLOYEETYPE"].ToString().Trim() +
                        temp.Rows[i]["GANGTYPE"].ToString().Trim() +
                        temp.Rows[i]["WAGECODE"].ToString().Trim();
                    }
                }

                IEnumerable<DataRow> query1 = from locks in temp.AsEnumerable()
                                              where locks.Field<string>("WAGECODE").TrimEnd() != "XXXX"
                                              where locks.Field<string>("SUBSECTION").TrimEnd() != "UNKNOWN"
                                              select locks;
                ;


                try
                {
                    nodupsParticipation = query1.CopyToDataTable<DataRow>();
                }
                catch
                {
                }

                TB.saveCalculations2(nodupsParticipation, Base.DBConnectionString, "", "PARTICIPATION");
                MessageBox.Show("Done", "Information", MessageBoxButtons.OK);


            #endregion

            }

            else
            {
                extractDictionaries();

                DataTable table = TB.ExtractParticipationRecords(Base.DBConnectionString);
                //Update the extra participation records table with the correct values

                foreach (DataRow dr in table.Rows)
                {
                    if (string.IsNullOrEmpty((string)(dr["WAGE_DESCRIPTION"])))
                    {
                        dr["WAGE_DESCRIPTION"] = "UNKNOWN";
                    }
                    try
                    {
                        dr["GANGPARTICIPATION"] = GangTypes[dr["GANGTYPE"].ToString().Trim()].ToString().Trim();
                    }
                    catch
                    {
                        dr["GANGPARTICIPATION"] = "0";
                    }
                    try
                    {
                        dr["EMPLOYEETYPEPARTICIPATION"] = Employeetypes[dr["EMPLOYEETYPE"].ToString().Trim()].ToString().Trim();
                    }
                    catch
                    {
                        dr["EMPLOYEETYPEPARTICIPATION"] = "0";
                    }

                    try
                    {
                        dr["WAGECODEPARTICIPATION"] = Wagecodetypes[dr["WAGECODE"].ToString().Trim()].ToString().Trim();
                    }
                    catch
                    {
                        dr["WAGECODEPARTICIPATION"] = "0";
                    }
                }

                table.AcceptChanges();

                TB.insertExtraParticipationRecords(Base.DBConnectionString, table);
                MessageBox.Show("Done", "Information", MessageBoxButtons.OK);

            }


            evaluateParticipation();
            this.Cursor = Cursors.Arrow;

        }

        private void extractDictionaries()
        {
            #region extract gangtypes percentage
            //Extract the rates for Surface Gangs, Dusty Gangs and Underground gangs
            IEnumerable<DataRow> query1 = from locks in Rates.AsEnumerable()
                                          where locks.Field<string>("RATE_TYPE").TrimEnd() == "ENGINEERINGGANGTYPES"
                                          select locks;
            ;


            try
            {
                DataTable gangtypes = query1.CopyToDataTable<DataRow>();
                foreach (DataRow dr in gangtypes.Rows)
                {

                    GangTypes.Add(dr["LOW_VALUE"].ToString().Trim(), dr["RATE"].ToString().Trim());

                }

            }
            catch
            {

            }
            #endregion

            #region extract employeetype percentage
            //Extract the rates for Artisans,Officials and Grp3-8 employees
            query1 = from locks in Rates.AsEnumerable()
                     where locks.Field<string>("RATE_TYPE").TrimEnd() == "ENGINEERINGEMPLOYEETYPE"
                     select locks;
            ;


            try
            {
                DataTable emptypes = query1.CopyToDataTable<DataRow>();
                foreach (DataRow dr in emptypes.Rows)
                {

                    Employeetypes.Add(dr["LOW_VALUE"].ToString().Trim(), dr["RATE"].ToString().Trim());

                }

            }
            catch
            {

            }
            #endregion

            #region extract min wagecode percentage
            //Extract the distint wagecodes with the minimum wagecode percentage from participation

            string strSQL = "select distinct wagecode, min(wagecodeparticipation) as wagecodeparticipation from participation group by WAGECODE";

            DataTable temp = TB.createDataTableWithAdapter(Base.DBConnectionString, strSQL);

            foreach (DataRow dr in temp.Rows)
            {
                Wagecodetypes.Add(dr["WAGECODE"].ToString().Trim(), dr["WAGECODEPARTICIPATION"].ToString().Trim());
            }

            #endregion
        }

        private void grdMineParameters_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {

            }
            else
            {
                txtCost_Actual.Text = grdMineParameters["Cost_Actual", e.RowIndex].Value.ToString().Trim();
                txtCost_Planned.Text = grdMineParameters["Cost_Planned", e.RowIndex].Value.ToString().Trim();
                txtTonsPerTEC_Actual.Text = grdMineParameters["TONSPERTEC_Actual", e.RowIndex].Value.ToString().Trim();
                txtTonsPerTEC_Planned.Text = grdMineParameters["TONSPERTEC_Planned", e.RowIndex].Value.ToString().Trim();
                txtMineSafety_Actual.Text = grdMineParameters["SAFETY_Actual", e.RowIndex].Value.ToString().Trim();
                //txtGoldActual.Text = grdMineParameters["GOLD_Actual", e.RowIndex].Value.ToString().Trim();
                //txtGoldActual.Text = grdMineParameters["GOLD_Planned", e.RowIndex].Value.ToString().Trim();
                //txtProdActual.Text = grdMineParameters["PRODUCTION_Actual", e.RowIndex].Value.ToString().Trim();
                //txtProdPlanned.Text = grdMineParameters["PRODUCTION_Planned", e.RowIndex].Value.ToString().Trim();


                btnUpdate.Enabled = true;
                btnInsertRow.Enabled = false;
                btnDeleteRow.Enabled = false;

            }
            Cursor.Current = Cursors.Arrow;

        }

        private void grdHOD_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex < 0)
            {

            }
            else
            {
                txtSafetyA.Enabled = true;
                txtTonsPerTECA.Enabled = true;
                txtTonsPerTECP.Enabled = true;
                txtCostA.Enabled = true;
                txtCostP.Enabled = true;

                txtSafetyA.BackColor = Color.Cornsilk;
                txtTonsPerTECA.BackColor = Color.Cornsilk;
                txtTonsPerTECP.BackColor = Color.Cornsilk;
                txtCostA.BackColor = Color.Cornsilk;
                txtCostP.BackColor = Color.Cornsilk;
                txtHODHODModel.BackColor = Color.Cornsilk;

                ParameterNames.Clear();
                IEnumerable<DataRow> query1 = from locks in DeptParameters.AsEnumerable()
                                              where locks.Field<string>("SUBSECTION").TrimEnd() == grdHOD["Subsection", e.RowIndex].Value.ToString().Trim()
                                              where locks.Field<string>("DEPARTMENT").TrimEnd() == grdHOD["Department", e.RowIndex].Value.ToString().Trim()
                                              select locks;

                try
                {
                    //Load the Personal PErformance descriptions and replace the labels with the description.
                    DataTable temp = query1.CopyToDataTable<DataRow>();
                    ParameterNames.Clear();
                    foreach (DataRow r in temp.Rows)
                    {

                        ParameterNames.Add(r["KPFPARAMETER"].ToString().Trim(), r["KPFPARAMETERDESC"].ToString().Trim());
                    }

                    label999.Text = ParameterNames["ITEM1_ACTUAL"];
                    label1000.Text = ParameterNames["ITEM1_PLANNED"];
                    label1001.Text = ParameterNames["ITEM2_ACTUAL"];
                    label1002.Text = ParameterNames["ITEM2_PLANNED"];
                    label1003.Text = ParameterNames["ITEM3_ACTUAL"];
                    label1004.Text = ParameterNames["ITEM3_PLANNED"];
                    label1005.Text = ParameterNames["ITEM4_ACTUAL"];
                    label1006.Text = ParameterNames["ITEM4_PLANNED"];
                    label1007.Text = ParameterNames["ITEM5_ACTUAL"];
                    label1008.Text = ParameterNames["ITEM5_PLANNED"];

                }
                catch
                {
                    label999.Text = "ITEM1_ACTUAL";
                    label1000.Text = "ITEM1_PLANNED";
                    label1001.Text = "ITEM2_ACTUAL";
                    label1002.Text = "ITEM2_PLANNED";
                    label1003.Text = "ITEM3_ACTUAL";
                    label1004.Text = "ITEM3_PLANNED";
                    label1005.Text = "ITEM4_ACTUAL";
                    label1006.Text = "ITEM4_PLANNED";
                    label1007.Text = "ITEM5_ACTUAL";
                    label1008.Text = "ITEM5_PLANNED";
                    Application.DoEvents();


                }

                ////Extract the mine parameters  into a dictionary
                //Dictionary<string, string> MineValues = new Dictionary<string, string>();
                //IEnumerable<DataRow> query2 = from locks in MineParameters.AsEnumerable()
                //                              select locks;

                //try
                //{
                //    //Load the dictionary
                //    DataTable temp = query2.CopyToDataTable<DataRow>();
                //    foreach (DataColumn dc in temp.Columns)
                //    {
                //        if (dc.ColumnName.Contains("ACTUAL") || dc.ColumnName.Contains("PLANNED"))
                //        {
                //            MineValues.Add(dc.ColumnName.ToString().Trim(), temp.Rows[0][dc.ColumnName.ToString().Trim()].ToString().Trim());
                //        }
                //    }
                //}
                //catch
                //{
                //}

                //Block the input text boxes if the mine parameters is a parameter in the parameternames dictionary.
                //If the parameter name does not begin with HOD, it is a mine parameter and should be blocked.

                foreach (KeyValuePair<string, String> kvp in ParameterNames)
                {
                    string parametername = kvp.Key;
                    switch (parametername)
                    {
                        case "SAFETY_ACTUAL":
                            txtSafetyA.BackColor = Color.AliceBlue;
                            txtSafetyA.Enabled = false;
                            //txtSafetyA.Text = MineValues["SAFETY_ACTUAL"];
                            //txtMineSafetyA.Text = MineValues["SAFETY_ACTUAL"];

                            break;

                        case "TONS_ACTUAL":

                            txtTonsPerTECA.BackColor = Color.AliceBlue;
                            txtTonsPerTECA.Enabled = false;
                            try
                            {
                                //txtTonsPerTECA.Text = MineValues["TEC_ACTUAL"];
                                //txtMineTECA.Text = MineValues["TEC_ACTUAL"];
                            }
                            catch
                            {
                                txtTonsPerTECA.Text = "ToBeCalculated";
                                //txtMineTECA.Text = "ToBeCalculated";
                            }
                            break;

                        case "TONS_PLANNED":
                            txtTonsPerTECP.BackColor = Color.AliceBlue;
                            txtTonsPerTECP.Enabled = false;
                            //txtTonsPerTECP.Text = MineValues["TONSPERTEC_PLANNED"];
                            //txtMineTECP.Text = MineValues["TONSPERTEC_PLANNED"];
                            break;

                        case "COST_ACTUAL":
                            txtCostA.BackColor = Color.AliceBlue;
                            txtCostA.Enabled = false;
                            //txtCostA.Text = MineValues["COST_ACTUAL"];
                            //txtMineCostA.Text = MineValues["COST_ACTUAL"];
                            break;

                        case "COST_PLANNED":
                            txtCostP.BackColor = Color.AliceBlue;
                            txtCostP.Enabled = false;
                            //txtCostP.Text = MineValues["COST_PLANNED"];
                            //txtMineCostP.Text = MineValues["COST_PLANNED"];
                            break;
                    }
                }

                txtHODEmployeeNo.Text = grdHOD["Employee_no", e.RowIndex].Value.ToString().Trim();
                txtHODEmployeeName.Text = grdHOD["Employee_name", e.RowIndex].Value.ToString().Trim();
                cboHODDesignation.Text = grdHOD["Designation", e.RowIndex].Value.ToString().Trim();
                txtHODDesignationDesc.Text = grdHOD["Designation_Desc", e.RowIndex].Value.ToString().Trim();
                cboHODSubsection.Text = grdHOD["Subsection", e.RowIndex].Value.ToString().Trim();
                cboHODDepartment.Text = grdHOD["Department", e.RowIndex].Value.ToString().Trim();
                txtHODShiftsWorked.Text = grdHOD["Shifts_Worked", e.RowIndex].Value.ToString().Trim();
                txtHODAwops.Text = grdHOD["Awop_shifts", e.RowIndex].Value.ToString().Trim();
                txtHODHODModel.Text = grdHOD["Hodmodel", e.RowIndex].Value.ToString().Trim();
                txtSafetyA.Text = grdHOD["Safety_Actual", e.RowIndex].Value.ToString().Trim();
                txtCostA.Text = grdHOD["Cost_Actual", e.RowIndex].Value.ToString().Trim();
                txtCostP.Text = grdHOD["Cost_Planned", e.RowIndex].Value.ToString().Trim();
                txtTonsPerTECA.Text = grdHOD["Tons_Actual", e.RowIndex].Value.ToString().Trim();
                txtTonsPerTECP.Text = grdHOD["Tons_Planned", e.RowIndex].Value.ToString().Trim();
                txtItem1Actual.Text = grdHOD["Item1_Actual", e.RowIndex].Value.ToString().Trim();
                txtItem2Actual.Text = grdHOD["Item2_Actual", e.RowIndex].Value.ToString().Trim();
                txtItem3Actual.Text = grdHOD["Item3_Actual", e.RowIndex].Value.ToString().Trim();
                txtItem4Actual.Text = grdHOD["Item4_Actual", e.RowIndex].Value.ToString().Trim();
                txtItem5Actual.Text = grdHOD["Item5_Actual", e.RowIndex].Value.ToString().Trim();
                txtItem1Planned.Text = grdHOD["Item1_Planned", e.RowIndex].Value.ToString().Trim();
                txtItem2Planned.Text = grdHOD["Item2_Planned", e.RowIndex].Value.ToString().Trim();
                txtItem3Planned.Text = grdHOD["Item3_Planned", e.RowIndex].Value.ToString().Trim();
                txtItem4Planned.Text = grdHOD["Item4_Planned", e.RowIndex].Value.ToString().Trim();
                txtItem5Planned.Text = grdHOD["Item5_Planned", e.RowIndex].Value.ToString().Trim();

            }

            btnUpdate.Enabled = true;
            btnInsertRow.Enabled = false;
            btnDeleteRow.Enabled = true;
        }

        private void updateTextBoxes()
        {
            //Convert all the actual and planned text boxes na zeroes.
            txtTotalPlanned.BackColor = Color.PowderBlue;
            if (string.IsNullOrEmpty(txtItem1Planned.Text))
            {
                txtItem1Planned.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem2Planned.Text))
            {
                txtItem2Planned.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem3Planned.Text))
            {
                txtItem3Planned.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem4Planned.Text))
            {
                txtItem4Planned.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem5Planned.Text))
            {
                txtItem5Planned.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem1Actual.Text))
            {
                txtItem1Actual.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem1Actual.Text))
            {
                txtItem1Actual.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem1Actual.Text))
            {
                txtItem1Actual.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem1Actual.Text))
            {
                txtItem1Actual.Text = "0";
            }
            if (string.IsNullOrEmpty(txtItem1Actual.Text))
            {
                txtItem1Actual.Text = "0";
            }

            txtTotalActuals.Text = Convert.ToString(Convert.ToDecimal(txtItem1Actual.Text) +
                                                   Convert.ToDecimal(txtItem2Actual.Text) +
                                                   Convert.ToDecimal(txtItem3Actual.Text) +
                                                   Convert.ToDecimal(txtItem4Actual.Text) +
                                                   Convert.ToDecimal(txtItem5Actual.Text));

            txtTotalPlanned.Text = Convert.ToString(Convert.ToDecimal(txtItem1Planned.Text) +
                                                 Convert.ToDecimal(txtItem2Planned.Text) +
                                                 Convert.ToDecimal(txtItem3Planned.Text) +
                                                 Convert.ToDecimal(txtItem4Planned.Text) +
                                                 Convert.ToDecimal(txtItem5Planned.Text));
        }

        private void btnUpdateMineParameters_Click(object sender, EventArgs e)
        {
            btnLockMineParameters_Click("Method", null);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            btnLockDeptParameters_Click("Method", null);
        }

        

        private void btnHODPrint_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("ENGSERFM2");
            this.Cursor = Cursors.Arrow;
        }

        private void printTeam_Click_1(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.StartReport("ENGSERTEAM");
            this.Cursor = Cursors.Arrow;
        }

        private void cboHODSubsection_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in SubsectionDept.AsEnumerable()
                                          where locks.Field<string>("Subsection").TrimEnd() == cboHODSubsection.Text.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            if (temp.Rows.Count > 0)
            {
                cboHODDepartment.Items.Clear();
                lstNames = TB.loadDistinctValuesFromColumn(temp, "Department");

                foreach (string s in lstNames)
                {
                    cboHODDepartment.Items.Add(s.Trim());

                }

                cboHODDepartment.Text = cboHODDepartment.Items[0].ToString();
            }
        }

        private void cboDeptParametersubsection_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in SubsectionDept.AsEnumerable()
                                          where locks.Field<string>("Subsection").TrimEnd() == cboDeptParametersSubsection.Text.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            if (temp.Rows.Count > 0)
            {
                cboDeptParametersDepartment.Items.Clear();
                lstNames = TB.loadDistinctValuesFromColumn(temp, "Department");

                foreach (string s in lstNames)
                {
                    cboDeptParametersDepartment.Items.Add(s.Trim());

                }

                cboDeptParametersDepartment.Text = cboDeptParametersDepartment.Items[0].ToString();
            }

            IEnumerable<DataRow> query2 = from locks in DeptParameters.AsEnumerable()
                                          where locks.Field<string>("Subsection").TrimEnd() == cboDeptParametersSubsection.Text.Trim()
                                          orderby locks.Field<string>("SubSection"), locks.Field<string>("Department"),
                                          locks.Field<string>("KPF"), locks.Field<string>("KPFParameter") ascending

                                          select locks;
            try
            {

                DataTable temp2 = query2.CopyToDataTable<DataRow>();

                if (temp2.Rows.Count > 0)
                {

                    grdDeptParameters.DataSource = temp2;
                }
                else
                {
                    grdDeptParameters.DataSource = KPF;
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.Message + " The selected subsection has no KPFs.", "Information", MessageBoxButtons.OK);
            }

        }

        private void cboDeptParametersKPF_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboDeptParametersKPFParameters.Items.Clear();
            IEnumerable<DataRow> query1 = from locks in KPF.AsEnumerable()
                                          where locks.Field<string>("KPF").TrimEnd() == cboDeptParametersKPF.Text.Trim()
                                          select locks;
            try
            {

                DataTable temp2 = query1.CopyToDataTable<DataRow>();

                if (temp2.Rows.Count > 0)
                {
                    foreach (DataRow dr in temp2.Rows)
                    {
                        cboDeptParametersKPFParameters.Items.Add(dr["KpfParameter"].ToString().Trim());
                    }

                    cboDeptParametersKPFParameters.Text = cboDeptParametersKPFParameters.Items[0].ToString().Trim();

                }
                else
                {

                }
            }
            catch
            {
                MessageBox.Show("The selected subsection has no KPFs.", "Information", MessageBoxButtons.OK);
            }

        }

        private void btnLockKPFCostLevel_Click_1(object sender, EventArgs e)
        {
            openTab(tabKPFCostLevel);
        }

        private void cboHODDepartment_SelectedIndexChanged(object sender, EventArgs e)
        {
            txtSafetyA.Enabled = true;
            txtTonsPerTECA.Enabled = true;
            txtTonsPerTECP.Enabled = true;
            txtCostA.Enabled = true;
            txtCostP.Enabled = true;

            txtSafetyA.BackColor = Color.Cornsilk;
            txtTonsPerTECA.BackColor = Color.Cornsilk;
            txtTonsPerTECP.BackColor = Color.Cornsilk;
            txtCostA.BackColor = Color.Cornsilk;
            txtCostP.BackColor = Color.Cornsilk;

            ParameterNames.Clear();

            IEnumerable<DataRow> query1 = from locks in SubsectionDept.AsEnumerable()
                                          where locks.Field<string>("SUBSECTION").TrimEnd() == cboHODSubsection.Text.Trim()
                                          where locks.Field<string>("DEPARTMENT").TrimEnd() == cboHODDepartment.Text.Trim()
                                          select locks;

            try
            {
                DataTable temp = query1.CopyToDataTable<DataRow>();
                txtHODHODModel.Text = temp.Rows[0]["HODMODEL"].ToString().Trim();

            }
            catch
            {

            }

            IEnumerable<DataRow> query2 = from locks in DeptParameters.AsEnumerable()
                                          where locks.Field<string>("SUBSECTION").TrimEnd() == cboHODSubsection.Text.Trim()
                                          where locks.Field<string>("DEPARTMENT").TrimEnd() == cboHODDepartment.Text.Trim()
                                          select locks;

            try
            {

                DataTable temp = query2.CopyToDataTable<DataRow>();
                ParameterNames.Clear();
                foreach (DataRow r in temp.Rows)
                {

                    ParameterNames.Add(r["KPFPARAMETER"].ToString().Trim(), r["KPFPARAMETERDESC"].ToString().Trim());
                }

                label999.Text = ParameterNames["ITEM1_ACTUAL"];
                label1000.Text = ParameterNames["ITEM1_PLANNED"];
                label1001.Text = ParameterNames["ITEM2_ACTUAL"];
                label1002.Text = ParameterNames["ITEM2_PLANNED"];
                label1003.Text = ParameterNames["ITEM3_ACTUAL"];
                label1004.Text = ParameterNames["ITEM3_PLANNED"];
                label1005.Text = ParameterNames["ITEM4_ACTUAL"];
                label1006.Text = ParameterNames["ITEM4_PLANNED"];
                label1007.Text = ParameterNames["ITEM5_ACTUAL"];
                label1008.Text = ParameterNames["ITEM5_PLANNED"];


            }
            catch
            {
                label999.Text = "ITEM1_ACTUAL";
                label1000.Text = "ITEM1_PLANNED";
                label1001.Text = "ITEM2_ACTUAL";
                label1002.Text = "ITEM2_PLANNED";
                label1003.Text = "ITEM3_ACTUAL";
                label1004.Text = "ITEM3_PLANNED";
                label1005.Text = "ITEM4_ACTUAL";
                label1006.Text = "ITEM4_PLANNED";
                label1007.Text = "ITEM5_ACTUAL";
                label1008.Text = "ITEM5_PLANNED";
                Application.DoEvents();


            }

            //Extract the mine parameters  
            Dictionary<string, string> MineValues = new Dictionary<string, string>();
            IEnumerable<DataRow> query3 = from locks in MineParameters.AsEnumerable()
                                          select locks;

            try
            {

                DataTable temp = query3.CopyToDataTable<DataRow>();
                foreach (DataColumn dc in temp.Columns)
                {
                    if (dc.ColumnName.Contains("ACTUAL") || dc.ColumnName.Contains("PLANNED"))
                    {
                        MineValues.Add(dc.ColumnName.ToString().Trim(), temp.Rows[0][dc.ColumnName.ToString().Trim()].ToString().Trim());
                    }
                }
            }
            catch
            {
            }

            //Block the input text boxes of if the mine parameters is a parameter in the parameternames dictionary
            foreach (KeyValuePair<string, String> kvp in ParameterNames)
            {
                string parametername = kvp.Key;
                switch (parametername)
                {
                    case "SAFETY_ACTUAL":
                        txtSafetyA.BackColor = Color.AliceBlue;
                        txtSafetyA.Enabled = false;
                        txtSafetyA.Text = MineValues["SAFETY_ACTUAL"];
                        txtMineSafetyA.Text = MineValues["SAFETY_ACTUAL"];

                        break;

                    case "TONS_ACTUAL":

                        txtTonsPerTECA.BackColor = Color.AliceBlue;
                        txtTonsPerTECA.Enabled = false;
                        try
                        {
                            txtTonsPerTECA.Text = MineValues["TEC_ACTUAL"];
                            txtMineTECA.Text = MineValues["TEC_ACTUAL"];
                        }
                        catch
                        {
                            txtTonsPerTECA.Text = "ToBeCalculated";
                            txtMineTECA.Text = "ToBeCalculated";
                        }
                        break;

                    case "TONS_PLANNED":
                        txtTonsPerTECP.BackColor = Color.AliceBlue;
                        txtTonsPerTECP.Enabled = false;
                        txtTonsPerTECP.Text = MineValues["TONSPERTEC_PLANNED"];
                        txtMineTECP.Text = MineValues["TONSPERTEC_PLANNED"];
                        break;

                    case "COST_ACTUAL":
                        txtCostA.BackColor = Color.AliceBlue;
                        txtCostA.Enabled = false;
                        txtCostA.Text = MineValues["COST_ACTUAL"];
                        txtMineCostA.Text = MineValues["COST_ACTUAL"];
                        break;

                    case "COST_PLANNED":
                        txtCostP.BackColor = Color.AliceBlue;
                        txtCostP.Enabled = false;
                        txtCostP.Text = MineValues["COST_PLANNED"];
                        txtMineCostP.Text = MineValues["COST_PLANNED"];
                        break;

                }

            }

            btnUpdate.Enabled = true;
            btnInsertRow.Enabled = false;
            btnDeleteRow.Enabled = false;

        }

        private void HOD_All()
        {

            DialogResult result = MessageBox.Show("The HOD table will be deleted and refreshed. Do you want to continue? - Be SURE!", "Refresh HOD", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    this.Cursor = Cursors.WaitCursor;
                    DataTable HODFromClockedShifts = Base.extractHOD(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType,
                                                 BusinessLanguage.BonusType, txtSelectedSection.Text.Trim(), BusinessLanguage.Period, strWagecodes);

                    if (HODFromClockedShifts.Rows.Count == 0)
                    {
                        MessageBox.Show("No HOD records were extracted from ADTeam.", "Information", MessageBoxButtons.OK);
                    }
                    else
                    {
                        foreach (DataRow row in HODFromClockedShifts.Rows)
                        {
                            DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, row["GANG"].ToString().Trim());
                            if (sub.Rows.Count > 0)
                            {
                                row["SUBSECTION"] = sub.Rows[0]["SUBSECTION"].ToString().Trim();
                                row["DEPARTMENT"] = sub.Rows[0]["DEPARTMENT"].ToString().Trim();
                            }
                            else
                            {
                                row["SUBSECTION"] = "UNKNOWN";
                                row["DEPARTMENT"] = "UNKNOWN";

                            }

                        }

                        TB.saveCalculations2(HODFromClockedShifts, Base.DBConnectionString, "", "HOD");
                        evaluateHOD();

                        MessageBox.Show("HODFromClockedShifts were updated", "Information", MessageBoxButtons.OK);
                        this.Cursor = Cursors.Arrow;
                    }
                    break;

                case DialogResult.No:
                    break;


            }

            this.Cursor = Cursors.Arrow;

        }

        private void cboHOD_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            //HOD - Refresh All
            //HOD - Refresh Shifts
            //HOD - Refresh Dept
            //HOD - Import New

            if (cboHOD.SelectedItem.ToString().Trim() == "HOD - Refresh All")
            {
                HOD_All();
            }
            else
            {
                if (cboHOD.SelectedItem.ToString().Trim() == "HOD - Refresh Shifts")
                {
                    //The HOD shifts are the Shifts_Worked + Q-Shifts.
                    foreach (DataRow row in HOD.Rows)
                    {
                        HOD_RefreshShifts(row["EMPLOYEE_NO"].ToString().Trim(), "316E004", row["GANG"].ToString().Trim());
                    }

                    evaluateHOD();
                }
                else
                {
                    if (cboHOD.SelectedItem.ToString().Trim() == "HOD - Refresh Employees")
                    {
                        //Add only new employees to the HOD table that 
                        this.Cursor = Cursors.WaitCursor;

                        DataTable temp = Base.extractHOD(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType,
                                                     BusinessLanguage.BonusType, txtSelectedSection.Text.Trim(), BusinessLanguage.Period, strWagecodes);

                        foreach (DataRow row in temp.Rows)
                        {
                            foreach (DataRow HODRow in HOD.Rows)
                            {
                                if (row["Employee_No"].ToString().Trim() == HODRow["Employee_No"].ToString().Trim())
                                {
                                    row["Employee_No"] = "XXX";
                                }
                            }
                        }

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            if (temp.Rows[i]["EMPLOYEE_NO"].ToString().Trim() == "XXX")
                            {
                                temp.Rows[i].Delete();
                            }
                            else
                            {
                            }
                        }

                        temp.AcceptChanges();
                        string strDelete = " where Bussunit = '999'";

                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "HOD");

                        evaluateHOD();
                        this.Cursor = Cursors.Arrow;
                    }
                    else
                    {
                        if (cboHOD.SelectedItem.ToString().Trim() == "HOD - Reset HOD Table")
                        {
                            this.Cursor = Cursors.WaitCursor;
                            
                            Base.clearProcessTable(Base.AnalysisConnectionString);
                            Base.updateTableWithNewPeriod(Base.DBConnectionString, "STATUS", BusinessLanguage.Period);
                            Base.updateTableWithNewPeriod(Base.DBConnectionString, "RATES", BusinessLanguage.Period);
                            this.Cursor = Cursors.Arrow;
                        }
                        else
                        {
                            foreach (DataRow row in HOD.Rows)
                            {
                                HOD_RefreshDept(row["Employee_no"].ToString().Trim(), row["Gang"].ToString().Trim());
                            }
                            evaluateHOD();
                        }

                    }

                }
            }

            MessageBox.Show("Done", "Information", MessageBoxButtons.OK);
        }

        private void HOD_RefreshDept(string employee_no, string Gang)
        {
            foreach (DataRow row in SubsectionDept.Rows)
            {

                if (row["HODMODEL"].ToString().Trim() == Gang.Substring(0, row["HODMODEL"].ToString().Trim().Length))
                {
                    TB.InsertData(Base.DBConnectionString,
                                 "Update HOD set Subsection = '" + row["SUBSECTION"].ToString().Trim() +
                                 "', Department  = '" + row["DEPARTMENT"].ToString().Trim() +
                                 "', HODMODEL  = '" + row["HODMODEL"].ToString().Trim() +
                                 "' where employee_no = '" + employee_no + "' AND Gang = '" + Gang + "'");
                }
            }
        }

        private void HOD_RefreshShifts(string employee_no, string Wagecode, string Gang)
        {

            //// Query the Bonusshifts for each HOD's
            IEnumerable<DataRow> query1 = from rec in Labour.AsEnumerable()
                                          where rec.Field<string>("EMPLOYEE_NO").Trim() == employee_no
                                          where rec.Field<string>("WAGECODE").Trim() == Wagecode
                                          where rec.Field<string>("Gang").Trim() == Gang

                                          select rec;
            try
            {
                DataTable testTB = query1.CopyToDataTable<DataRow>();
                if (testTB.Rows.Count == 1)
                {
                    TB.InsertData(Base.DBConnectionString,
                                  "Update HOD set shifts_worked = '" + (Convert.ToInt32(testTB.Rows[0]["Shifts_Worked"].ToString().Trim()) + Convert.ToInt32(testTB.Rows[0]["Q_Shifts"].ToString().Trim())) +
                                  "', Awop_shifts  = '" + testTB.Rows[0]["Awop_Shifts"].ToString().Trim() +
                                  "' where employee_no = '" + employee_no + "' AND Gang = '" + Gang + "'");

                }
                else
                {

                }
            }
            catch
            {


            }


        }

        private void Engineering_RefreshShifts(string employee_no, string Gang, string SUD)
        {
            //// Query the Bonusshifts for each Artisan
            IEnumerable<DataRow> query1 = from rec in Labour.AsEnumerable()
                                          where rec.Field<string>("EMPLOYEE_NO").Trim() == employee_no
                                          where rec.Field<string>("SUD").Trim() == SUD
                                          where rec.Field<string>("Gang").Trim() == Gang

                                          select rec;
            try
            {
                DataTable testTB = query1.CopyToDataTable<DataRow>();
                if (testTB.Rows.Count == 1)
                {
                    switch (SUD)
                    {
                        case "U":
                            TB.InsertData(Base.DBConnectionString,
                                  "Update Artisans set shifts_worked = '" + (Convert.ToInt32(testTB.Rows[0]["Shifts_Worked"].ToString().Trim()) +
                                  Convert.ToInt32(testTB.Rows[0]["Q_Shifts"].ToString().Trim())) +
                                  "', Awop_shifts  = '" + testTB.Rows[0]["Awop_Shifts"].ToString().Trim() +
                                  "' where employee_no = '" + employee_no + "' AND Gang = '" + Gang + "'");

                            break;

                        case "O":
                            TB.InsertData(Base.DBConnectionString,
                                  "Update Officials set shifts_worked = '" + (Convert.ToInt32(testTB.Rows[0]["Shifts_Worked"].ToString().Trim()) +
                                  Convert.ToInt32(testTB.Rows[0]["Q_Shifts"].ToString().Trim())) +
                                  "', Awop_shifts  = '" + testTB.Rows[0]["Awop_Shifts"].ToString().Trim() +
                                  "' where employee_no = '" + employee_no + "' AND Gang = '" + Gang + "'");

                            break;

                    }

                }
                else
                {

                }
            }
            catch
            {


            }

        }

        private void Artisan_All()
        {

            DialogResult result = MessageBox.Show("The Artisan table will be deleted and refreshed. Do you want to continue? - Be SURE!", "Refresh HOD", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    MessageBox.Show("Please be patient. It will take a while.", "Confirm", MessageBoxButtons.OKCancel);
                    this.Cursor = Cursors.WaitCursor;
                    DataTable ArtisanFromClockedShifts = Base.extractEngineeringEmployees(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType,
                                                 BusinessLanguage.BonusType, txtSelectedSection.Text.Trim(), BusinessLanguage.Period, "U");

                    if (ArtisanFromClockedShifts.Rows.Count == 0)
                    {
                        MessageBox.Show("No Artisan records were extracted from ADTeam.", "Information", MessageBoxButtons.OK);
                    }
                    else
                    {
                        foreach (DataRow row in ArtisanFromClockedShifts.Rows)
                        {
                            DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, row["GANG"].ToString().Trim());
                            if (sub.Rows.Count > 0)
                            {
                                row["SUBSECTION"] = sub.Rows[0]["SUBSECTION"].ToString().Trim();
                                row["DEPARTMENT"] = sub.Rows[0]["DEPARTMENT"].ToString().Trim();
                            }
                            else
                            {
                                row["SUBSECTION"] = "UNKNOWN";
                                row["DEPARTMENT"] = "UNKNOWN";

                            }

                        }

                        TB.saveCalculations2(ArtisanFromClockedShifts, Base.DBConnectionString, "", "Artisans");
                        evaluateArtisans();

                        MessageBox.Show("Artisans From ClockedShifts were updated", "Information", MessageBoxButtons.OK);
                        this.Cursor = Cursors.Arrow;
                    }
                    break;

                case DialogResult.No:
                    break;


            }

            this.Cursor = Cursors.Arrow;

        }

        private void Officials_All()
        {

            DialogResult result = MessageBox.Show("The Official table will be deleted and refreshed. Do you want to continue? - Be SURE!", "Refresh HOD", MessageBoxButtons.YesNo);

            switch (result)
            {
                case DialogResult.Yes:
                    MessageBox.Show("Please be patient. It will take a while.", "Confirm", MessageBoxButtons.OKCancel);
                    this.Cursor = Cursors.WaitCursor;
                    DataTable OfficialsFromClockedShifts = Base.extractEngineeringEmployees(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType,
                                                 BusinessLanguage.BonusType, txtSelectedSection.Text.Trim(), BusinessLanguage.Period, "O");

                    if (OfficialsFromClockedShifts.Rows.Count == 0)
                    {
                        MessageBox.Show("No Official records were extracted from ADTeam.", "Information", MessageBoxButtons.OK);
                    }
                    else
                    {
                        foreach (DataRow row in OfficialsFromClockedShifts.Rows)
                        {
                            if (row["WAGECODE"].ToString().Trim() == "316E004")
                            {
                                //This employee is a HOD.
                                //Remove the row.

                            }
                            else
                            {
                                DataTable sub = Base.extractSubsectionAndDept(SubsectionDept, row["GANG"].ToString().Trim());
                                if (sub.Rows.Count > 0)
                                {
                                    row["SUBSECTION"] = sub.Rows[0]["SUBSECTION"].ToString().Trim();
                                    row["DEPARTMENT"] = sub.Rows[0]["DEPARTMENT"].ToString().Trim();
                                }
                                else
                                {
                                    row["SUBSECTION"] = "UNKNOWN";
                                    row["DEPARTMENT"] = "UNKNOWN";

                                }
                            }

                        }

                        TB.saveCalculations2(OfficialsFromClockedShifts, Base.DBConnectionString, "", "Officials");

                        TB.InsertData(Base.DBConnectionString, "Delete from Officials where subsection = 'XXX'");
                        //evaluateOfficials();

                        MessageBox.Show("Officials From ClockedShifts were updated", "Information", MessageBoxButtons.OK);
                        this.Cursor = Cursors.Arrow;
                    }
                    break;

                case DialogResult.No:
                    break;


            }

            this.Cursor = Cursors.Arrow;

        }

        private void txtItem1Planned_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem2Planned_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem3Planned_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem4Planned_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem5Planned_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void button4_Click(object sender, EventArgs e)
        {
            //Extract only employees that has q-shifts
            //Lynx....LINQ

            IEnumerable<DataRow> query1 = from locks in Labour.AsEnumerable()
                                          where Convert.ToInt32(locks.Field<string>("Q_Shifts").TrimEnd()) > 0
                                          select locks;


            try
            {
                DataTable QShifts = query1.CopyToDataTable<DataRow>();
                grdLabour.DataSource = QShifts;
            }
            catch
            {

            }

        }

        private void txtItem1Actual_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem2Actual_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem3Actual_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();

        }

        private void txtItem4Actual_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();
        }

        private void txtItem5Actual_TextChanged(object sender, EventArgs e)
        {
            updateTextBoxes();
        }

        private void btnRefreshMineParameters_Click(object sender, EventArgs e)
        {

            if (txtSafetyA.BackColor == Color.AliceBlue)
            {
                txtSafetyA.Text = txtMineSafetyA.Text;
            }
            if (txtCostA.BackColor == Color.AliceBlue)
            {
                txtCostA.Text = txtMineCostA.Text;
            }
            if (txtCostP.BackColor == Color.AliceBlue)
            {
                txtCostP.Text = txtMineCostP.Text;
            }
            if (txtTonsPerTECA.BackColor == Color.AliceBlue)
            {
                txtTonsPerTECA.Text = txtMineTECA.Text;
            }
            if (txtTonsPerTECP.BackColor == Color.AliceBlue)
            {
                txtTonsPerTECP.Text = txtMineTECP.Text;
            }

        }

        private void btnPrintParameters_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("departmentparameters");
            this.Cursor = Cursors.Arrow;
        }

       

        private void cboArtisans_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Artisans - Refresh All
            //Artisans - Refresh Shifts
            //Artisans - Refresh Dept
            //Artisans - Import New

            if (cboArtisans.SelectedItem.ToString().Trim() == "Artisan - Refresh All")
            {
                Artisan_All();
            }
            else
            {
                if (cboArtisans.SelectedItem.ToString().Trim() == "Artisan - Refresh Shifts")
                {
                    DialogResult result = MessageBox.Show("REFRESH shifts of ARTISANS?", "Confirm", MessageBoxButtons.OKCancel);

                    switch (result)
                    {

                        case DialogResult.OK:
                            MessageBox.Show("Please be patient. It will take a while.", "Confirm", MessageBoxButtons.OKCancel);
                            pictBox.Visible = true;
                            this.Cursor = Cursors.WaitCursor;
                            //The Artisan shifts are the Shifts_Worked + Q-Shifts.
                            foreach (DataRow row in Artisans.Rows)
                            {
                                Engineering_RefreshShifts(row["EMPLOYEE_NO"].ToString().Trim(), row["GANG"].ToString().Trim(), "U");
                            }

                            evaluateArtisans();
                            pictBox.Visible = false;
                            MessageBox.Show("Done", "Information");
                            this.Cursor = Cursors.Arrow;
                            break;

                        case DialogResult.Cancel:
                            break;
                    }


                }
                else
                {
                    if (cboArtisans.SelectedItem.ToString().Trim() == "Artisan - Refresh Employees")
                    {
                        //Add only new employees to the HOD table that 
                        this.Cursor = Cursors.WaitCursor;
                        DataTable temp = Base.extractEngineeringEmployees(Base.DBConnectionString, BusinessLanguage.BussUnit, BusinessLanguage.MiningType,
                                                     BusinessLanguage.BonusType, txtSelectedSection.Text.Trim(), BusinessLanguage.Period, "U");

                        foreach (DataRow row in temp.Rows)
                        {
                            foreach (DataRow ArtisanRow in Artisans.Rows)
                            {
                                if (row["Employee_No"].ToString().Trim() == ArtisanRow["Employee_No"].ToString().Trim())
                                {
                                    row["Employee_No"] = "XXX";
                                }
                            }
                        }

                        for (int i = 0; i <= temp.Rows.Count - 1; i++)
                        {
                            if (temp.Rows[i]["EMPLOYEE_NO"].ToString().Trim() == "XXX")
                            {
                                temp.Rows[i].Delete();
                            }
                            else
                            {
                            }
                        }

                        temp.AcceptChanges();
                        string strDelete = " where Bussunit = '999'";
                        TB.saveCalculations2(temp, Base.DBConnectionString, strDelete, "Artisans");

                        evaluateArtisans();
                        this.Cursor = Cursors.Arrow;
                    }
                    else
                    {
                        foreach (DataRow row in Artisans.Rows)
                        {
                            //Artisan_RefreshDept(row["Employee_no"].ToString().Trim(), row["Gang"].ToString().Trim());
                        }
                        evaluateArtisans();
                    }

                }
            }

        }

        private void grdArtisans_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (e.RowIndex < 0)
            {
            }
            else
            {
                txtArtisanEmployeeNo.Text = grdArtisans["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
                txtArtisanEmployeeName.Text = grdArtisans["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
                txtArtisanPPActual.Text = grdArtisans["PERSONALPERFORMANCE_ACTUAL", e.RowIndex].Value.ToString().Trim();
                cboBonusShiftsWageCode.Text = grdArtisans["WAGECODE", e.RowIndex].Value.ToString().Trim();
                txtArtisanShiftsWorked.Text = grdArtisans["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
                txtArtisanAwops.Text = grdArtisans["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
                txtArtisanHODModel.Text = grdArtisans["HODModel", e.RowIndex].Value.ToString().Trim();
                cboArtisanDepartment.Text = grdArtisans["Department", e.RowIndex].Value.ToString().Trim();
                cboArtisanSubsection.Text = grdArtisans["Subsection", e.RowIndex].Value.ToString().Trim();

            }

            Cursor.Current = Cursors.Arrow;

        }

        private void btnArtisanPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("ENGSERARTISANS");
            this.Cursor = Cursors.Arrow;
        }

        private void btnArtisanParameters_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("ENGARTISANS");
            this.Cursor = Cursors.Arrow;
        }

        private void DBDefault_Click(object sender, EventArgs e)
        {
            TB.TBName = "";
            BackupDB(Base.DBConnectionString, Base.DBName, Base.BackupPath);

        }

        private void BackupDB(string connectionstring, string dbname, string dbPath)
        {
            bool check = false;
            check = Base.backupDatabase3(connectionstring, dbname, dbPath);

            //Copy the file to the C:\drive
            if (check == true)
            {
                MessageBox.Show("Source = " + dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) + "\\ICALC", "X:") +
                                dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak", "Information", MessageBoxButtons.OK);

                Path = dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2), "C:") + dbname +
                       DateTime.Today.ToString("yyyyMMdd") + " \\\\";

                createZipFolder(Path, dbname);

                MessageBox.Show("dest = " + Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak", "Information", MessageBoxButtons.OK);
                check = BusinessLanguage.copyBackupFile(dbPath.ToUpper().Replace(dbPath.ToUpper().Substring(0, 2) +
                        "\\ICALC", "X:") + dbname + DateTime.Today.ToString("yyyyMMdd") + ".bak",
                        Path + dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak");

                if (check == true)
                {
                    string filename = dbname + DateTime.Today.ToString("yyyyMMdd") + "xxx.bak";
                    FastZipCompress(Path + "\\", dbname + DateTime.Today.ToString("yyyyMMdd"));
                    MessageBox.Show("Backup Done to : " + Path, "Information", MessageBoxButtons.OK);
                }
                else
                {
                    MessageBox.Show("Copy unsuccessfull from : " + dbPath.Substring(0, 2) + "   Copy unsuccessfull to :" + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
                }
            }
            else
            {
                MessageBox.Show("Backup unsuccessfull to : " + dbPath.Replace(dbPath.Substring(0, 2), "C:"), "Information", MessageBoxButtons.OK);
            }

            Cursor.Current = Cursors.Arrow;

        }

        private static void FastZipCompress(string pathDBBackup)//JVDW
        {
            //FastZip fZip = new FastZip();

            //fZip.CreateZip("C:\\ZipTest\\test.zip", pathDBBackup, false, ".bak$");

        }

        private static void FastZipCompress(string pathDBBackup, string zipname)
        {
            FastZip fZip = new FastZip();

            fZip.CreateZip("C:\\icalc\\" + zipname + ".zip", pathDBBackup.Replace("xxx.bak", ""), false, ".bak$");

        }

        private bool createZipFolder(string path, string databasename)
        {
            path = Base.BackupPath.Replace(Base.BackupPath.Substring(0, 2), "C:") + "\\" + databasename + DateTime.Today.ToString("yyyyMMdd");
            try
            {
                // Try to create the directory.
                DirectoryInfo di = Directory.CreateDirectory(path);
                return true;
            }
            catch
            {
                return false;
            }
        }

        //private void grdOfficials_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    Cursor.Current = Cursors.WaitCursor;
        //    if (e.RowIndex < 0)
        //    {
        //    }
        //    else
        //    {
        //        txtOfficialsEmployeeNo.Text = grdOfficials["EMPLOYEE_NO", e.RowIndex].Value.ToString().Trim();
        //        txtOfficialsEmployeeName.Text = grdOfficials["EMPLOYEE_NAME", e.RowIndex].Value.ToString().Trim();
        //        txtOfficialsPPActual.Text = grdOfficials["PERSONALPERFORMANCE_ACTUAL", e.RowIndex].Value.ToString().Trim();
        //        txtOfficialsShiftsWorked.Text = grdOfficials["SHIFTS_WORKED", e.RowIndex].Value.ToString().Trim();
        //        txtOfficialsAwops.Text = grdOfficials["AWOP_SHIFTS", e.RowIndex].Value.ToString().Trim();
        //        txtOfficialsHODModel.Text = grdOfficials["HODModel", e.RowIndex].Value.ToString().Trim();
        //        cboOfficialsDepartment.Text = grdOfficials["Department", e.RowIndex].Value.ToString().Trim();
        //        cboOfficialsSubsection.Text = grdOfficials["Subsection", e.RowIndex].Value.ToString().Trim();

        //    }

        //    Cursor.Current = Cursors.Arrow;
        //}

        private void btnOfficialsPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.StartReport("ENGSEROfficials");
            this.Cursor = Cursors.Arrow;
        }

        private void grdRates_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdRates);
            }
        }

        private void grdParticipation_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdParticipation);
            }
        }


        private void btnSelect_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void cboColumnNames_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(newDataTable, cboColumnNames.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboColumnValues.Items.Add(s.Trim());
            }
        }

        private void cboColumnValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboColumnNames.SelectedItem.ToString()).TrimEnd() == cboColumnValues.SelectedItem.ToString()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdActiveSheet.DataSource = temp;

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, TB.TBName, Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdActiveSheet.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdActiveSheet.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                        {
                            grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }
        }

        //private void cboColumnShow_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    grdActiveSheet.Columns[cboColumnShow.SelectedItem.ToString()].DisplayIndex = 0;
        //}

        private void cboParticipationColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(newDataTable, cboParticipationColumns.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboParticipationValues.Items.Add(s.Trim());
            }
        }

        private void cboParticipationValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in newDataTable.AsEnumerable()
                                          where locks.Field<string>(cboParticipationColumns.SelectedItem.ToString()).TrimEnd() == cboParticipationValues.Text.Trim()
                                          select locks;


            DataTable temp = query1.CopyToDataTable<DataRow>();

            grdParticipation.DataSource = temp;

            AConn = Analysis.AnalysisConnection;
            AConn.Open();
            DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, "Participation", Base.AnalysisConnectionString);

            foreach (DataRow dt in tempDataTable.Rows)
            {
                string strValue = dt["Calc_Name"].ToString().Trim();
                int intValue = grdParticipation.Columns.Count - 1;

                for (int i = intValue; i >= 3; --i)
                {
                    string strHeader = grdParticipation.Columns[i].HeaderText.ToString().Trim();
                    if (strValue == strHeader)
                    {
                        for (int j = 0; j <= grdParticipation.Rows.Count - 1; j++)
                        {
                            grdParticipation[i, j].Style.BackColor = Color.Lavender;
                        }
                    }
                }
            }
        }

        //private void cboParticipationShow_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    grdParticipation.Columns[cboColumnShow.SelectedItem.ToString()].DisplayIndex = 0;
        //}

        private void cboDeptParametersColumns_SelectedIndexChanged(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(DeptParameters, cboDeptParametersColumns.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboDeptParametersValues.Items.Add(s.Trim());
            }
        }

        //private void DeptParametersShow_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    grdDeptParameters.Columns[cboColumnShow.SelectedItem.ToString()].DisplayIndex = 0;
        //}

        private void cboDeptParametersValues_SelectedIndexChanged(object sender, EventArgs e)
        {
            IEnumerable<DataRow> query1 = from locks in DeptParameters.AsEnumerable()
                                          where locks.Field<string>(cboDeptParametersColumns.Text).TrimEnd() == cboDeptParametersValues.Text.Trim()
                                          select locks;

            try
            {
                DataTable temp = query1.CopyToDataTable<DataRow>();

                grdDeptParameters.DataSource = temp;

                AConn = Analysis.AnalysisConnection;
                AConn.Open();
                DataTable tempDataTable = Analysis.selectTableFormulas(TB.DBName, "DeptParameters", Base.AnalysisConnectionString);

                foreach (DataRow dt in tempDataTable.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdDeptParameters.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdDeptParameters.Columns[i].HeaderText.ToString().Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdDeptParameters.Rows.Count - 1; j++)
                            {
                                grdDeptParameters[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //private void cboDeptParametersShow_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    grdDeptParameters.Columns[cboColumnShow.SelectedItem.ToString()].DisplayIndex = 0;
        //}

        private void grdMineParameters_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                autoSizeGrid(grdMineParameters);
            }
        }

        private void btnMISDeptParameters_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("deptparametersMIS");
            this.Cursor = Cursors.Arrow;
        }

        private void DB2SpreadSheets(object sender, EventArgs e)
        {
            //The database-tables and formulas will be stored on spreadsheets.

            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("No tables to backup", "Backup Failure", MessageBoxButtons.OK);
            }
            else
            {
                foreach (string s in listBox1.Items)
                {
                    TB.TBName = s.Trim();
                    saveTheSpreadSheet();
                }
            }

            //Extract the formulas of the database
            extractDatabaseFormulas();
            TB.TBName = "";

            Base.backupDatabase3(Base.DBConnectionString, Base.DBName, "c:\\iCalc\\Harmony\\Phakisa\\Development\\Databases\\Backups");

            MessageBox.Show("Backup Done to:  c:\\iCalc\\Harmony\\Phakisa\\Development\\Databases\\Backups ", "Information", MessageBoxButtons.OK);

            string filename = Base.DBName + DateTime.Today.ToString("yyyyMMdd");//JVDW
            string pathDBBackup = "c:\\iCalc\\Harmony\\Phakisa\\Development\\Databases\\Backups\\";//JVDW
            FastZipCompress(pathDBBackup);
        }

        private void DBAnalysis_Click(object sender, EventArgs e)
        {
            TB.TBName = "";
            BackupDB(Base.AnalysisConnectionString, "ANALYSIS", Base.BackupPath);
        }

        private void btnMineParmPrint_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            MetaReportRuntime.App mm = new MetaReportRuntime.App();
            mm.Init(strMetaReportCode);
            mm.ProjectsPath = "c:\\icalc\\Harmony\\Phakisa\\" + strServerPath + "\\REPORTS\\";
            mm.StartReport("EngserMine");
            this.Cursor = Cursors.Arrow;
        }

        private void cboHODDesignation_SelectedIndexChanged(object sender, EventArgs e)
        {

            IEnumerable<DataRow> query1 = from locks in Designations.AsEnumerable()
                                          where locks.Field<string>("DESIGNATION").TrimEnd() == cboHODDesignation.Text.Trim()
                                          select locks;


            try
            {
                DataTable Description = query1.CopyToDataTable<DataRow>();
                txtHODDesignationDesc.Text = Description.Rows[0]["Designation_desc"].ToString().Trim();
            }
            catch
            {
                txtHODDesignationDesc.Text = "UNKNOWN";
            }
        }

        private void cboColumnNames_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            List<string> lstColumnValues = lstNames = TB.loadDistinctValuesFromColumn(newDataTable, cboColumnNames.SelectedItem.ToString());

            foreach (string s in lstColumnValues)
            {
                cboColumnValues.Items.Add(s.Trim());
            }
        }

        private void btnShow_Click(object sender, EventArgs e)
        {
            if (blTablenames == false && listBox1.SelectedItems.Count > 0)
            {
                if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                {
                    grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                {
                    grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                {
                    grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                }


                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    if (listBox1.SelectedItems.Contains(listBox1.Items[i]))
                    {

                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = true;
                    }
                    else
                    {
                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = false;
                    }
                }

                if (grdActiveSheet.Columns.Contains("SECTION"))
                {
                    grdActiveSheet.Columns["SECTION"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains("PERIOD"))
                {
                    grdActiveSheet.Columns["PERIOD"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains(cboColumnNames.Text.Trim()))
                {
                    grdActiveSheet.Columns[cboColumnNames.Text.Trim()].Visible = true;
                }

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void btnResetListBos_Click(object sender, EventArgs e)
        {
            this.Cursor = Cursors.WaitCursor;
            extractDBTableNames(listBox1);

            this.Cursor = Cursors.Arrow;
        }

        private void btnHide_Click(object sender, EventArgs e)
        {
            if (blTablenames == false && listBox1.SelectedItems.Count > 0)
            {
                //unhide first all the columns.
                for (int i = 0; i <= grdActiveSheet.Columns.Count - 1; i++)
                {
                    grdActiveSheet.Columns[i].Visible = true;
                }

                if (grdActiveSheet.Columns.Contains("BUSSUNIT"))
                {
                    grdActiveSheet.Columns["BUSSUNIT"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("MININGTYPE"))
                {
                    grdActiveSheet.Columns["MININGTYPE"].Visible = false;
                }
                if (grdActiveSheet.Columns.Contains("BONUSTYPE"))
                {
                    grdActiveSheet.Columns["BONUSTYPE"].Visible = false;
                }


                for (int i = 0; i <= listBox1.Items.Count - 1; i++)
                {
                    if (listBox1.SelectedItems.Contains(listBox1.Items[i]))
                    {

                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = false;
                    }
                    else
                    {
                        grdActiveSheet.Columns[listBox1.Items[i].ToString().Trim()].Visible = true;
                    }
                }

                if (grdActiveSheet.Columns.Contains("SECTION"))
                {
                    grdActiveSheet.Columns["SECTION"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains("PERIOD"))
                {
                    grdActiveSheet.Columns["PERIOD"].Visible = true;
                }
                if (grdActiveSheet.Columns.Contains(cboColumnNames.Text.Trim()))
                {
                    grdActiveSheet.Columns[cboColumnNames.Text.Trim()].Visible = true;
                }

                foreach (DataRow dt in _formulas.Rows)
                {
                    string strValue = dt["Calc_Name"].ToString().Trim();
                    int intValue = grdActiveSheet.Columns.Count - 1;

                    for (int i = intValue; i >= 3; --i)
                    {
                        string strHeader = grdActiveSheet.Columns[i].HeaderText.Trim();
                        if (strValue == strHeader)
                        {
                            for (int j = 0; j <= grdActiveSheet.Rows.Count - 1; j++)
                            {
                                grdActiveSheet[i, j].Style.BackColor = Color.Lavender;
                            }
                        }
                    }
                }
            }
        }

        private void cboParticipationShow_SelectedIndexChanged(object sender, EventArgs e)
        {

        }



    }

}