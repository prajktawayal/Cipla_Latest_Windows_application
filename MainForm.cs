using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Globalization;
using Cipla_Bangalore_API4_091.BAL;
using System.Data.Odbc;


namespace Cipla_Bangalore_API4_091
{
    public partial class MainForm : Form
    {
        private readonly DAL.DAL DAL;
        #region Declare Globally
        public string queryText;
        public string connetionString1 = string.Empty;
        public static string connetionString = null;
        public static string connectionStringFetch_BatchNo = null;
        string connetionStringAlarmlive = null;
        public String StepNumber = "";
        string ReactorName = null;
        string SolventName = null;
        string DryerName = null;
        string Equip_Name = null;
        public string viewName = null;
        string BatchNo = null;
        public static string BatchEndDateTime = null;
        public static string BatchStartDateTime = null;
        public string logAdminSession = string.Empty;
        public string UserID = string.Empty;
        public static string fromDateTime = "";
        public static string ToDateTime = "";
        string AE_Equip_Name = "";
        public static string rptFromdate = "";
        public static string rptTodate = "";
        public static string rptBatchstarttime = "";
        public static string rptBatchEndtime = "";
        public static string rptBatchNo = "";
        public static string rptProductname = "";
        public static string rptReactorname = "";
        string BatchStartDateTimeTimeIntervalAdded = null;
        public DateTime dtbatchst;
        public string btstrt = "";
        public string btend = "";
        public string RHTempName = "";

        public DateTime dtbatchend;
        private readonly BAL.BAL BAL;
        ReportView vr = new ReportView();

        #endregion
        public MainForm()
        {
            InitializeComponent();
            BAL = new BAL.BAL();
        }

        private void lblClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            MainForm obj = new MainForm();
            obj.UserID = UserID;
            GroupboxReactorVisibitilty();
            LoadTimeIntervals();
        }

        public void GroupboxReactorVisibitilty()
        {
            // Code to comment the Groupbox which are not required for Reactor Batch Report Generation           

            btnEventReport.Enabled = true;
            btnTrendReport.Enabled = true;
            btnAlarmReport.Enabled = true;
            btnAuditTrailReport.Enabled = true;
            if (rbBatchReport.Checked == true)
            {
                grpBx_BatchSelection.Visible = true;
                grpBx_Utility.Visible = false;
                grpBx_Solvent.Visible = false;
                btnBatchReport.Enabled = true;
                btnAllSolvent.Enabled = false;
                btnSolvent.Enabled = false;
            }
            else if (rbUtilityReport.Checked == true)
            {
                grpBx_Utility.Visible = true;
                grpBx_BatchSelection.Visible = false;
                grpBx_Solvent.Visible = false;
                btnBatchReport.Enabled = true;
                btnAllSolvent.Enabled = false;
                btnSolvent.Enabled = false;
            }
            else if (rbSolventReport.Checked == true)
            {
                grpBx_Utility.Visible = false;
                grpBx_BatchSelection.Visible = false;
                grpBx_Solvent.Visible = true;
                btnBatchReport.Enabled = false;
                btnAllSolvent.Enabled = true;
                btnSolvent.Enabled = true;
            }

            grpBx_Dryer.Visible = false;

            //Below are false as they are not required in Reactor Batch Report

            lblVenturiPressAccept.Visible = false;
            lblGrindPressAccept.Visible = false;
            lblVenPresAccp.Visible = false;
            lblGrnPresAcc.Visible = false;
            lblKgcm1.Visible = false;
            lblKgcm2.Visible = false;
            txtGrindPressAccept.Visible = false;
            txtVenturiPressAccept.Visible = false;
            txtGrindPresAccep2.Visible = false;
            txtVentPresAccep2.Visible = false;

            lblRelativeHumidityAccp.Visible = false;
            lblHumidityAccp.Visible = false;
            lblPercentHumidityAccp.Visible = false;
            txtRelHumidityAccp1.Visible = false;
            txtRelHumidityAccp2.Visible = false;

            RHcomboBox.Visible = false;
            rhtemplbl.Visible = false;

            btnTempRHReport.Enabled = false;

            lblAuditLocation.Visible = false;
            cmbAuditLocation.Visible = false;
            btnAuditTrailReport.Visible = false;

        }

        public void GroupboxDryerVisibitilty()
        {
            // Code to comment the Groupbox which are not required for Reactor Batch Report Generation

            grpBx_Utility.Visible = false;
            grpBx_Solvent.Visible = false;

            if (rbDryerReport.Checked == true)
            {
                grpBx_Dryer.Visible = true;
                grpBx_BatchSelection.Visible = false;
            }
            else if (rbCustomReport.Checked == true)
            {
                grpBx_BatchSelection.Visible = true;
                grpBx_Dryer.Visible = false;
            }

            btnBatchReport.Enabled = true;
            btnEventReport.Enabled = true;
            btnTrendReport.Enabled = true;
            btnAlarmReport.Enabled = true;
            btnAuditTrailReport.Enabled = true;

            lblRelativeHumidityAccp.Visible = false;
            lblHumidityAccp.Visible = false;
            lblPercentHumidityAccp.Visible = false;
            txtRelHumidityAccp1.Visible = false;
            txtRelHumidityAccp2.Visible = false;

            RHcomboBox.Visible = false;
            rhtemplbl.Visible = false;

            //Below are false as they are not required in Reactor Batch Report

            lblVenturiPressAccept.Visible = true;
            lblGrindPressAccept.Visible = true;
            lblVenPresAccp.Visible = true;
            lblGrnPresAcc.Visible = true;
            lblKgcm1.Visible = true;
            lblKgcm2.Visible = true;
            txtGrindPressAccept.Visible = true;
            txtVenturiPressAccept.Visible = true;
            txtGrindPresAccep2.Visible = true;
            txtVentPresAccep2.Visible = true;

            btnTempRHReport.Enabled = false;
            btnAllSolvent.Enabled = false;
            btnSolvent.Enabled = false;

            lblAuditLocation.Visible = false;
            cmbAuditLocation.Visible = false;
            btnAuditTrailReport.Visible = false;
        }

        public void GroupboxTempRHVisibitilty()
        {
            // Code to comment the Groupbox which are not required for Reactor Batch Report Generation
            grpBx_BatchSelection.Visible = true;
            grpBx_Utility.Visible = false;
            grpBx_Solvent.Visible = false;
            grpBx_Dryer.Visible = false;
          


            //Below are false as they are not required in Reactor Batch Report

            lblVenturiPressAccept.Visible = false;
            lblGrindPressAccept.Visible = false;
            lblVenPresAccp.Visible = false;
            lblGrnPresAcc.Visible = false;
            lblKgcm1.Visible = false;
            lblKgcm2.Visible = false;
            txtGrindPressAccept.Visible = false;
            txtVenturiPressAccept.Visible = false;
            txtGrindPresAccep2.Visible = false;
            txtVentPresAccep2.Visible = false;

            lblRelativeHumidityAccp.Visible = true;
            lblHumidityAccp.Visible = true;
            lblPercentHumidityAccp.Visible = true;
            txtRelHumidityAccp1.Visible = true;
            txtRelHumidityAccp2.Visible = true;

            RHcomboBox.Visible = true;
            rhtemplbl.Visible = true;

            btnAllSolvent.Enabled = false;
            btnSolvent.Enabled = false;
            btnBatchReport.Enabled = false;
            btnEventReport.Enabled = false;
            btnTrendReport.Enabled = false;

            lblAuditLocation.Visible = true;
            cmbAuditLocation.Visible = true;
            btnAuditTrailReport.Visible = true;

            btnAlarmReport.Enabled = true;
            btnAuditTrail.Enabled = true;
            btnTempRHReport.Enabled = true;
        }

        private void LoadTimeIntervals()
        {
            try
            {
                // Set connection string dynamically based on your requirement
                string connectionString = @"DSN=ANFD;Uid=sa;pwd=Cipla@123";
                DataTable dt = BAL.GetTimeIntervals(connectionString);

                if (dt.Rows.Count > 0)
                {
                    cmbTimeInterval.DisplayMember = "TimeInterval";
                    cmbTimeInterval.ValueMember = "TimeVal";
                    cmbTimeInterval.DataSource = dt;
                }
                else
                {
                    MessageBox.Show("No Record Found!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void btnSearchBatch_Click(object sender, EventArgs e)
        {
            #region Control Validation

            fromDateTime = dtpFromDate.Value.ToString();
            ToDateTime = dtpToDate.Value.ToString();

            if (rbBatchReport.Checked)
            {
                // Validate From and To dates
                if (dtpFromDate.Value >= dtpToDate.Value)
                {
                    MessageBox.Show("From date must be earlier than To date", "Error Detected in Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Exit the method if validation fails
                }

                // Validate Reactor Name
                if (string.IsNullOrEmpty(cmbReactorName.Text))
                {
                    MessageBox.Show("Please select Reactor Name", "Error Detected in Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Exit the method if validation fails
                }

                SearchBatchNumberConnString();
                LoadBatchNo();
            }
            else if (rbDryerReport.Checked)
            {
                // Validate From and To dates
                if (dtpFromDate.Value >= dtpToDate.Value)
                {
                    MessageBox.Show("From date must be earlier than To date", "Error Detected in Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Exit the method if validation fails
                }

                // Validate Reactor Name
                if (string.IsNullOrEmpty(Dryer_cmb.Text))
                {
                    MessageBox.Show("Please select Dryer Name", "Error Detected in Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Exit the method if validation fails
                }

                SearchBatchNumberConnString();
                LoadBatchNo();
            }
            else if (rbSolventReport.Checked == true)
            {
                // Validate From and To dates
                if (dtpFromDate.Value >= dtpToDate.Value)
                {
                    MessageBox.Show("From date must be earlier than To date", "Error Detected in Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Exit the method if validation fails
                }

                // Validate Reactor Name
                if (string.IsNullOrEmpty(cmbSolvent.Text))
                {
                    MessageBox.Show("Please select Reactor Name", "Error Detected in Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return; // Exit the method if validation fails
                }

                SearchBatchNumberConnString();
                LoadBatchNo();
            }

            #endregion
        }

        private void SearchBatchNumberConnString()
        {
            ReactorName = cmbReactorName.Text;
            DryerName = Dryer_cmb.Text;
            SolventName = cmbSolvent.Text;

            connetionString = "";
            Equip_Name = "";
            viewName = "";

            if (rbBatchReport.Checked || rbCustomReport.Checked || rbUtilityReport.Checked)
            {
                switch (ReactorName)
                {
                    case "Room_No_487":
                        connetionString = @"DSN=API_IV_Room_No_487;Uid=sa;pwd=Cipla@123";
                        Equip_Name = "E2R20";
                        viewName = "v_API_IV_Room_No_487";
                        break;
                    case "E4-ANFD-371":
                    case "E4-ANFD-375":
                        connetionString = @"DSN=ANFD;Uid=sa;pwd=Cipla@123";
                        string[] reactorParts = ReactorName.Split('-');
                        viewName = "v_API4_" + reactorParts[1] + reactorParts[2];
                        Equip_Name = reactorParts[0] + reactorParts[1] + reactorParts[2];
                        break;
                    case "E4-ANFD-63":
                        connetionString = @"DSN=ANFD;Uid=sa;pwd=Cipla@123";
                        Equip_Name = "E4ANFD63";
                        viewName = "v_API4_ANFD63";
                        break;
                    case "E3-GLR-19":
                        connetionString = @"DSN=E3R19;Uid=sa;pwd=Cipla@123";
                        Equip_Name = "E3R19";
                        viewName = "v_E3R19";
                        break;
                    case "E4-GLR-04":
                        connetionString = @"DSN=ALL_INTERMEDIATE;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_GLR04";
                        Equip_Name = "E4GLR04";
                        break;
                    case "E4-GLR-06":
                        connetionString = @"DSN=ALL_INTERMEDIATE;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_GLR06";
                        Equip_Name = "E4GLR06";
                        break;
                    case "E4-SSR-07":
                    case "E4-SSR-09":
                    case "E4-SSR-11":
                    case "E4-GLR-12":
                    case "E4-GLR-13":
                    case "E4-GLR-310":
                    case "E4-GLR-17":
                    case "E4-SSR-19":
                    case "E4-GLR-28":
                    case "E4-GLR-01":
                    case "E4-SSR-29":
                    case "E4-SSR-43":
                    case "E4-SSR-44":
                    case "E4-GLR-45":
                    case "E4-GLR-348":
                        connetionString = @"DSN=ALL_INTERMEDIATE;Uid=sa;pwd=Cipla@123";
                        // Extract the relevant part of the ReactorName for the viewName
                        reactorParts = ReactorName.Split('-');
                        viewName = "v_API4_" + reactorParts[1] + reactorParts[2];
                        Equip_Name = reactorParts[0] + reactorParts[1] + reactorParts[2];
                        break;

                    case "E4-GLR-36-PH":
                        connetionString = @"DSN=API_IV_GLR36PH;Uid=sa;pwd=Cipla@123";
                        viewName = "v_GLR36PH";
                        Equip_Name = "E4-GLR-36-PH";
                        break;

                    case "E4-SSR-311":
                        connetionString = @"DSN=API_IV_SSR311;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_SSR311";
                        Equip_Name = "E4SSR311";
                        break;
                    /****** LOGIC FOR CRYSTILLSER DSN CONNECTION STARTS FROM HERE ******/
                    case "E4-GLR-23":
                    case "E4-SSR-24":
                    case "E4-SSR-37":
                    case "E4-SSR-368":
                    case "E4-GLR-32":
                    case "E4-GLR-33":
                    case "E4-GLR-36":
                        connetionString = @"DSN=CRYSTILLSER;Uid=sa;pwd=Cipla@123";
                        reactorParts = ReactorName.Split('-');
                        viewName = "v_API4_" + reactorParts[1] + reactorParts[2];
                        Equip_Name = reactorParts[0] + reactorParts[1] + reactorParts[2];
                        break;
                    /****** LOGIC FOR CRYSTILLSER DSN CONNECTION ENDS HERE ******/
                    /****** LOGIC FOR SOLVENT CONNECTION THROUGH DSN STARTS FROM HERE (SOLVANTS WHEN CLICKED ON CUSTOM REPORT) */

                    case "E4_CF65B_SOLVANT":
                        connetionString = @"DSN=ALL SOLVANT API 4;Uid=sa;pwd=Cipla@123";
                        viewName = "E4_CF65B_SOLVANT";
                        break;
                    case "E4_SSV03_SOLVANT":
                    case "E4_SSV251_SOLVANT":
                    case "E4_SSV347_SOLVANT":
                    case "E4_SSV373_SOLVANT":
                    case "E4_SSV46_SOLVANT":

                        connetionString = @"DSN=ALL SOLVANT API 4;Uid=sa;pwd=Cipla@123";
                        // Extract the relevant part of the ReactorName for the viewName
                        viewName = "v_" + ReactorName;
                        break;

                    /****** LOGIC FOR SOLVENT CONNECTION THROUGH DSN ENDS HERE *****/

                    /****** LOGIC FOR ALL_JTM_MM DSN CONNECTION STARTS HERE *****/
                    case "E4-JTM86":
                        connetionString = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_JTM86";
                        Equip_Name = "JTM-86";
                        break;
                    case "E4-JTM87":
                        connetionString = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_JTM87";
                        Equip_Name = "JTM-87";
                        break;
                    /****** LOGIC FOR ALL_JTM_MM DSN CONNECTION END HERE *****/

                    /****** LOGIC FOR JTM_MM DSN CONNECTION STARTS HERE *****/

                    case "E4-MM-70":
                    case "E4-MM-78":
                    case "E4-MM-82":
                        connetionString = @"DSN=JTM_MM;Uid=sa;pwd=Cipla@123";
                        reactorParts = ReactorName.Split('-');
                        viewName = "v_MM_" + reactorParts[2];
                        Equip_Name = reactorParts[1] + "-" + reactorParts[2];
                        break;

                    /****** LOGIC FOR JTM_MM DSN CONNECTION END HERE *****/

                    /****** LOGIC FOR API_IV_OGB DSN CONNECTION STARTS HERE *****/

                    case "OGB73":
                    case "OGB88":

                        connetionString = @"DSN=API_IV_OGB;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_" + ReactorName;
                        Equip_Name = ReactorName;
                        break;

                    /****** LOGIC FOR API_IV_OGB DSN CONNECTION END HERE *****/

                    case "E4-RCVD-350":
                    case "E4-RCVD-370":
                    case "E4-RCVD-61":
                    case "E4-RCVD-69":
                    case "E4-RCVD-76":

                        connetionString = @"DSN=RCVD;Uid=sa;pwd=Cipla@123";
                        reactorParts = ReactorName.Split('-');
                        viewName = "v_API4_" + reactorParts[1] + reactorParts[2];
                        Equip_Name = reactorParts[0] + reactorParts[1] + reactorParts[2];
                        break;

                    case "E4-HWS-180":
                    case "E4-HWS-181":
                    case "E4-HWS-122":

                        connetionString = @"DSN=ALL_HWS;Uid=sa;pwd=Cipla@123";
                        reactorParts = ReactorName.Split('-');
                        viewName = "v_API4_" + reactorParts[1] + reactorParts[2];
                        Equip_Name = reactorParts[0] + reactorParts[1] + reactorParts[2];
                        break;

                    default:
                        // Handle unknown ReactorName
                        break;
                }
            }
            else if (RHradioButton.Checked)
            {
                switch (RHTempName)
                {
                    case "DRY AREA TEMP RH":
                        connetionString = @"DSN=DRY_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                        viewName = "vFloatTable";
                        break;
                    case "WET AREA TEMP RH":
                        connetionString = @"DSN=WET_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                        viewName = "vFloatTable";
                        break;
                    // Add other cases as needed
                    default:
                        // Handle unknown RHTempName
                        break;
                }
            }
            else if (rbDryerReport.Checked)
            {
                switch (DryerName)
                {
                    case "E4-RCVD-350":
                    case "E4-RCVD-370":
                    case "E4-RCVD-61":
                    case "E4-RCVD-69":
                    case "E4-RCVD-76":

                        connetionString = @"DSN=RCVD;Uid=sa;pwd=Cipla@123";
                        String[] reactorParts = DryerName.Split('-');
                        viewName = "v_API4_" + reactorParts[1] + reactorParts[2];
                        Equip_Name = reactorParts[0] + reactorParts[1] + reactorParts[2];
                        break;
                    /****** LOGIC FOR ALL_JTM_MM DSN CONNECTION STARTS HERE *****/
                    case "E4-JTM86":
                        connetionString = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_JTM86";
                        Equip_Name = "JTM-86";
                        break;
                    case "E4-JTM87":
                        connetionString = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                        viewName = "v_API4_JTM87";
                        Equip_Name = "JTM-87";
                        break;
                    /****** LOGIC FOR ALL_JTM_MM DSN CONNECTION END HERE *****/
                    default:
                        // Handle unknown DryerName
                        break;
                }
            }
            else if (rbSolventReport.Checked)
            {
                switch (SolventName)
                {
                    case "E4_ANFD371_SOLVANT":
                    case "E4_ANFD375_SOLVANT":
                    case "E4_ANFD63_SOLVANT":
                    case "E4_GLR01_SOLVANT":
                    case "E4_GLR04_SOLVANT":
                    case "E4_GLR06_SOLVANT":
                    case "E4_GLR12_SOLVANT":
                    case "E4_GLR13_SOLVANT":
                    case "E4_GLR17_SOLVANT":
                    case "E4_GLR23_SOLVANT":
                    case "E4_GLR28_SOLVANT":
                    case "E4_GLR32_SOLVANT":
                    case "E4_GLR33_SOLVANT":
                    case "E4_GLR348_SOLVANT":
                    case "E4_GLR36_SOLVANT":
                    case "E4_GLR45_SOLVANT":
                    case "E4_SSR07_SOLVANT":
                    case "E4_SSR09_SOLVANT":
                    case "E4_SSR11_SOLVANT":
                    case "E4_SSR19_SOLVANT":
                    case "E4_SSR24_SOLVANT":
                    case "E4_SSR29_SOLVANT":
                    case "E4_SSR368_SOLVANT":
                    case "E4_SSR37_SOLVANT":
                    case "E4_SSR43_SOLVANT":
                    case "E4_SSR44_SOLVANT":
                        connetionString = @"DSN=ALL SOLVANT API 4;Uid=sa;pwd=Cipla@123";
                        viewName = "v_" + SolventName;
                        break;
                    default:
                        // Handle unknown SolventName
                        break;
                }
            }

           

        }
       

        private void LoadBatchNo()
        {
            try
            {
                string connetionString = "";
                string queryText = "";
                string batchColumn = "";

                // Determine connection string, query, and batch column based on selected radio button
                DetermineQueryParameters(ref connetionString, ref queryText, ref batchColumn);

                // Retrieve batch data from the database
                DataTable batchData = BAL.FetchBatchNoData(connetionString, queryText);

                // Bind the batch data to the ComboBox
                BindBatchData(batchData, batchColumn);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void DetermineQueryParameters(ref string connectionStringFetch_BatchNo, ref string queryText, ref string batchColumn)
        {
            if (rbBatchReport.Checked)
            {
                if (cmbReactorName.Text == "E4-MM-70" || cmbReactorName.Text == "E4-MM-78" || cmbReactorName.Text == "E4-MM-82")
                {
                    connectionStringFetch_BatchNo = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                    queryText = "Select Distinct(BatchNo) as BatchNo from " + viewName + " where DateAndTime between '" + fromDateTime + "' and '" + ToDateTime + "' and BatchNo !='' ";
                    batchColumn = "BatchNo";
                }
                else if (cmbReactorName.Text == "OGB73" || cmbReactorName.Text == "OGB88")
                {
                    connectionStringFetch_BatchNo = @"DSN=API_IV_OGB;Uid=sa;pwd=Cipla@123";
                    queryText = "Select Distinct(BatchNo) as BatchNo from " + viewName + " where DateAndTime between '" + fromDateTime + "' and '" + ToDateTime + "' and BatchNo !='' ";
                    batchColumn = "BatchNo";
                }
                else
                {
                    connectionStringFetch_BatchNo = @"DSN=HMI_Alarm_Event;Uid=sa;pwd=Cipla@123";
                    queryText = "Select Distinct(Tag3Value) as BatchNo from UDV_AllEvents where EventTimeStamp between '" + fromDateTime + "' and '" + ToDateTime + "' and Tag3Value !='' and Tag4Value='" + ReactorName + "' ";
                    batchColumn = "BatchNo";
                }
            }
            else if (rbDryerReport.Checked)
            {
                connectionStringFetch_BatchNo = @"DSN=HMI_Alarm_Event;Uid=sa;pwd=Cipla@123";
                if (DryerName == "E4-JTM86" || DryerName == "E4-JTM87")
                {
                    queryText = "Select Distinct BatchNo from UDV_Events_Jetmill where EventTimeStamp between '" + fromDateTime + "' and '" + ToDateTime + "' and BatchNo !='' and Equipment_Name='" + DryerName + "' ";
                }
                else
                {
                    queryText = "Select Distinct(Tag3Value) as BatchNo from UDV_AllEvents where EventTimeStamp between '" + fromDateTime + "' and '" + ToDateTime + "' and Tag3Value !='' and Tag4Value='" + DryerName + "' ";
                }
                batchColumn = "BatchNo";
            }
            else if (rbSolventReport.Checked)
            {
                connectionStringFetch_BatchNo = @"DSN=ALL SOLVANT API 4;Uid=sa;pwd=Cipla@123";
                queryText = "Select Distinct(BatchNo) as BatchNo from " + viewName + " where DateAndTime between '" + fromDateTime + "' and '" + ToDateTime + "' and BatchNo !='' ";
                batchColumn = "BatchNo";
            }
        }

        private void BindBatchData(DataTable batchData, string batchColumn)
        {
            if (batchData.Rows.Count > 0)
            {
                cmbBatchNo.DisplayMember = batchColumn;
                cmbBatchNo.ValueMember = batchColumn;
                cmbBatchNo.DataSource = batchData;
            }
            else
            {
                cmbBatchNo.DisplayMember = null;
                cmbBatchNo.ValueMember = null;
                MessageBox.Show("No Record Found!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void rbDryerReport_CheckedChanged(object sender, EventArgs e)
        {
            GroupboxDryerVisibitilty();
        }

        private void RHradioButton_CheckedChanged(object sender, EventArgs e)
        {
            GroupboxTempRHVisibitilty();
        }

        private void rbCustomReport_CheckedChanged(object sender, EventArgs e)
        {
            GroupboxDryerVisibitilty();
    
        }
        private void rbBatchReport_CheckedChanged(object sender, EventArgs e)
        {
            GroupboxReactorVisibitilty();
        }

        private void rbUtilityReport_CheckedChanged(object sender, EventArgs e)
        {
            GroupboxReactorVisibitilty();
        }

        private void rbSolventReport_CheckedChanged(object sender, EventArgs e)
        {
            GroupboxReactorVisibitilty();
        }

        private void btnBatchReport_Click(object sender, EventArgs e)
        {
            string timeInterval = cmbTimeInterval.Text;
            string customTA = custta.Text;
            string customVA = custva.Text;
            string customRPMA = custrpma.Text;
            string tAcceptance = ta.Text;
            string stepNum = stepno.Text;
            string venturiAcceptance1 = txtVenturiPressAccept.Text;
            string venturiAcceptance2 = txtVentPresAccep2.Text;
            string grindingAcceptance1 = txtGrindPressAccept.Text;
            string grindingAcceptance2 = txtGrindPresAccep2.Text;

            try
            {
                timeInterval = cmbTimeInterval.SelectedValue.ToString();
                string reportType = "";
                string stepNumber = StepNocomboBox.Text;

                //Batchstart();

                if (string.IsNullOrEmpty(BatchStartDateTime) || string.IsNullOrEmpty(BatchEndDateTime))
                {
                    MessageBox.Show("Cannot generate Batch report. Please check Batch Start/End Date Time is correctly logged!");
                    return;
                }

                BatchStartDateTime = Convert.ToDateTime(BatchStartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                BatchEndDateTime = Convert.ToDateTime(BatchEndDateTime).ToString("yyyy-MM-dd HH:mm:ss");

                ReportView vr = new ReportView();

                if (rbDryerReport.Checked)
                {
                    reportType = "DryerReport";
                    vr.GenerateDryerReport(
                        BatchStartDateTime, BatchEndDateTime, DryerName, viewName, BatchNo, connetionString,
                        reportType, stepNumber, UserID, "", customTA, customVA, customRPMA, tAcceptance,
                        stepNum, timeInterval, venturiAcceptance1, venturiAcceptance2, grindingAcceptance1, grindingAcceptance2
                    );
                }
                else if (rbBatchReport.Checked)
                {
                    reportType = "BatchReport";
                    vr.GenerateBatchReport(
                        BatchStartDateTime, BatchEndDateTime, ReactorName, viewName, BatchNo,
                        timeInterval, connetionString, reportType, stepNumber, UserID, "",
                        customTA, customVA, customRPMA, tAcceptance, stepNum, venturiAcceptance1, venturiAcceptance2, grindingAcceptance1, grindingAcceptance2
                    );
                }
                else if (rbCustomReport.Checked || rbUtilityReport.Checked)
                {
                    string fromDateTime = DateTime.ParseExact(dtpFromDate.Text, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture)
                                          .ToString("yyyy-MM-dd HH:mm:ss");
                    string toDateTime = DateTime.ParseExact(dtpToDate.Text, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture)
                                        .ToString("yyyy-MM-dd HH:mm:ss");
                    ReactorName = cmbReactorName.Text;
                    reportType = "BatchReport";

                    string customBmrSection = custbmrsection.Text;
                    string customBmrPageNumber = custbmrnumber.Text;
                    string customProductName = custproductname.Text;
                    string customBatchNumber = cmbBatchNo.Text;
                    string batchNo = cmbBatchNo.Text;

                    if (rbCustomReport.Checked)
                    {
                        vr.GenerateCustomBatchReport(
                            fromDateTime, toDateTime, batchNo, ReactorName, viewName, BatchNo,
                            timeInterval, connetionString, reportType, stepNumber, UserID, "",
                            customBatchNumber, customBmrPageNumber, customBmrSection, customProductName, customTA,
                            customVA, customRPMA, tAcceptance, stepNum, venturiAcceptance1, venturiAcceptance2,
                            grindingAcceptance1, grindingAcceptance2
                        );
                    }
                    else if (rbUtilityReport.Checked)
                    {
                        string HWSName = utilitycomboBox.Text;
                        vr.GenerateCustomBatchReport(
                            fromDateTime, toDateTime, batchNo, HWSName, viewName, BatchNo,
                            timeInterval, connetionString, reportType, stepNumber, UserID, "",
                            customBatchNumber, customBmrPageNumber, customBmrSection, customProductName, customTA,
                            customVA, customRPMA, tAcceptance, stepNum, venturiAcceptance1, venturiAcceptance2,
                            grindingAcceptance1, grindingAcceptance2
                        );
                    }
                }
                else
                {
                    MessageBox.Show("Please Select Report Type", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private void cmbBatchNo_SelectedIndexChanged(object sender, EventArgs e)
        {
            BatchNo = cmbBatchNo.Text;
            Batchstart();
        }
        public void Batchstart()
        {
            if (rbBatchReport.Checked == true)
            {

                string reactorName = cmbReactorName.Text;

                connetionString1 = GetConnectionString();
                FetchBatchData(reactorName, connetionString1);
            }
            else if (rbDryerReport.Checked == true)
            {
                connetionString1 = GetConnectionString();
                string DryerName = Dryer_cmb.Text;
                FetchBatchData(DryerName, connetionString1);
            }
            else if (rbSolventReport.Checked == true)
            {
                connetionString1 = GetConnectionString();
                string SolventName = cmbSolvent.Text;
                FetchBatchData(SolventName, connetionString1);
            }
        }
        private void FetchBatchData(string reactorOrDryerName, string connetionString1)
        {
            try
            {
                connectionStringFetch_BatchNo = connetionString1;
                string frmdt = dtpFromDate.Text;
                string tdt = dtpToDate.Text;

                DateTime dtfromdt = DateTime.ParseExact(frmdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                DateTime dttodt = DateTime.ParseExact(tdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                string fromDateTime = dtfromdt.ToString("yyyy-MM-dd HH:mm:ss");
                string toDateTime = dttodt.ToString("yyyy-MM-dd HH:mm:ss");

                string batchNo = cmbBatchNo.Text;

                DataTable dtBatchData = FetchBatchData(connectionStringFetch_BatchNo, reactorOrDryerName, fromDateTime, toDateTime, batchNo, viewName);

                if (dtBatchData.Rows.Count > 0)
                {
                    BatchStartDateTime = dtBatchData.Rows[0]["BatchStartDateTime"].ToString();
                    BatchEndDateTime = dtBatchData.Rows[0]["BatchEndDateTime"].ToString();

                    dtpBatchStartDttm.Text = BatchStartDateTime;
                    dtpBatchEndDttm.Text = BatchEndDateTime;

                    DateTime dtbatchst = Convert.ToDateTime(BatchStartDateTime);
                    DateTime dtbatchend = Convert.ToDateTime(BatchEndDateTime);

                    string btstrt = dtbatchst.ToString("yyyy-MM-dd HH:mm:ss");
                    string btend = dtbatchend.ToString("yyyy-MM-dd HH:mm:ss");

                    DataTable dtprodname = BAL.GetProductName(btstrt, btend, reactorOrDryerName, connectionStringFetch_BatchNo, viewName); // Uncomment and implement if needed
                    string productName = "";
                    if (dtprodname.Rows.Count > 0)
                    {
                        productName = dtprodname.Rows[0]["Val"].ToString();
                        lblProdName.Text = productName;
                    }
                    else
                    {
                        lblProdName.Text = "NA";
                    }
                }
                else
                {
                    MessageBox.Show("Batch Start event not Found!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        public string GetConnectionString()
        {
            if (rbBatchReport.Checked == true)
            {
                if (cmbReactorName.Text == "OGB73" || cmbReactorName.Text == "OGB88")
                {
                    return @"DSN=API_IV_OGB;Uid=sa;pwd=Cipla@123";
                }
                else if (cmbReactorName.Text == "E4-MM-70" || cmbReactorName.Text == "E4-MM-78" || cmbReactorName.Text == "E4-MM-82")
                {
                    return @"DSN=JTM_MM;Uid=sa;pwd=Cipla@123";
                }
                else
                {
                    return @"DSN=HMI_Alarm_Event;Uid=sa;pwd=Cipla@123";
                }
            }
            else if (rbDryerReport.Checked == true)
            {
                return @"DSN=HMI_Alarm_Event;Uid=sa;pwd=Cipla@123";
            }
            else if (rbSolventReport.Checked == true)
            {
                return @"DSN=ALL SOLVANT API 4;Uid=sa;pwd=Cipla@123";
            }
            else
            {
                // Add a default return value or throw an exception if no conditions are met
                return string.Empty; // or throw new InvalidOperationException("No valid report type selected.");
            }
        }

        private void btnAllSolvent_Click(object sender, EventArgs e)
        {

            string reportType = "";
            string stepNumber = StepNocomboBox.Text;
            string hr = "";
            SolventName = cmbSolvent.Text;

            if (string.IsNullOrEmpty(BatchStartDateTime) || string.IsNullOrEmpty(BatchEndDateTime))
            {
                MessageBox.Show("Cannot generate Batch report. Please check Batch Start/End Date Time is correctly logged!");
                return;
            }
            string connectionString = GetConnectionString();
            if (string.IsNullOrEmpty(connectionString))
            {
                MessageBox.Show("Connection string is not defined. Please check your configuration.");
                return;
            }

            // Ensure other required variables are not null or empty

            reportType = "AllSolventReport";
            ReportView vr = new ReportView();

            vr.GenerateAllSolventReport(
                BatchStartDateTime,
                BatchEndDateTime,
                SolventName,
                viewName,
                BatchNo,
                connectionString, // corrected the typo here
                reportType

            );
        }



        public DataTable FetchBatchData(string connectionString, string reactorOrDryerName, string fromDateTime, string toDateTime, string batchNo, string viewName)
        {


            if (reactorOrDryerName == "E4-JTM86" || reactorOrDryerName == "E4-JTM87")
            {
                queryText = $"SELECT Min(EventTimeStamp) as BatchStartDateTime, Max(EventTimeStamp) as BatchEndDateTime FROM UDV_Events_Jetmill WHERE GroupPath='{reactorOrDryerName}_E' AND EventTimeStamp BETWEEN '{fromDateTime}' AND '{toDateTime}' AND BatchNo='{batchNo}'";
            }
            else if (reactorOrDryerName == "E4-MM-70" || reactorOrDryerName == "E4-MM-78" || reactorOrDryerName == "E4-MM-82")
            {
                queryText = $"SELECT Min(DateAndTime) as BatchStartDateTime, Max(DateAndTime) as BatchEndDateTime FROM {viewName} WHERE BatchNo='{batchNo}' AND DateAndTime BETWEEN '{fromDateTime}' AND '{toDateTime}'";
            }
            else if (rbSolventReport.Checked == true)
            {
                queryText = $"SELECT Min(DateAndTime) as BatchStartDateTime, Max(DateAndTime) as BatchEndDateTime FROM {viewName} WHERE BatchNo='{batchNo}' AND DateAndTime BETWEEN '{fromDateTime}' AND '{toDateTime}'";
            }
            else
            {
                queryText = $"SELECT Min(EventTimeStamp) as BatchStartDateTime, Max(EventTimeStamp) as BatchEndDateTime FROM UDV_AllEvents WHERE Tag3Value='{batchNo}' AND Tag4Value='{reactorOrDryerName}' AND EventTimeStamp BETWEEN '{fromDateTime}' AND '{toDateTime}'";
            }
            DataTable result = BAL.FetchBatchData(connectionString, queryText);
            return result;
        }

        private void btnSolvent_Click(object sender, EventArgs e)
        {
            string reportType = "";
            string stepNumber = StepNocomboBox.Text;
            string hr = "";
            SolventName = cmbSolvent.Text;

            if (string.IsNullOrEmpty(BatchStartDateTime) || string.IsNullOrEmpty(BatchEndDateTime))
            {
                MessageBox.Show("Cannot generate Batch report. Please check Batch Start/End Date Time is correctly logged!");
                return;
            }
            string connectionString = GetConnectionString();
            if (string.IsNullOrEmpty(connectionString))
            {
                MessageBox.Show("Connection string is not defined. Please check your configuration.");
                return;
            }

            // Ensure other required variables are not null or empty

            reportType = "SolventReport";
            ReportView vr = new ReportView();

            vr.GenerateSolventReport(
                BatchStartDateTime,
                BatchEndDateTime,
                SolventName,
                viewName,
                BatchNo,
                connectionString, // corrected the typo here
                reportType

            );
        }

        private void btnTrendReport_Click(object sender, EventArgs e)
        {
            if (rbSolventReport.Checked == true)
            {
                string reporttype = "SolventTrendReport";
                ReportView vr = new ReportView();
                string timeInterval = cmbTimeInterval.SelectedValue.ToString();
                //int TimeInterval = vr.ConvertTimeIntervalToSeconds(timeInterval);
                vr.GenerateSolventTrendReport(BatchStartDateTime, BatchEndDateTime, SolventName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval);
            }
            else if (rbDryerReport.Checked == true)
            {
                string timeInterval = cmbTimeInterval.SelectedValue.ToString();
                string stepnum = stepno.Text;
                string reporttype;
                switch (DryerName)
                {
                    case "E4-RCVD-350":
                    case "E4-RCVD-69":
                    case "E4-RCVD-61":
                    case "E4-RCVD-76":
                    case "E4-RCVD-370":
                        reporttype = "DryerTrendReport";
                        break;
                    case "E4-JTM86":
                        reporttype = "TrendReportE4_JMM_86";
                        break;
                    case "E4-JTM87":
                        reporttype = "TrendReportE4_JMM_87";
                        break;
                    default:
                        //Handle unknown DryerName here
                        reporttype = ""; // or any default value you want
                        break;
                }

                ReportView vr = new ReportView();
                vr.GenerateDryerTrendReport(BatchStartDateTime, BatchEndDateTime, DryerName, viewName, BatchNo, timeInterval, connetionString, reporttype, UserID, stepnum);
            }
            else if (rbBatchReport.Checked == true)
            {
                string reporttype= "BatchReport";
                string timeInterval = cmbTimeInterval.SelectedValue.ToString();
                string customta = custta.Text;
                string customva = custva.Text;
                string customrpma = custrpma.Text;
                string tacceptance = ta.Text;
                string stepnum = stepno.Text;
                vr.GenerateTrendReport(BatchStartDateTime, BatchEndDateTime, SolventName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval, customta, customva, customrpma, tacceptance, stepnum);
            }
            else if (rbCustomReport.Checked == true)
            {
                string frmdt = dtpFromDate.Text;
                string tdt = dtpToDate.Text;
                ConvertDateTimeForReportFormat(frmdt, tdt);

                string reporttype = "SolventCustomTrendReport";
                SearchBatchNumberConnString();
                string timeInterval = cmbTimeInterval.SelectedValue.ToString();
                vr.GenerateSolventTrendReport(fromDateTime, ToDateTime, SolventName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval);

            }
        }

        private void cmbTimeInterval_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnEventReport_Click(object sender, EventArgs e)
        {
            ReportView vr = new ReportView();
            string timeInterval = cmbTimeInterval.SelectedValue.ToString();
            if(rbSolventReport.Checked == true)
            {
                string reporttype = "SolventEventReport";
                vr.SolventEventReport(BatchStartDateTime, BatchEndDateTime, SolventName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval);
            }
            else if (rbDryerReport.Checked == true)
            {
                string stepnum = stepno.Text;
                string reporttype;
                switch (DryerName)
                {
                    case "E4-RCVD-350":
                    case "E4-RCVD-69":
                    case "E4-RCVD-61":
                    case "E4-RCVD-76":
                    case "E4-RCVD-370":
                        reporttype = "DryerEventReport";
                        break;
                    case "E4-JTM86":
                        reporttype = "DryerEventReport-JMM";
                        break;
                    case "E4-JTM87":
                        reporttype = "DryerEventReport-JMM";
                        break;
                    default:
                        // Handle unknown DryerName here
                        reporttype = ""; // or any default value you want
                        break;
                }


                vr.DryerEventReport(BatchStartDateTime, BatchEndDateTime, DryerName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval);
            }
            else if (rbBatchReport.Checked == true)
            {
                string customta = custta.Text;
                string customva = custva.Text;
                string customrpma = custrpma.Text;
                string tacceptance = ta.Text;
                string stepnum = stepno.Text;
                string reporttype;
                timeInterval = cmbTimeInterval.SelectedValue.ToString();
                vr = new ReportView();

                if (cmbReactorName.Text == "E4-ANFD-63" || cmbReactorName.Text == "E4-ANFD-375")
                {
                    reporttype = "DryerEventReportANFD";
                    vr.GenerateEventBatchReport(BatchStartDateTime, BatchEndDateTime, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval,  customta,  customva,  customrpma,  tacceptance,  stepnum);
                }
                else if (cmbReactorName.Text == "E4-ANFD-371")
                {
                    reporttype = "DryerEventReportANFD";
                    vr.GenerateEventBatchReport(BatchStartDateTime, BatchEndDateTime, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval, customta, customva, customrpma, tacceptance, stepnum);
                }
                else if (cmbReactorName.Text == "OGB88")
                {
                    reporttype = "EventReportOGB88";
                    vr.GenerateEventBatchReport(BatchStartDateTime, BatchEndDateTime, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval, customta, customva, customrpma, tacceptance, stepnum);

                }
                else if (cmbReactorName.Text == "OGB73")
                {
                    reporttype = "EventReportOGB73";
                    vr.GenerateEventBatchReport(BatchStartDateTime, BatchEndDateTime, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval, customta, customva, customrpma, tacceptance, stepnum);

                }
                else
                {
                    reporttype = "EventReport";
                    vr.GenerateEventBatchReport(BatchStartDateTime, BatchEndDateTime, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reporttype, timeInterval, customta, customva, customrpma, tacceptance, stepnum);

                }
            }
            else if (rbCustomReport.Checked == true)
            {
                string frmdt = dtpFromDate.Text;
                string tdt = dtpToDate.Text;
                string reactorName = cmbReactorName.Text;

                ConvertDateTimeForReportFormat(frmdt, tdt);

                string reporttype = "SolventCustomEventReport";
                vr.SolventCustomEventReport(fromDateTime, ToDateTime,  reactorName, viewName, BatchNo,  connetionString,  StepNumber, AE_Equip_Name,  UserID,  reporttype,  timeInterval);
            }
        }

        private void btnAuditTrail_Click(object sender, EventArgs e)
        {
            cmbAuditLocation.Visible = true;
            lblAuditLocation.Visible = true;
            btnAuditTrailReport.Visible = true;
        }

        private void btnAuditTrailReport_Click(object sender, EventArgs e)
        {
            string ProductName = lblProdName.Text;
            string BatchNo = cmbBatchNo.Text;
            string ReactorName = cmbReactorName.Text;
            string HWSName = utilitycomboBox.Text;

            String BatchStartDateTime1 = dtpBatchStartDttm.Text;
            String BatchEndtDateTime1 = dtpBatchEndDttm.Text;
            String AE_Equip_Name = Equip_Name;
            string areaName = cmbAuditLocation.Text;
            StepNumber = StepNocomboBox.Text;
            string frmdt = dtpFromDate.Text;
            string tdt = dtpToDate.Text;
            string CustomReactorReport = "";
            DateTime dtfromdt1 = DateTime.ParseExact(frmdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            DateTime dttodt1 = DateTime.ParseExact(tdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            string fromDateTime1 = dtfromdt1.ToString("yyyy-MM-dd HH:mm:ss");
            string ToDateTime1 = dttodt1.ToString("yyyy-MM-dd HH:mm:ss");
            string audittable = "";

            try
            {
                //if (BatchStartDateTime1 == "" && BatchEndtDateTime1 == "" && Eventradiobtn.Checked == false && Location == "")
                if (BatchStartDateTime1 == "" && BatchEndtDateTime1 == "" && areaName == "")
                {
                    MessageBox.Show("Can not generate Batch report. Please check Batch Start/End Date Time is correct logged!");
                }
                else
                {
                    DateTime dtfromdt = DateTime.ParseExact(BatchStartDateTime1, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                    DateTime dttodt = DateTime.ParseExact(BatchEndtDateTime1, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);

                    fromDateTime = dtfromdt.ToString("yyyy-MM-dd HH:mm:ss");
                    ToDateTime = dttodt.ToString("yyyy-MM-dd HH:mm:ss");

                }


                #region declare ConnectionString areaWise

                string reporttype = "";
                switch (areaName)
                {
                    case "INBLRENGTB-001":
                    case "INBLRENGTB-002":
                    case "INBLRENGTB-003":
                        connetionString1 = @"DSN=DSN_AUDIT_COMBINE;Uid=sa;pwd=Cipla@123";
                        audittable = "Audit1";
                        reporttype = "AuditReport";
                        break;
                    case "INBLRENGW-090":
                        connetionString1 = @"DSN=INBLRENGW_090_New;Uid=sa;pwd=Cipla@123";
                        audittable = "INBLRENGW_090";
                        reporttype = "AuditReport_INBLRENGW_090";
                        break;
                    case "INBLRENGW-091":
                        connetionString1 = @"DSN=INBLRENGW_091_New;Uid=sa;pwd=Cipla@123";
                        audittable = "INBLRENGW_091";
                        reporttype = "AuditReport";
                        break;
                }
                #endregion
                ReportView vr = new ReportView();
                if (rbBatchReport.Checked == true)
                {

                    vr.GenerateAuditReportINBLRENGW091(fromDateTime1, ToDateTime1, areaName, connetionString1, ProductName, ReactorName, BatchNo, connetionString, viewName, AE_Equip_Name, UserID, audittable, reporttype);

                }

                if (rbCustomReport.Checked == true)
                {
                    reporttype = "AuditReport";
                    vr.GenerateAuditReport(fromDateTime1, ToDateTime1, areaName, connetionString1, ProductName, SolventName, BatchNo, connetionString, viewName, AE_Equip_Name, UserID, audittable, reporttype);

                }
                else if (rbDryerReport.Checked == true)
                {

                    reporttype = "AuditReport";
                    vr.GenerateAuditReportINBLRENGW091(fromDateTime1, ToDateTime1, areaName, connetionString1, ProductName, ReactorName, BatchNo, connetionString, viewName, AE_Equip_Name, UserID, audittable, reporttype);

                }
                else if (rbUtilityReport.Checked == true)
                {
                    reporttype = "AuditReport";
                    vr.GenerateAuditReport(fromDateTime1, ToDateTime1, areaName, connetionString1, ProductName, HWSName, BatchNo, connetionString, viewName, AE_Equip_Name, UserID, audittable, reporttype);

                }
                else if (RHradioButton.Checked == true)
                {
                    string tenprh = RHcomboBox.Text;
                    reporttype = "AuditReport";
                    vr.GenerateAuditReport(fromDateTime1, ToDateTime1, areaName, connetionString1, ProductName, tenprh, BatchNo, connetionString, viewName, AE_Equip_Name, UserID, audittable, reporttype);

                }




            }
            catch (Exception ex)
            {
                //LogError(ex);
            }
        }

        private void btnTempRHReport_Click(object sender, EventArgs e)
        {
            string timeInterval = cmbTimeInterval.Text;
            string customTA = custta.Text;
            string customVA = custva.Text;
            string customRPMA = custrpma.Text;
            string tAcceptance = ta.Text;
            string stepNum = stepno.Text;
           

            try
            {
                timeInterval = cmbTimeInterval.SelectedValue.ToString();
                string reportType = "";
                string stepNumber = StepNocomboBox.Text;
                string RHTemp = RHcomboBox.Text;
                string ReportType = "";
                string customta = custta.Text;
                string tacceptance = ta.Text;
                string relativeaccpt = txtRelHumidityAccp1.Text;
                string relativehumaccpt = txtRelHumidityAccp2.Text;

                //Batchstart();

                if (string.IsNullOrEmpty(BatchStartDateTime) || string.IsNullOrEmpty(BatchEndDateTime))
                {
                    MessageBox.Show("Cannot generate Batch report. Please check Batch Start/End Date Time is correctly logged!");
                    return;
                }

                BatchStartDateTime = Convert.ToDateTime(BatchStartDateTime).ToString("yyyy-MM-dd HH:mm:ss");
                BatchEndDateTime = Convert.ToDateTime(BatchEndDateTime).ToString("yyyy-MM-dd HH:mm:ss");

                ReportView vr = new ReportView();


                    if (RHTemp == "DRY AREA TEMP RH")
                    {
                        ReportType = "DRYRHTemp";
                        connetionString = @"DSN=DRY_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                        viewName = "vFloatTable";
                    }
                    if (RHTemp == "WET AREA TEMP RH")
                    {
                        connetionString = @"DSN=WET_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                        viewName = "vFloatTable";
                        ReportType = "WETRHTemp";
                    }
                    if (RHTemp == "COLD ROOM 57")
                    {
                        connetionString = @"DSN=WET_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                        viewName = "vFloatTable";
                        ReportType = "COLD_ROOM_57";
                    }
                    if (RHTemp == "COLD ROOM 92")
                    {
                        connetionString = @"DSN=DRY_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                        viewName = "vFloatTable";
                        ReportType = "COLD_ROOM_92";
                    }

                    if (RHradioButton.Checked == true)
                    {
                        vr.GenerateRHTEMP(fromDateTime, ToDateTime, tacceptance, relativeaccpt, relativehumaccpt, customta, RHTemp, viewName, BatchNo, timeInterval, connetionString, ReportType, StepNumber, UserID, customRPMA, customTA);
                    }

                else
                {
                    MessageBox.Show("Please Select Report Type", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            AuditPanel.Visible = false;
           
        }

        private void btnAlarmReport_Click(object sender, EventArgs e)
        {
            string timeInterval = cmbTimeInterval.SelectedValue.ToString();
            string reportType = "";
            string customta = custta.Text;
            string customva = custva.Text;
            string customrpma = custrpma.Text;
            string tacceptance = ta.Text;
            string stepnum = stepno.Text;
            string RHTemp = RHcomboBox.Text;
            AE_Equip_Name = Equip_Name;
            if (rbBatchReport.Checked == true)
            {
             
                reportType = "AlarmReport";
                ReportView vr3 = new ReportView();
                if (cmbReactorName.Text == "E4-ANFD-63")
                {
                    reportType = "DryerAlarmReportE4-ANFD-63";
                    vr3.GenerateAlarmReport(BatchStartDateTime, BatchEndDateTime, timeInterval, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reportType, customta, customva, customrpma, tacceptance, stepnum);
                }
                else if (cmbReactorName.Text == "E4-ANFD-375" || cmbReactorName.Text == "E4-ANFD-371")
                {
                    reportType = "DryerAlarmReportE4-ANFD-375";
                    vr3.GenerateAlarmReport(BatchStartDateTime, BatchEndDateTime, timeInterval, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reportType, customta, customva, customrpma, tacceptance, stepnum);
                }
                else
                {
                    vr3.GenerateAlarmReport(BatchStartDateTime, BatchEndDateTime, timeInterval, ReactorName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reportType, customta, customva, customrpma, tacceptance, stepnum);

                }
            }
            if (rbDryerReport.Checked == true)
            {
                ReportView vr1 = new ReportView();
                reportType = "DryerAlarmReport"; 

                if (Dryer_cmb.Text == "E4-JTM86")
                {
                    connetionString = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                    viewName = "v_API4_JTM86";
                    Equip_Name = "JTM86";
                }
                else if (Dryer_cmb.Text == "E4-JTM87")
                {
                    connetionString = @"DSN=ALL_JTM_MM;Uid=sa;pwd=Cipla@123";
                    viewName = "v_API4_JTM87";
                    Equip_Name = "JTM87";
                }
                vr1.GenerateAlarmReport(BatchStartDateTime, BatchEndDateTime, timeInterval, DryerName, viewName, BatchNo, connetionString, StepNumber, AE_Equip_Name, UserID, reportType, customta, customva, customrpma, tacceptance, stepnum);
            }
            if (RHradioButton.Checked == true)
            {

                if (RHTemp == "DRY AREA TEMP RH")
                {
                    connetionString = @"DSN=DRY_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                    viewName = "vFloatTable";
                    reportType = "AlarmReportRH_Temp";
                }
                if (RHTemp == "WET AREA TEMP RH")
                {
                    connetionString = @"DSN=WET_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                    viewName = "vFloatTable";
                    reportType = "AlarmReport_WET_TEMP_RH";
                }
                if (RHTemp == "COLD ROOM 57")
                {
                    connetionString = @"DSN=WET_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                    viewName = "vFloatTable";
                    reportType = "Alarm_COLD_ROOM_57";
                }
                if (RHTemp == "COLD ROOM 92")
                {
                    connetionString = @"DSN=DRY_AREA_TEMP_RH;Uid=sa;pwd=Cipla@123";
                    viewName = "vFloatTable";
                    reportType = "Alarm_COLD_ROOM_92";
                }

                string frmdt = dtpFromDate.Text;
                string tdt = dtpToDate.Text;
                ConvertDateTimeForReportFormat(frmdt, tdt);
                

                ReportView vr = new ReportView();
                vr.GenerateAlarmReport(fromDateTime, ToDateTime,  timeInterval, RHTemp, viewName, BatchNo,  connetionString,  StepNumber, AE_Equip_Name,  UserID, reportType,  customta,  customva,  customrpma,  tacceptance,  stepnum);
            }
        }

        private void ConvertDateTimeForReportFormat(string frmdt, string tdt)
        {


            DateTime dtfromdt = DateTime.ParseExact(frmdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            DateTime dttodt = DateTime.ParseExact(tdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            fromDateTime = dtfromdt.ToString("yyyy-MM-dd HH:mm:ss");
            ToDateTime = dttodt.ToString("yyyy-MM-dd HH:mm:ss");
            frmdt = dtpFromDate.Text;
            tdt = dtpToDate.Text;

            dtfromdt = DateTime.ParseExact(frmdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);
            dttodt = DateTime.ParseExact(tdt, "dd-MM-yyyy HH:mm:ss", CultureInfo.InvariantCulture);

            fromDateTime = dtfromdt.ToString("yyyy-MM-dd HH:mm:ss");
            ToDateTime = dttodt.ToString("yyyy-MM-dd HH:mm:ss");
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void cmbReactorName_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void lblRelativeHumidityAccp_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }
    }
}

