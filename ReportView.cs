using Microsoft.Reporting.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cipla_Bangalore_API4_091
{
    public partial class ReportView : Form
    {
        private readonly BAL.BAL BAL;
        string RDLCNAME;
        public ReportView()
        {
            InitializeComponent();
        }
        private void ReportView_Load(object sender, EventArgs e)
        {

            this.reportViewer1.RefreshReport();
            reportViewer1.ShowExportButton = true;

            foreach (RenderingExtension extension in reportViewer1.LocalReport.ListRenderingExtensions())
            {
                if (extension.Name == "EXCELOPENXML" || extension.Name == "WORD" || extension.Name == "EXCEL" || extension.Name == "WORDOPENXML")
                {
                    FieldInfo fi = extension.GetType().GetField("m_isVisible", BindingFlags.Instance | BindingFlags.NonPublic);
                    fi.SetValue(extension, false);
                }

            }

        }


        #region GroupPath
        public string GetGroupPath(string reactorName, string reporttype, string Grouppath)
        {

        
                switch (reactorName)
                {
                    case "E4_ANFD371_SOLVANT":
                        Grouppath = "ANFD371SOL_E";
                        break;
                    case "E4_ANFD375_SOLVANT":
                        Grouppath = "ANFD375SOL_E";
                        break;
                    case "E4_ANFD63_SOLVANT":
                        Grouppath = "ANFD63SOL_E";
                        break;
                    case "E4_CF65B_SOLVANT":
                        Grouppath = "CF65BSOL_E";
                        break;
                    case "E4_GLR01_SOLVANT":
                        Grouppath = "GLR01SOL_E";
                        break;
                    case "E4_GLR04_SOLVANT":
                        Grouppath = "GLR04SOL_E";
                        break;
                    case "E4_GLR06_SOLVANT":
                        Grouppath = "GLR06SOL_E";
                        break;
                    case "E4_GLR12_SOLVANT":
                        Grouppath = "GLR12SOL_E";
                        break;
                    case "E4_GLR13_SOLVANT":
                        Grouppath = "GLR13SOL_E";
                        break;
                    case "E4_GLR17_SOLVANT":
                        Grouppath = "GLR17SOL_E";
                        break;
                    case "E4_GLR23_SOLVANT":
                        Grouppath = "GLR23SOL_E";
                        break;
                    case "E4_GLR28_SOLVANT":
                        Grouppath = "GLR28SOL_E";
                        break;
                    case "E4_GLR32_SOLVANT":
                        Grouppath = "GLR32SOL_E";
                        break;
                    case "E4_GLR33_SOLVANT":
                        Grouppath = "GLR33SOL_E";
                        break;
                    case "E4_GLR348_SOLVANT":
                        Grouppath = "GLR348SOL_E";
                        break;
                    case "E4_GLR36_SOLVANT":
                        Grouppath = "GLR36SOL_E";
                        break;
                    case "E4_GLR45_SOLVANT":
                        Grouppath = "GLR45SOL_E";
                        break;
                    case "E4_SSR07_SOLVANT":
                        Grouppath = "SSR07SOL_E";
                        break;
                    case "E4_SSR09_SOLVANT":
                        Grouppath = "SSR09SOL_E";
                        break;
                    case "E4_SSR11_SOLVANT":
                        Grouppath = "SSR11SOL_E";
                        break;
                    case "E4_SSR19_SOLVANT":
                        Grouppath = "SSR19SOL_E";
                        break;
                    case "E4_SSR24_SOLVANT":
                        Grouppath = "SSR24SOL_E";
                        break;
                    case "E4_SSR29_SOLVANT":
                        Grouppath = "SSR29SOL_E";
                        break;
                    case "E4_SSR368_SOLVANT":
                        Grouppath = "SSR368SOL_E";
                        break;
                    case "E4_SSR37_SOLVANT":
                        Grouppath = "SSR37SOL_E";
                        break;
                    case "E4_SSR43_SOLVANT":
                        Grouppath = "SSR43SOL_E";
                        break;
                    case "E4_SSR44_SOLVANT":
                        Grouppath = "SSR44SOL_E";
                        break;
                    case "E4_SSV03_SOLVANT":
                        Grouppath = "SSV03SOL_E";
                        break;
                    case "E4_SSV251_SOLVANT":
                        Grouppath = "SSV251SOL_E";
                        break;
                    case "E4_SSV347_SOLVANT":
                        Grouppath = "SSV347SOL_E";
                        break;
                    case "E4_SSV373_SOLVANT":
                        Grouppath = "SSV373SOL_E";
                        break;
                    case "E4_SSV46_SOLVANT":
                        Grouppath = "SSV46SOL_E";
                        break;
                    case "E4-RCVD-61":
                        Grouppath = "RCVD61_E";
                        break;
                    case "E4-RCVD-69":
                        Grouppath = "RCVD69_E";
                        break;
                    case "E4-RCVD-76":
                        Grouppath = "RCVD76_E";
                        break;
                    case "E4-RCVD-81":
                        Grouppath = "RCVD81_E";
                        break;
                    case "E4-RCVD-370":
                        Grouppath = "RCVD370_E";
                        break;
                default:
                    // Handle unknown SolventName here
                    Grouppath = string.Empty;
                    break;
                }
            
            

            if (reporttype == "SolventEventReport")
            {
                if (reactorName == "E4-ANFD-63")
                {
                    Grouppath = "ANFD63_E";
                }
                else if (reactorName == "E4-ANFD-375")
                {
                    Grouppath = "ANFD375_E";

                }
            }
            if (reporttype == "SolventCustomEventReport")
            {
                if (reactorName == "E4_CF65B_SOLVANT")
                {
                    Grouppath = "CF65BSOL_E";
                }

                if (reactorName == "E4_SSV03_SOLVANT")
                {
                    Grouppath = "SSV03SOL_E";

                }
                if (reactorName == "E4_SSV251_SOLVANT")
                {
                    Grouppath = "SSV251SOL_E";
                }
                if (reactorName == "E4_SSV347_SOLVANT")
                {
                    Grouppath = "SSV347SOL_E";

                }
                if (reactorName == "E4_SSV373_SOLVANT")
                {
                    Grouppath = "SSV373SOL_E";
                }
                if (reactorName == "E4_SSV46_SOLVANT")
                {
                    Grouppath = "SSV46SOL_E";
                }
            }

            return Grouppath;
        }
        public string GetAlarmGroupPath(string reactorName, string reporttype, string Grouppath)
        {
            var groupPaths = new Dictionary<string, string>
                {
                    { "E4-HWS-180", "E4_HWS_180" },
                    { "E4-HWS-181", "E4_HWS_181" },
                    { "E4-HWS-122", "E4_HWS_122" },
                    { "E4-ANFD-63", "ANFD_63_A" },
                    { "E4-ANFD-375", "ANFD_375_A" },
                    { "E4-ANFD-371", "ANFD_371_A" },
                    { "E4-RCVD-350", "RCVD_350_A" },
                    { "E4-RCVD-69", "RCVD_69_A" },
                    { "E4-RCVD-370", "RCVD_370_A" },
                    { "E4-RCVD-61", "RCVD_61_A" },
                    { "E4-RCVD-76", "RCVD_76_A" },
                    { "E4-JTM86", "JTM86_A" },
                    { "E4-JTM87", "JTM87_A" },
                    { "E4-GLR-36", "GLR36_A" },
                    { "E4-SSR-37", "SSR_37_A" },
                    { "E4-GLR-32", "GLR32_A" },
                    { "E4-GLR-23", "GLR23_A" },
                    { "E4-GLR-33", "GLR33_A" },
                    { "E4-SSR-24", "SSR24_A" },
                    { "E4-SSR-368", "SSR_368_A" },
                    { "E4-SSR-43", "SSR_43_A" },
                    { "JTM-86", "JTM86_A" },
                    { "JTM-87", "JTM87_A" },
                    { "E4-GLR-01", "GLR01_A" },
                    { "E4-GLR-04", "GLR04_A" },
                    { "E4-GLR-06", "GLR06_A" },
                    { "E4-GLR-12", "GLR12_A" },
                    { "E4-GLR-13", "GLR13_A" },
                    { "E4-GLR-17", "GLR17_A" },
                    { "E4-GLR-28", "GLR_28_A" },
                    { "E4-GLR-310", "GLR310_A" },
                    { "E4-GLR-348", "GLR348_A" },
                    { "E4-GLR-45", "GLR45_A" },
                    { "E4-SSR-07", "SSR07_A" },
                    { "E4-SSR-09", "SSR09_A" },
                    { "E4-SSR-11", "SSR11_A" },
                    { "E4-SSR-19", "SSR19_A" },
                    { "E4-SSR-29", "SSR29_A" },
                    { "E4-SSR-43", "SSR43_A" },
                    { "E4-SSR-44", "SSR44_A" }
                };

            Grouppath = groupPaths.ContainsKey(reactorName) ? groupPaths[reactorName] : "default_path";
            return Grouppath;
        }
        #endregion


        #region  BatchReport
        public void GenerateBatchReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string timeInterval, string connetionString, string ReportType, string StepNumber, string UserID, string hr, string customta, string customva, string customrpma, string tacceptance, String stepnum, String venturiAcceptance1, String venturiAcceptance2, String grindingAcceptance1, String grindingAcceptance2)
        {
            DataTable dt1 = new DataTable();
            string RDLCName;
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
              
                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetBatchData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, timeInterval);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsReactor(dt1, StepNumber, customta, customva, customrpma, timeInterval, tacceptance, stepnum, venturiAcceptance1, venturiAcceptance2, grindingAcceptance1, grindingAcceptance2);
                        }
                        if(reactorName== "E4-GLR-23" || reactorName== "E4-GLR-32" || reactorName== "E4-GLR-33" || reactorName== "E4-SSR-37")
                        {
                             RDLCName = "E4-GLR-23";
                        }
                        else if (reactorName == "E4-SSR24" || reactorName == "E4-SSR368")
                        {
                             RDLCName = "E4-SSR24";
                        }
                        else if (reactorName == "E4-RCVD-350" || reactorName == "E4-RCVD-370" || reactorName == "E4-RCVD-61" || reactorName == "E4-RCVD-69" || reactorName == "E4-RCVD-76")
                        {
                            RDLCName = "E4-RCVD";
                        }
                        else {  RDLCName = reactorName; }
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        public void GenerateDryerReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string ReportType, string StepNumber, string UserID, string hr, string customta, string customva, string customrpma, string tacceptance, string stepnum, string TimeInterval1, string Venturi_Acceptance1, string Venturi_Acceptance2, string Grinding_Acceptance1, string Grinding_Acceptance2)
        {
            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        // Construct the SQL query string dynamically
                        OdbcCommand query = new OdbcCommand(@"
                            WITH added_row_number AS (
                                SELECT *,
                                       ROW_NUMBER() OVER (PARTITION BY CONVERT(Char(16), [DateAndTime], 20) ORDER BY [DateAndTime] ASC) AS row_number
                                FROM " + viewname + @"
                                WHERE DateAndTime BETWEEN ? AND ? AND BatchNo = ?
                            )
                            SELECT *
                            FROM added_row_number
                            WHERE row_number = 1
                                AND DATEPART(MINUTE, DATEADD(MINUTE," + TimeInterval1 + ", [DateAndTime])) % " + TimeInterval1 + " = 0 ORDER BY DateAndTime ASC;", Conn);
                        query.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
                        query.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
                        query.Parameters.AddWithValue("?", batchNo);  // Assuming batchNo is a valid string parameter

                        // Execute the query
                        OdbcDataAdapter adp = new OdbcDataAdapter(query);
                        query.CommandTimeout = 60000;
                        adp.Fill(dt1);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumns(dt1, customta, customva, TimeInterval1, tacceptance, stepnum, Venturi_Acceptance1, Venturi_Acceptance2, Grinding_Acceptance1, Grinding_Acceptance2);
                        }
                        string RDLCNAME;
                        if (reactorName == "E4-JTM86") { RDLCNAME = reactorName; }
                        else if (reactorName == "E4-JTM87") { RDLCNAME = reactorName; }
                        else { RDLCNAME = "DryerBatchReport"; }

                        PrintReport(dt1, reactorName, RDLCNAME);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        public void GenerateCustomBatchReport(string batchStartDateTime, string batchEndDateTime, string batchno, string reactorName, string viewname, string batchNo, string timeInterval, string connetionString, string ReportType, string StepNumber, string UserID, string hr, string custombatchnumber, string custombmrpagenumber, string custombmrsection, string customproductname, string customTA, string customVA, string custoMrpmA, string tacceptance, string stepnum, string Venturi_Acceptance1, string Venturi_Acceptance2, string Grinding_Acceptance1, string Grinding_Acceptance2)
        {
            DataTable dt1 = new DataTable();

            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
     
               

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetCustomData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, timeInterval);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsReactor(dt1, StepNumber, customTA, customVA, custoMrpmA, timeInterval.ToString(), tacceptance, stepnum, Venturi_Acceptance1, Venturi_Acceptance2, Grinding_Acceptance1, Grinding_Acceptance2);
                            dt1 = AddAdditionalColumnsCustomReactor(dt1, custombatchnumber, customproductname, custombmrpagenumber, custombmrsection, batchno, batchStartDateTime, batchEndDateTime, Venturi_Acceptance1, Venturi_Acceptance2, Grinding_Acceptance1, Grinding_Acceptance2);
                        }
                        else if (reactorName == "E4-RCVD-350" || reactorName == "E4-RCVD-370" || reactorName == "E4-RCVD-61" || reactorName == "E4-RCVD-69" || reactorName == "E4-RCVD-76")
                        {
                            RDLCNAME = "E4-RCVD";
                        }
                        else { RDLCNAME = reactorName; }
                        PrintReport(dt1, reactorName, RDLCNAME);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }

        public void GenerateEventBatchReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval, string customta, string customva, string customrpma, string tacceptance, string stepnum)
        {


            DataTable dt1 = new DataTable();
            DataTable dt = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                string GroupPath = "";

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();
                        GroupPath = GetGroupPath(reactorName, reporttype, GroupPath);
                        if (reactorName == "E4-ANFD-63" || reactorName == "E4-ANFD-375")
                        {
                            dt = GetEventBatchData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval, GroupPath, reactorName);

                        }
                        else if (reactorName == "E4-ANFD-371" || reactorName == "OGB88" || reactorName == "OGB73")
                        {

                        }
                        else
                        {
                            dt = GetEventBatchData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval, GroupPath, reactorName);
                        }
                                     

                        if (reactorName == "E6-DT-62-SOLVENT" || reactorName == "E6-DT-63-SOLVENT" || reactorName == "E6-DT-64-SOLVENT" || reactorName == "E6-DT-65-SOLVENT" || reactorName == "E6-DT-66-SOLVENT" || reactorName == "E6-CHBT-76-SOLVENT" || reactorName == "E6-SSR-03-SOLVENT" || reactorName == "E6-GLR-08-SOLVENT" || reactorName == "E6-GLR-09-SOLVENT" || reactorName == "E6-GLR-10-SOLVENT" || reactorName == "E6-SSR-11-SOLVENT" || reactorName == "E6-SSR-99-SOLVENT")
                        {
                            DataColumn dcSolvent = new DataColumn("SolventName1", typeof(String));
                            dt1.Columns.Add(dcSolvent);
                            if (Convert.ToInt32(dt1.Rows[0]["SolventName"]) == 25)
                            {
                                dt1.Rows[0]["SolventName1"] = "MDC";
                            }
                            if (Convert.ToInt32(dt1.Rows[0]["SolventName"]) == 26)
                            {
                                dt1.Rows[0]["SolventName1"] = "IPA";
                            }
                            if (Convert.ToInt32(dt1.Rows[0]["SolventName"]) == 27)
                            {
                                dt1.Rows[0]["SolventName1"] = "Acetone";
                            }
                            if (Convert.ToInt32(dt1.Rows[0]["SolventName"]) == 28)
                            {
                                dt1.Rows[0]["SolventName1"] = "Methanol";
                            }
                            if (Convert.ToInt32(dt1.Rows[0]["SolventName"]) == 29)
                            {
                                dt1.Rows[0]["SolventName1"] = "DIPE";
                            }
                            if (Convert.ToInt32(dt1.Rows[0]["SolventName"]) == 30)
                            {
                                dt1.Rows[0]["SolventName1"] = "EA";
                            }
                        }
                        if (dt.Rows.Count > 0)
                        {
                            dt = AddDryerAlarmColoumns(dt, StepNumber, customva, customta, customrpma, tacceptance, stepnum, batchStartDateTime, batchEndDateTime);
                        }
                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }

        #endregion BatchReport

        #region GetData

        private DataTable GetAuditData(OdbcConnection connection, string areaName, string batchStartDateTime, string batchEndDateTime, string reporttype, string audittable)
        {
            DataTable dataTable = new DataTable();

            OdbcCommand query = new OdbcCommand("select DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) AS TimeStmp,UserID,Severity,Location,[MessageText] from " + audittable + " " +
                                             "where DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) >= '" + batchStartDateTime + "' " +
                                             " and DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) <= '" + batchEndDateTime + "' "
                                      + " and Location='" + areaName + "'and UserID not Like '%NT AUTHORITY\\SYSTEM%' and UserID not Like '%NT AUTHORITY\\LOCAL SERVICE%' "
                                        + " and UserID not Like '%CIPLA\\INPTG01A3W-006$%' and UserID not Like '%CIPLA\\INPTGSRSDS-01$%' "
                                        + " and UserID not Like '%NT AUTHORITY\\LOCAL SERVICE%' and UserID not Like '%LOCAL SERVICE%' and Severity!=1 and Severity!=2 "
                                        + " and UserID not Like '%CIPLA%'"
                                        + "  and UserID not Like '%FactoryTalk Service%' "
                                        + "   and UserID not like '%INBLRSRENG-09\\ADMINISTRATOR%' and UserID not like '%Administrator%' order by TimeStmp asc", connection);
            OdbcDataAdapter adp = new OdbcDataAdapter(query);
            query.CommandTimeout = 120000;
            adp.Fill(dataTable);


            return dataTable;
        }
        private DataTable GetAuditDataINBLRENGW(OdbcConnection connection, string areaName, string batchStartDateTime, string batchEndDateTime, string reporttype, string audittable)
        {
            DataTable dataTable = new DataTable();

            OdbcCommand query = new OdbcCommand("select DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) AS TimeStmp,UserID,Severity,Location,[MessageText] from test " +
                                                "where DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) >= '" + batchStartDateTime + "' " +
                                                " and DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) <= '" + batchEndDateTime + "' "
                                         + " and Location='" + areaName + "'and UserID not Like '%NT AUTHORITY\\SYSTEM%' and UserID not Like '%NT AUTHORITY\\LOCAL SERVICE%' "
                                           + " and UserID not Like '%CIPLA\\INPTG01A3W-006$%' and UserID not Like '%CIPLA\\INPTGSRSDS-01$%' "
                                           + " and UserID not Like '%NT AUTHORITY\\LOCAL SERVICE%' and UserID not Like '%LOCAL SERVICE%' and Severity!=1 and Severity!=2 "
                                           + " and UserID not Like '%CIPLA%'  "
                                           + "  and UserID not Like '%FactoryTalk Service%' "
                                           + " and UserID not like '%INBLRSRENG-09\\ADMINISTRATOR%' and UserID not like '%Administrator%' order by TimeStmp asc", connection);
            OdbcDataAdapter adp = new OdbcDataAdapter(query);
            query.CommandTimeout = 120000;
            adp.Fill(dataTable);


            return dataTable;
        }
        private DataTable GetBatchData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string timeInterval)
        {
            DataTable dataTable = new DataTable();

            OdbcCommand cmd = new OdbcCommand(@"
                WITH added_row_number AS (
                    SELECT *,
                           ROW_NUMBER() OVER (PARTITION BY CONVERT(Char(16), [DateAndTime], 20) ORDER BY [DateAndTime] ASC) AS row_number
                    FROM " + viewname + @"
                    WHERE DateAndTime BETWEEN ? AND ? 
                )
                SELECT *
                FROM added_row_number
                WHERE row_number = 1
                    AND DATEPART(MINUTE, DATEADD(MINUTE," + timeInterval + ", [DateAndTime])) % " + timeInterval + " = 0 ORDER BY DateAndTime ASC;", connection);

                cmd.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
                cmd.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
                OdbcDataAdapter adp = new OdbcDataAdapter(cmd);
                cmd.CommandTimeout = 120000;
                adp.Fill(dataTable);
            

            return dataTable;
        }
        private DataTable GetCustomData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string timeInterval)
        {
            DataTable dataTable = new DataTable();

            OdbcCommand query = new OdbcCommand(@"
                WITH added_row_number AS (
                    SELECT *,
                           ROW_NUMBER() OVER (PARTITION BY CONVERT(Char(16), [DateAndTime], 20) ORDER BY [DateAndTime] ASC) AS row_number
                    FROM " + viewname + @"
                    WHERE DateAndTime BETWEEN ? AND ? 
                )
                SELECT *
                FROM added_row_number
                WHERE row_number = 1
                    AND DATEPART(MINUTE, DATEADD(MINUTE," + timeInterval + ", [DateAndTime])) % " + timeInterval + " = 0 ORDER BY DateAndTime ASC;", connection);

            OdbcDataAdapter adp = new OdbcDataAdapter(query);
                query.CommandTimeout = 120000;

                adp.Fill(dataTable);
            

            return dataTable;
        }
        private DataTable GetSolventData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string reporttype)
        {
            DataTable dataTable = new DataTable();
            if (reporttype == "AllSolventReport")
            {
                OdbcCommand query = new OdbcCommand("WITH NumberedRows AS(SELECT[DateAndTime],[SET_QTY], LAG([SET_QTY]) OVER(ORDER BY[DateAndTime] ASC) AS PrevSetQty FROM  " + viewname + " WHERE [DateAndTime] BETWEEN '" + batchStartDateTime + "' AND '" + batchEndDateTime + "')SELECT MinSolventDateTime, MaxSolventDateTime, DISPENSE_QTY, SolventName, EquipmentNo, TotalSetQty FROM(SELECT MIN([DateAndTime]) AS MinSolventDateTime, MAX([DateAndTime]) AS MaxSolventDateTime, DISPENSE_QTY, SolventName, EquipmentNo FROM " + viewname + " WHERE[DateAndTime] BETWEEN '" + batchStartDateTime + "' and '" + batchEndDateTime + "' AND EquipmentNo != ''GROUP BY DISPENSE_QTY, SolventName, EquipmentNo) AS SecondQueryResults JOIN(SELECT SUM([SET_QTY]) AS TotalSetQty FROM NumberedRows WHERE[SET_QTY] != PrevSetQty OR PrevSetQty IS NULL) AS FirstQueryResults ON 1 = 1", connection);
                OdbcDataAdapter adp = new OdbcDataAdapter(query);
                query.CommandTimeout = 120000;
                adp.Fill(dataTable);
            }
            if (reporttype == "SolventReport")
            {
                OdbcCommand query = new OdbcCommand("select min(dateandtime) as MinSolventDateTime,MAX(DateAndTime)as MaxSolventDateTime ,SET_QTY,DISPENSE_QTY,SolventName,BatchNo,ProductName,EquipmentNo from " + viewname + " WHERE DateAndTime Between '" + batchStartDateTime + "' and '" + batchEndDateTime + "' and EquipmentNo != '' group by SET_QTY,DISPENSE_QTY,SolventName,BatchNo,ProductName,EquipmentNo order by Min(DateAndTime) asc", connection);
                OdbcDataAdapter adp = new OdbcDataAdapter(query);
                query.CommandTimeout = 120000;
                adp.Fill(dataTable);
            }



            return dataTable;
        }
        private DataTable GetTrendData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string reporttype, string timeInterval)
        {
            DataTable dataTable = new DataTable();

            if (reporttype == "SolventCustomTrendReport")
            {
                OdbcCommand query = new OdbcCommand(@"
                WITH added_row_number AS (
                    SELECT *,
                           ROW_NUMBER() OVER (PARTITION BY CONVERT(Char(16), [DateAndTime], 20) ORDER BY [DateAndTime] ASC) AS row_number
                    FROM " + viewname + @"
                    WHERE DateAndTime BETWEEN ? AND ? 
                )
                SELECT *
                FROM added_row_number
                WHERE row_number = 1
                    AND DATEPART(MINUTE, DATEADD(MINUTE," + timeInterval + ", [DateAndTime])) % " + timeInterval + " = 0 ORDER BY DateAndTime ASC;", connection);

                // Adding parameters

                query.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
                query.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
                OdbcDataAdapter adp = new OdbcDataAdapter(query);
                query.CommandTimeout = 120000;
                adp.Fill(dataTable);
            }
            else
            {
                OdbcCommand query = new OdbcCommand(@"
                    WITH added_row_number AS (
                        SELECT *,
                               ROW_NUMBER() OVER (PARTITION BY CONVERT(Char(16), [DateAndTime], 20) ORDER BY [DateAndTime] ASC) AS row_number
                        FROM " + viewname + @"
                        WHERE DateAndTime BETWEEN ? AND ? AND BatchNo = ?
                    )
                    SELECT *
                    FROM added_row_number
                    WHERE row_number = 1
                        AND DATEPART(MINUTE, DATEADD(MINUTE," + timeInterval + ", [DateAndTime])) % " + timeInterval + " = 0 ORDER BY DateAndTime ASC;", connection);

                // Adding parameters

                query.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
                query.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
                query.Parameters.AddWithValue("?", batchNo);  // Assuming batchNo is a valid string parameter


                OdbcDataAdapter adp = new OdbcDataAdapter(query);
                query.CommandTimeout = 120000;
                adp.Fill(dataTable);
            }

            return dataTable;
        }
        private DataTable GetEventData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string reporttype, string timeInterval, string GroupPath)
        {
            DataTable dataTable = new DataTable();
            OdbcCommand query = new OdbcCommand(@"
                            WITH added_row_number AS (
                                SELECT *,
                                       ROW_NUMBER() OVER (PARTITION BY CONVERT(Char(16), [DateAndTime], 20) ORDER BY [DateAndTime] ASC) AS row_number
                                FROM " + viewname + @"
                                WHERE DateAndTime BETWEEN ? AND ? AND BatchNo = ?
                            )
                            SELECT *
                            FROM added_row_number
                            WHERE row_number = 1
                                AND DATEPART(MINUTE, DATEADD(MINUTE," + timeInterval + ", [DateAndTime])) % " + timeInterval + " = 0 ORDER BY DateAndTime ASC;", connection);
            query.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
            query.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
            query.Parameters.AddWithValue("?", batchNo);  // Assuming batchNo is a valid string parameter


            OdbcDataAdapter adp = new OdbcDataAdapter(query);
            query.CommandTimeout = 120000;
            adp.Fill(dataTable);

            return dataTable;
        }

        private DataTable GetEventBatchData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string reporttype, string timeInterval, string GroupPath, string reactorName)
        {
            DataTable dataTable = new DataTable();
            string query;
            if (reactorName == "E6-DT-62-SOLVENT" || reactorName == "E6-DT-63-SOLVENT" || reactorName == "E6-DT-64-SOLVENT" || reactorName == "E6-DT-65-SOLVENT" || reactorName == "E6-DT-66-SOLVENT" || reactorName == "E6-CHBT-76-SOLVENT" || reactorName == "E6-SSR-03-SOLVENT" || reactorName == "E6-GLR-08-SOLVENT" || reactorName == "E6-GLR-09-SOLVENT" || reactorName == "E6-GLR-10-SOLVENT" || reactorName == "E6-SSR-11-SOLVENT" || reactorName == "E6-SSR-99-SOLVENT")
            {
                query = "SELECT Min(DateAndTime) as DateAndTime, EventMessage,EquipmentNo,BatchNo,ProductName,SolventName FROM " + viewname + " WHERE BatchNo = '" + batchNo + "' and DateAndTime >= '" + batchStartDateTime + "' and DateAndTime <= '" + batchEndDateTime + "' and EventMEssage not like '%IDLE%' Group by EventMessage,EquipmentNo,BatchNo,ProductName,SolventName order by Min(DateAndTime) asc";
            }
            else
            {
                query = "Select CONVERT(DATETIME, CONVERT(VARCHAR(20), EventTimeStamp, 120)) AS EventTimeStamp, Message ,Tag2Value ,Tag3Value,Active, Tag4Value,GroupPath  from UDV_AllEvents where EventTimeStamp >= '" + batchStartDateTime + "' and EventTimeStamp <= '" + batchEndDateTime + "'  and Tag3Value ='" + batchNo + "'and GroupPath= '" + GroupPath + "'and Active=1 order by EventTimeStamp asc";
            }
            OdbcCommand cmd = new OdbcCommand(query, connection);
            cmd.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
            cmd.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
            cmd.Parameters.AddWithValue("?", batchNo);  // Assuming batchNo is a valid string parameter


            OdbcDataAdapter adp = new OdbcDataAdapter(cmd);
            cmd.CommandTimeout = 120000;
            adp.Fill(dataTable);

            return dataTable;
        }

        private DataTable GetEventCustomData(OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string reporttype, string timeInterval, string GroupPath, string reactorName)
        {
            DataTable dataTable = new DataTable();
            string query;
           
            query = "Select CONVERT(DATETIME, CONVERT(VARCHAR(20), EventTimeStamp, 120)) AS EventTimeStamp, Message ,Tag2Value ,Tag3Value,Active, Tag4Value,GroupPath  from UDV_AllEvents where EventTimeStamp >= '" + batchStartDateTime + "' and EventTimeStamp <= '" + batchEndDateTime + "'  and Tag3Value ='" + batchNo + "'and GroupPath= '" + GroupPath + "'and Active=1 order by EventTimeStamp asc";
            OdbcCommand cmd = new OdbcCommand(query, connection);
            cmd.Parameters.AddWithValue("?", batchStartDateTime);  // Assuming FRMDTTM is a valid DateTime parameter
            cmd.Parameters.AddWithValue("?", batchEndDateTime);  // Assuming TODTTM is a valid DateTime parameter
            cmd.Parameters.AddWithValue("?", batchNo);  // Assuming batchNo is a valid string parameter


            OdbcDataAdapter adp = new OdbcDataAdapter(cmd);
            cmd.CommandTimeout = 120000;
            adp.Fill(dataTable);

            return dataTable;
        }

        private DataTable GetAlarmData(DataTable dataTable, string reactorName, OdbcConnection connection, string viewname, string batchStartDateTime, string batchEndDateTime, string batchNo, string reporttype, string timeInterval, string Grouppath)
        {
            connection.Open();
            OdbcCommand query = new OdbcCommand();
            if (reactorName == "E4-JTM86" || reactorName == "E4-JTM87")
            {
                query = new OdbcCommand("select  * from UDV_Alarms_Jetmill where(EventTimeStamp >= '" + batchStartDateTime + "' and EventTimeStamp <= '" + batchEndDateTime + "') AND BatchNo='" + batchNo + "'and  GroupPath= '" + Grouppath + "' and Active is not null order by EventTimeStamp asc", connection);
            }
            else
            {
                query = new OdbcCommand("select  * from UDV_Alarms where(EventTimeStamp >= '" + batchStartDateTime + "' and EventTimeStamp <= '" + batchEndDateTime + "') AND BatchNo='" + batchNo + "'and  GroupPath= '" + Grouppath + "' and Active is not null order by EventTimeStamp asc", connection);
            }
            OdbcDataAdapter adp = new OdbcDataAdapter(query);
            query.CommandTimeout = 120000;
            DataTable dt1 = new DataTable();
            adp.Fill(dt1);

            connection.Close();

            return dataTable;
        }

        private DataTable GetEventMessageData(DataTable dt, string connetionString1, string batchStartDateTime, string batchEndDateTime, string batchNo, string GroupPath,string reactorName)
        {
            OdbcConnection Conn1 = new OdbcConnection(connetionString1);
            DataTable dt1 = new DataTable();
            OdbcCommand query = new OdbcCommand("Select CONVERT(DATETIME, CONVERT(VARCHAR(20), EventTimeStamp, 120)) AS EventTimeStamp, Message ,Tag2Value ,Tag3Value,Active, Tag4Value,GroupPath  from UDV_AllEvents where EventTimeStamp >= '" + batchStartDateTime + "' and EventTimeStamp <= '" + batchEndDateTime + "'  and Tag3Value ='" + batchNo + "'and GroupPath= '" + GroupPath + "'and Active=1 order by EventTimeStamp asc", Conn1);

            OdbcDataAdapter adp = new OdbcDataAdapter(query);
            query.CommandTimeout = 120000;
            //DataTable dt1 = new DataTable();
            adp.Fill(dt1);

            if (dt1.Rows.Count > 0)
            {
                if (!dt1.Columns.Contains("batchStartDateTime"))
                    dt1.Columns.Add(new DataColumn("batchStartDateTime", typeof(string)));
                if (!dt1.Columns.Contains("batchEndDateTime"))
                    dt1.Columns.Add(new DataColumn("batchEndDateTime", typeof(string)));

                dt1.Rows[0]["batchStartDateTime"] = batchStartDateTime;
                dt1.Rows[0]["batchEndDateTime"] = batchEndDateTime;


                DataColumn Tag4Value = new DataColumn("Tag4Value", typeof(String));
                dt.Columns.Add(Tag4Value);
                dt.Rows[0]["Tag4Value"] = reactorName;

                DataColumn Tag1Value = new DataColumn("Tag1Value", typeof(String));
                dt.Columns.Add(Tag1Value);
                dt.Rows[0]["Tag1Value"] = dt.Rows[0]["BatchNo"];

                DataColumn Tag2Value = new DataColumn("Tag2Value", typeof(String));
                dt.Columns.Add(Tag2Value);
                dt.Rows[0]["Tag2Value"] = dt.Rows[0]["ProductName"];

                DataColumn Tag3Value = new DataColumn("Tag3Value", typeof(String));
                dt.Columns.Add(Tag3Value);
                dt.Rows[0]["Tag3Value"] = dt.Rows[0]["StepNumber"];

                
            }

            return dt;
        }

        #endregion

        #region Solvent Report
        public void GenerateAllSolventReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string reporttype)
        {
            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;


                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetSolventData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsALLSolvent(dt1, batchStartDateTime, batchEndDateTime);
                        }
                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        public void GenerateSolventReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string reporttype)
        {
            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;


                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetSolventData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsSolvent(dt1, batchStartDateTime, batchEndDateTime);
                        }
                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        public void GenerateSolventTrendReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval)
        {


            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;


                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetTrendData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsSolvent(dt1, batchStartDateTime, batchEndDateTime);
                        }
                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        public void GenerateTrendReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval, string customta, string customva, string customrpma, string tacceptance, string stepnum)
        {


            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;


                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetTrendData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddDryerAlarmColoumns(dt1, StepNumber, customva, customta, customrpma, tacceptance, stepnum, batchStartDateTime, batchEndDateTime);
                        }
                        string RDLCName = GetRDLCNameReactorWise(reactorName);
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }

        public void GenerateSolventCustomTrendReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval)
        {


            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;


                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetTrendData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsSolvent(dt1, batchStartDateTime, batchEndDateTime);
                        }
                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }

        public void SolventEventReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval)
        {


            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                string GroupPath = "";
                string connetionString1 = @"DSN=HMI_Alarm_Event;Uid=sa;pwd=Cipla@123";

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();
                        GroupPath = GetGroupPath(reactorName, reporttype, GroupPath);
                        dt1 = GetEventData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval, GroupPath);

                        dt1 = GetEventMessageData(dt1, connetionString1, batchStartDateTime, batchEndDateTime, batchNo, GroupPath, reactorName);

                        if (dt1.Rows.Count > 0)
                        {
                            if (!dt1.Columns.Contains("batchStartDateTime"))
                                dt1.Columns.Add(new DataColumn("batchStartDateTime", typeof(string)));
                            if (!dt1.Columns.Contains("batchEndDateTime"))
                                dt1.Columns.Add(new DataColumn("batchEndDateTime", typeof(string)));

                            dt1.Rows[0]["batchStartDateTime"] = batchStartDateTime;
                            dt1.Rows[0]["batchEndDateTime"] = batchEndDateTime;
                        }
                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }

        public void SolventCustomEventReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval)
        {


            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                string GroupPath = "";
                string connetionString1 = @"DSN=HMI_Alarm_Event;uid=sa;pwd=Cipla@123";
                using (OdbcConnection Conn = new OdbcConnection(connetionString1))
                {
                    try
                    {
                        Conn.Open();
                        GroupPath = GetGroupPath(reactorName, reporttype, GroupPath);
                        dt1 = GetEventData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval, GroupPath);

                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }

        #endregion

        #region DryerReport

        public void GenerateDryerTrendReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string timeInterval, string connetionString, string reporttype, string UserID, string stepnum)
        {
            DataTable dt1 = new DataTable();
            reportViewer1.ProcessingMode = ProcessingMode.Local;


            using (OdbcConnection Conn = new OdbcConnection(connetionString))
            {
                try
                {
                    Conn.Open();

                    dt1 = GetTrendData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval);

                    if (dt1.Rows.Count > 0)
                    {
                        if (!dt1.Columns.Contains("batchStartDateTime"))
                            dt1.Columns.Add(new DataColumn("batchStartDateTime", typeof(string)));
                        if (!dt1.Columns.Contains("batchEndDateTime"))
                            dt1.Columns.Add(new DataColumn("batchEndDateTime", typeof(string)));
                        if (!dt1.Columns.Contains("stepnum"))
                            dt1.Columns.Add(new DataColumn("stepnum", typeof(string)));


                        dt1.Rows[0]["batchStartDateTime"] = batchStartDateTime;
                        dt1.Rows[0]["batchEndDateTime"] = batchEndDateTime;
                        dt1.Rows[0]["stepnum"] = stepnum;
                    }
                    string RDLCName = reporttype;
                    PrintReport(dt1, reactorName, RDLCName);

                    Conn.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                }
            }

        }
        public void DryerEventReport(string batchStartDateTime, string batchEndDateTime, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Event_Equip_Name, string UserID, string reporttype, string timeInterval)
        {


            DataTable dt1 = new DataTable();
            DataTable dt = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                string GroupPath = "";

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();
                        GroupPath = GetGroupPath(reactorName, reporttype, GroupPath);
                        dt = GetEventData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval, GroupPath);
                        Conn.Close();


                        string connetionString1 = @"DSN=HMI_Alarm_Event;Uid=sa;pwd=Cipla@123";
                        dt1 = GetEventMessageData(dt1, connetionString1, batchStartDateTime, batchEndDateTime, batchNo, GroupPath, reactorName);






                        string RDLCName = reporttype;
                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        public void GenerateAlarmReport(string batchStartDateTime, string batchEndDateTime, string timeInterval, string reactorName, string viewname, string batchNo, string connetionString, string StepNumber, string Alarm_Equip_Name, string UserID, string reporttype, string customta, string customva, string customrpma, string tacceptance, string stepnum)
        {
            string Grouppath = "";
            DataTable dt = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                string connetionString1 = @"DSN=HMI_Alarm_Event;uid=sa;pwd=Cipla@123";
                string RDLCName = reporttype;

                using (OdbcConnection Conn = new OdbcConnection(connetionString1))
                {
                    try
                    {
                        Grouppath = GetAlarmGroupPath(reactorName, reporttype, Grouppath);
                        dt = GetAlarmData(dt, reactorName, Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, reporttype, timeInterval, Grouppath);
                        dt = AddDryerAlarmColoumns(dt, StepNumber, customva, customta, customrpma, tacceptance, stepnum, batchStartDateTime, batchEndDateTime);
                        if (dt.Rows.Count > 0)
                        {
                            if (dt.Rows.Count > 0)
                            {
                                DataColumn dc4 = new DataColumn("Active1", typeof(String));
                                dt.Columns.Add(dc4);

                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    if (((bool)dt.Rows[i]["Active"] == true))
                                    {

                                        string Active = "Active";
                                        dt.Rows[i]["Active1"] = Active;

                                    }
                                    else if ((bool)dt.Rows[i]["Active"] == false)
                                    {
                                        string Inactive = "Inactive";
                                        dt.Rows[i]["Active1"] = Inactive;
                                    }
                                }
                            }


                            if (reactorName == "E4-JTM86" || reactorName == "E4-JTM87")
                            {
                                RDLCName = "DryerAlarmReport-JMM";
                            }
                            else
                            {
                                RDLCName = reporttype;
                            }
                            PrintReport(dt, reactorName, RDLCName);

                        }

                        else
                        {
                            OdbcConnection Conn2 = new OdbcConnection(connetionString);
                            Conn2.Open();
                            DataTable dt2 = new DataTable();
                            string datetime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                            OdbcCommand query12 = new OdbcCommand("SELECT top 1 * FROM " + viewname + " WHERE DateAndTime BETWEEN '" + batchStartDateTime + "' AND '" + batchEndDateTime + "' ORDER BY DateAndTime ASC ", Conn2);
                            OdbcDataAdapter adp12 = new OdbcDataAdapter(query12);
                            query12.CommandTimeout = 60000;
                            adp12.Fill(dt2);

                            RDLCName = "AlarmReportnotGenerated";

                            PrintReport(dt, reactorName, RDLCName);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Alarm Report Error!" + ex.ToString());
                    }
                    finally
                    {
                        Conn.Close();
                    }



                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Alarm Report Error!" + ex.ToString());
            }
        }

        #endregion

        #region TempRHReport

        public void GenerateRHTEMP(string batchStartDateTime, string batchEndDateTime, string tacceptance, string relativeaccpt, string relativehumaccpt, string customta, string reactorName, string viewname, string batchNo, string timeInterval, string connetionString, string ReportType, string StepNumber, string UserIDr, string custoMrpmA, string customTA)
        {
            DataTable dt1 = new DataTable();
            string RDLCName;
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();

                        dt1 = GetTrendData(Conn, viewname, batchStartDateTime, batchEndDateTime, batchNo, ReportType, timeInterval);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsTempRH(dt1, StepNumber, timeInterval, custoMrpmA, customTA, relativeaccpt, relativehumaccpt, batchStartDateTime, tacceptance, batchEndDateTime);
                        }

                        var reportTypeReactors = new List<string>
                {
                    "WET AREA TEMP RH",
                    "COLD ROOM 57",
                    "COLD ROOM 92"
                };

                        var reactorNameReactors = new List<string>
                {
                    "E4-GLR-28",
                    "E4-GLR-310",
                    "E4-HWS-180",
                    "E4-HWS-181",
                    "E4-HWS-122"
                };

                        if (reactorName == "DRY AREA TEMP RH")
                        {
                            RDLCName = ReportType;
                        }
                        else if (reportTypeReactors.Contains(reactorName))
                        {
                            RDLCName = ReportType;
                        }
                        else if (reactorNameReactors.Contains(reactorName))
                        {
                            RDLCName = reactorName;
                        }
                        else
                        {
                            RDLCName = "temp";
                        }

                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }
        }
        #endregion


        #region Print  Report
        public void PrintReport(DataTable dataTable, string reactorName, string RDLCName)
        {
            if (dataTable.Rows.Count > 0)
            {
                string connectionStringUser = @"DSN=INBLRENGW_091_New;uid=sa;pwd=Cipla@123";
                OdbcConnection ConnUser = new OdbcConnection(connectionStringUser);

                //ConnUser.Open();

                //try
                //{
                //    OdbcCommand queryprintby = new OdbcCommand("Select top 1 DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) AS TimeStmp ,UserID,Severity,Location,[MessageText] from INBLRENGW_091 " +
                //        "where UserID not Like '%NT AUTHORITY\\SYSTEM%' and UserID not Like '%NT AUTHORITY\\LOCAL SERVICE%'"
                //        + " and UserID not Like '%CIPLA\\INPTG01A3W-006$%' and UserID not Like '%CIPLA\\INPTGSRSDS-01$%' "
                //        + " and UserID not Like '%NT AUTHORITY\\LOCAL SERVICE%' and UserID not Like '%LOCAL SERVICE%' and Severity!=1 and Severity!=2 "
                //        + " and UserID not Like '%CIPLA%' and MessageText not like '%Executed macro:%'"
                //        + " and UserID not Like '%WORKGROUP%' and UserID not Like '%FactoryTalk Service%' and Location = 'INBLRENGW-091' order by DATEADD(MINUTE, 30, DATEADD(HOUR, 5, TimeStmp)) desc ", ConnUser);
                //    OdbcDataAdapter adp_Audit = new OdbcDataAdapter(queryprintby);
                //    queryprintby.CommandTimeout = 120000;
                //    DataTable dt_Audit = new DataTable();
                //    adp_Audit.Fill(dt_Audit);

                //    if (dt_Audit.Rows.Count > 0)
                //    {
                //        string userName = Convert.ToString(dt_Audit.Rows[0]["UserID"]);
                //        string datetime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");
                //        string PrintBy = userName + " " + datetime;
                //        DataColumn dc1 = new DataColumn("PrintBy", typeof(String));
                //        dataTable.Columns.Add(dc1);
                //        dataTable.Rows[0]["PrintBy"] = PrintBy;
                //    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Error getting printing information: " + ex.ToString());
                //}
                //finally
                //{
                //    ConnUser.Close();
                //}


                reportViewer1.LocalReport.ReportPath = Path.Combine(Application.StartupPath, @"..\..\RDLCFiles", RDLCName + ".rdlc");

                ReportDataSource reportdataSource = new ReportDataSource("ReactorData", dataTable);
                reportViewer1.LocalReport.EnableHyperlinks = true;
                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.DataSources.Add(reportdataSource);
                reportViewer1.LocalReport.Refresh();

                this.Show();
            }
            else
            {
                MessageBox.Show("No Record Found!!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #endregion

        #region AuditReport
        public void GenerateAuditReportINBLRENGW091(string batchStartDateTime, string batchEndDateTime, string areaName, string connetionString1, string productName, string reactorName, string BatchNo, string connetionString, string ViewName, string AE_Equip_Name, string UserID, string audittable, string ReportType)
        {
            string batchStartDateTime1 = batchStartDateTime;
            string batchEndDateTime1 = batchEndDateTime;
            DataTable dt1 = new DataTable();
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();
                        dt1 = GetAuditDataINBLRENGW(Conn, areaName, batchStartDateTime, batchEndDateTime, ReportType, audittable);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsAuditReactor(dt1, BatchNo, reactorName, areaName, batchStartDateTime, batchEndDateTime, productName);
                        }
                        string RDLCName = reactorName;
                        if (dt1.Rows.Count > 0)
                        {
                            if (reactorName == "E4-HWS-180")
                            {
                                RDLCName = "AuditReportHWS180.rdlc";
                            }
                            else if (reactorName == "E4-HWS-181")
                            {
                                RDLCName = "AuditReportHWS181.rdlc";
                            }
                            else if (reactorName == "E4-HWS-122")
                            {
                                RDLCName = "AuditReportHWS122.rdlc";
                            }
                            else if (reactorName == "DRY AREA TEMP RH")
                            {
                                RDLCName = "AuditReportDRYAREA.rdlc";
                            }
                            else if (reactorName == "COLD ROOM 92")
                            {
                                RDLCName = "AuditReportCOLDROOM92.rdlc";
                            }
                            else if (reactorName == "COLD ROOM 57")
                            {
                                RDLCName = "AuditReportCOLDROOM57.rdlc";

                            }
                            else if (reactorName == "WET AREA TEMP RH")
                            {
                                RDLCName = "AuditReportWETAREA.rdlc";
                            }

                            else
                            {
                                RDLCName = ReportType;
                            }

                        }

                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }

        }
        public void GenerateAuditReport(string batchStartDateTime, string batchEndDateTime, string areaName, string connetionString1, string productName, string reactorName, string BatchNo, string connetionString, string ViewName, string AE_Equip_Name, string UserID, string audittable, string ReportType)
        {
            string batchStartDateTime1 = batchStartDateTime;
            string batchEndDateTime1 = batchEndDateTime;
            try
            {
                reportViewer1.ProcessingMode = ProcessingMode.Local;
                DataTable dt1 = new DataTable();

                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    try
                    {
                        Conn.Open();
                        dt1 = GetAuditData(Conn, areaName, batchStartDateTime, batchEndDateTime, ReportType, audittable);

                        if (dt1.Rows.Count > 0)
                        {
                            dt1 = AddAdditionalColumnsAuditReactor(dt1, BatchNo, reactorName, areaName, batchStartDateTime, batchEndDateTime, productName);
                        }
                        string RDLCName = reactorName;
                        if (dt1.Rows.Count > 0)
                        {
                            if (reactorName == "E4-HWS-180")
                            {
                                RDLCName = "AuditReportHWS180.rdlc";
                            }
                            else if (reactorName == "E4-HWS-181")
                            {
                                RDLCName = "AuditReportHWS181.rdlc";
                            }
                            else if (reactorName == "E4-HWS-122")
                            {
                                RDLCName = "AuditReportHWS122.rdlc";
                            }
                            else if (reactorName == "DRY AREA TEMP RH")
                            {
                                RDLCName = "AuditReportDRYAREA.rdlc";
                            }
                            else if (reactorName == "COLD ROOM 92")
                            {
                                RDLCName = "AuditReportCOLDROOM92.rdlc";
                            }
                            else if (reactorName == "COLD ROOM 57")
                            {
                                RDLCName = "AuditReportCOLDROOM57.rdlc";

                            }
                            else if (reactorName == "WET AREA TEMP RH")
                            {
                                RDLCName = "AuditReportWETAREA.rdlc";
                            }

                            else
                            {
                                RDLCName = ReportType;
                            }

                        }

                        PrintReport(dt1, reactorName, RDLCName);

                        Conn.Close();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Generate Batch Report Error!" + ex.ToString());
            }

        }
        #endregion


      

        #region ADD coloumns
        public DataTable AddAdditionalColumnsALLSolvent(DataTable dataTable, string batchStartDateTime, string batchEndDateTime)
        {
            // Add new columns if they do not already exist
            if (!dataTable.Columns.Contains("SolventName1"))
                dataTable.Columns.Add(new DataColumn("SolventName1", typeof(string)));
            if (!dataTable.Columns.Contains("TotalMethanol"))
                dataTable.Columns.Add(new DataColumn("TotalMethanol", typeof(int)));
            if (!dataTable.Columns.Contains("TotalToulene"))
                dataTable.Columns.Add(new DataColumn("TotalToulene", typeof(int)));
            if (!dataTable.Columns.Contains("TotalAcetone"))
                dataTable.Columns.Add(new DataColumn("TotalAcetone", typeof(int)));
            if (!dataTable.Columns.Contains("TotalMDC"))
                dataTable.Columns.Add(new DataColumn("TotalMDC", typeof(int)));
            if (!dataTable.Columns.Contains("TotalEthylAcetate"))
                dataTable.Columns.Add(new DataColumn("TotalEthylAcetate", typeof(int)));
            if (!dataTable.Columns.Contains("batchStartDateTime"))
                dataTable.Columns.Add(new DataColumn("batchStartDateTime", typeof(string)));
            if (!dataTable.Columns.Contains("batchEndDateTime"))
                dataTable.Columns.Add(new DataColumn("batchEndDateTime", typeof(string)));

            int TotalMethanol = 0;
            int TotalToulene = 0;
            int TotalAcetone = 0;
            int TotalMDC = 0;
            int TotalEthylAcetate = 0;

            foreach (DataRow row in dataTable.Rows)
            {
                int solventName;
                int totalSetQty;

                if (int.TryParse(row["SolventName"].ToString(), out solventName) && int.TryParse(row["TotalSetQty"].ToString(), out totalSetQty))
                {
                    switch (solventName)
                    {
                        case 1:
                            row["SolventName1"] = "Methanol";
                            TotalMethanol += totalSetQty;
                            break;
                        case 2:
                            row["SolventName1"] = "Toulene";
                            TotalToulene += totalSetQty;
                            break;
                        case 3:
                            row["SolventName1"] = "Acetone";
                            TotalAcetone += totalSetQty;
                            break;
                        case 4:
                            row["SolventName1"] = "MDC";
                            TotalMDC += totalSetQty;
                            break;
                        case 5:
                            row["SolventName1"] = "EthylAcetate";
                            TotalEthylAcetate += totalSetQty;
                            break;
                        default:
                            row["SolventName1"] = "Unknown";
                            break;
                    }
                }
            }

            // Assign the totals to the first row
            if (dataTable.Rows.Count > 0)
            {
                dataTable.Rows[0]["TotalMethanol"] = TotalMethanol;
                dataTable.Rows[0]["TotalToulene"] = TotalToulene;
                dataTable.Rows[0]["TotalAcetone"] = TotalAcetone;
                dataTable.Rows[0]["TotalMDC"] = TotalMDC;
                dataTable.Rows[0]["TotalEthylAcetate"] = TotalEthylAcetate;

                dataTable.Rows[0]["batchStartDateTime"] = batchStartDateTime;
                dataTable.Rows[0]["batchEndDateTime"] = batchEndDateTime;
            }

            return dataTable;
        }
        public DataTable AddAdditionalColumnsSolvent(DataTable dataTable, string batchStartDateTime, string batchEndDateTime)
        {
            // Add new columns if they do not already exist
            if (!dataTable.Columns.Contains("SolventName1"))
                dataTable.Columns.Add(new DataColumn("SolventName1", typeof(string)));
            if (!dataTable.Columns.Contains("batchStartDateTime"))
                dataTable.Columns.Add(new DataColumn("batchStartDateTime", typeof(string)));
            if (!dataTable.Columns.Contains("batchEndDateTime"))
                dataTable.Columns.Add(new DataColumn("batchEndDateTime", typeof(string)));



            foreach (DataRow row in dataTable.Rows)
            {

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    switch (Convert.ToInt32(dataTable.Rows[i]["SolventName"]))
                    {
                        case 1:
                            row["SolventName1"] = "Methanol";
                            break;
                        case 2:
                            row["SolventName1"] = "Toulene";
                            break;
                        case 3:
                            row["SolventName1"] = "Acetone";

                            break;
                        case 4:
                            row["SolventName1"] = "MDC";
                            break;
                        case 5:
                            row["SolventName1"] = "EthylAcetate";
                            break;
                        default:
                            row["SolventName1"] = "Unknown";
                            break;
                    }
                }

            }

            // Assign the totals to the first row
            if (dataTable.Rows.Count > 0)
            {
                dataTable.Rows[0]["batchStartDateTime"] = batchStartDateTime;
                dataTable.Rows[0]["batchEndDateTime"] = batchEndDateTime;
            }

            return dataTable;
        }

        public DataTable AddAdditionalColumnsTempRH(DataTable dataTable, string StepNumber, string timeInterval, string custoMrpmA, string customTA, string relativeaccpt, string relativehumaccpt, string tacceptance, string batchStartDateTime, string batchEndDateTime)
        {
            // Add new columns if they do not already exist
            if (!dataTable.Columns.Contains("SolventName1"))
                dataTable.Columns.Add(new DataColumn("SolventName1", typeof(string)));
            if (!dataTable.Columns.Contains("batchStartDateTime"))
                dataTable.Columns.Add(new DataColumn("batchStartDateTime", typeof(string)));
            if (!dataTable.Columns.Contains("batchEndDateTime"))
                dataTable.Columns.Add(new DataColumn("batchEndDateTime", typeof(string)));



            foreach (DataRow row in dataTable.Rows)
            {

                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    switch (Convert.ToInt32(dataTable.Rows[i]["SolventName"]))
                    {
                        case 1:
                            row["SolventName1"] = "Methanol";
                            break;
                        case 2:
                            row["SolventName1"] = "Toulene";
                            break;
                        case 3:
                            row["SolventName1"] = "Acetone";

                            break;
                        case 4:
                            row["SolventName1"] = "MDC";
                            break;
                        case 5:
                            row["SolventName1"] = "EthylAcetate";
                            break;
                        default:
                            row["SolventName1"] = "Unknown";
                            break;
                    }
                }

            }

            // Assign the totals to the first row
            if (dataTable.Rows.Count > 0)
            {
                dataTable.Rows[0]["batchStartDateTime"] = batchStartDateTime;
                dataTable.Rows[0]["batchEndDateTime"] = batchEndDateTime;
            }

            return dataTable;
        }
        public DataTable AddAdditionalColumnsReactor(DataTable dataTable, string StepNumber, string customta, string customva, string customrpma, string timeInterval, string tacceptance, string stepnum, string venturiAcceptance1, string venturiAcceptance2, string grindingAcceptance1, string  grindingAcceptance2)
        {
            string datetime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

            DataColumn dc2 = new DataColumn("StepNo", typeof(String));
            dataTable.Columns.Add(dc2);
            dataTable.Rows[0]["StepNo"] = StepNumber;

            DataColumn dc3 = new DataColumn("customva", typeof(String));
            dataTable.Columns.Add(dc3);
            dataTable.Rows[0]["customva"] = customva;

            DataColumn dc4 = new DataColumn("customta", typeof(String));
            dataTable.Columns.Add(dc4);
            dataTable.Rows[0]["customta"] = customta;

            DataColumn dc5 = new DataColumn("customrpma", typeof(String));
            dataTable.Columns.Add(dc5);
            dataTable.Rows[0]["customrpma"] = customrpma;

            DataColumn dc6 = new DataColumn("timeInterval", typeof(String));
            dataTable.Columns.Add(dc6);
            dataTable.Rows[0]["timeInterval"] = timeInterval;

            DataColumn dc7 = new DataColumn("tacceptance", typeof(String));
            dataTable.Columns.Add(dc7);
            dataTable.Rows[0]["tacceptance"] = tacceptance;

            DataColumn dc8 = new DataColumn("stepnum", typeof(String));
            dataTable.Columns.Add(dc8);
            dataTable.Rows[0]["stepnum"] = stepnum;

            DataColumn dcventuriAcceptance1 = new DataColumn("venturiAcceptance1", typeof(String));
            dataTable.Columns.Add(dcventuriAcceptance1);
            dataTable.Rows[0]["stepnum"] = venturiAcceptance1;

            DataColumn dcventuriAcceptance2 = new DataColumn("venturiAcceptance2", typeof(String));
            dataTable.Columns.Add(dcventuriAcceptance2);
            dataTable.Rows[0]["stepnum"] = venturiAcceptance2;

            DataColumn dcgrindingAcceptance1 = new DataColumn("grindingAcceptance1", typeof(String));
            dataTable.Columns.Add(dcgrindingAcceptance1);
            dataTable.Rows[0]["stepnum"] = grindingAcceptance1;

            DataColumn dcgrindingAcceptance2 = new DataColumn("grindingAcceptance2", typeof(String));
            dataTable.Columns.Add(dcgrindingAcceptance2);
            dataTable.Rows[0]["stepnum"] = grindingAcceptance2;



            return dataTable;
        }
        public DataTable AddDryerAlarmColoumns(DataTable dt, string StepNumber, string customva, string customta, string customrpma, string tacceptance, string stepnum, string batchStartDateTime, string batchEndDateTime)
        {
            DataColumn dcStepNo = new DataColumn("stepno", typeof(String));
            dt.Columns.Add(dcStepNo);
            dt.Rows[0]["stepno"] = StepNumber;

            DataColumn dc9 = new DataColumn("customva", typeof(String));
            dt.Columns.Add(dc9);
            dt.Rows[0]["customva"] = customva;

            DataColumn dc10 = new DataColumn("customta", typeof(String));
            dt.Columns.Add(dc10);
            dt.Rows[0]["customta"] = customta;


            DataColumn dc5 = new DataColumn("customrpma", typeof(String));
            dt.Columns.Add(dc5);
            dt.Rows[0]["customrpma"] = customrpma;



            DataColumn dc7 = new DataColumn("tacceptance", typeof(String));
            dt.Columns.Add(dc7);
            dt.Rows[0]["tacceptance"] = tacceptance;

            DataColumn dc8 = new DataColumn("stepnum", typeof(String));
            dt.Columns.Add(dc8);
            dt.Rows[0]["stepnum"] = stepnum;

            DataColumn dc11 = new DataColumn("batchStartDateTime", typeof(String));
            dt.Columns.Add(dc11);
            dt.Rows[0]["batchStartDateTime"] = batchStartDateTime;

            DataColumn dc12 = new DataColumn("batchEndDateTime", typeof(String));
            dt.Columns.Add(dc12);
            dt.Rows[0]["batchEndDateTime"] = batchEndDateTime;

            return dt;
        }

        public DataTable AddAdditionalColumnsAuditReactor(DataTable dataTable, string BatchNo, string reactorName, string areaName, string batchStartDateTime, string batchEndDateTime, string productName)
        {
            DataColumn dcbatchNo = new DataColumn("BatchNo", typeof(String));
            dataTable.Columns.Add(dcbatchNo);
            dataTable.Rows[0]["BatchNo"] = BatchNo;

            DataColumn dcReactor = new DataColumn("ReactorName", typeof(String));
            dataTable.Columns.Add(dcReactor);
            dataTable.Rows[0]["ReactorName"] = reactorName;

            DataColumn dcproductName = new DataColumn("ProductName", typeof(String));
            dataTable.Columns.Add(dcproductName);
            dataTable.Rows[0]["ProductName"] = productName;


            DataColumn area = new DataColumn("areaName", typeof(String));
            dataTable.Columns.Add(area);
            dataTable.Rows[0]["areaName"] = areaName;

            DataColumn dcBatchstart = new DataColumn("BatchStartTime", typeof(String));
            dataTable.Columns.Add(dcBatchstart);
            string StartdateTime = (Convert.ToDateTime(batchStartDateTime)).ToString("dd.MM.yyyy HH:mm:ss");
            dataTable.Rows[0]["BatchStartTime"] = StartdateTime;

            DataColumn dcBatchEnd = new DataColumn("BatchEndTime", typeof(String));
            dataTable.Columns.Add(dcBatchEnd);
            string enddateTime = (Convert.ToDateTime(batchEndDateTime)).ToString("dd.MM.yyyy HH:mm:ss");

            dataTable.Rows[0]["BatchEndTime"] = enddateTime;
            dataTable.Rows[0]["BatchEndTime"] = enddateTime;

            return dataTable;
        }


        public DataTable AddAdditionalColumns(DataTable dataTable, string customta, string customva, string TimeInterval1, string tacceptance, string stepnum, string Venturi_Acceptance1, string Venturi_Acceptance2, string Grinding_Acceptance1, string Grinding_Acceptance2)
        {
            string datetime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

            DataColumn dc3 = new DataColumn("customva", typeof(String));
            dataTable.Columns.Add(dc3);
            dataTable.Rows[0]["customva"] = customva;

            DataColumn dc4 = new DataColumn("customta", typeof(String));
            dataTable.Columns.Add(dc4);
            dataTable.Rows[0]["customta"] = customta;

            DataColumn dc6 = new DataColumn("timeInterval", typeof(String));
            dataTable.Columns.Add(dc6);
            dataTable.Rows[0]["timeInterval"] = TimeInterval1;

            DataColumn dc7 = new DataColumn("tacceptance", typeof(String));
            dataTable.Columns.Add(dc7);
            dataTable.Rows[0]["tacceptance"] = tacceptance;

            DataColumn dc8 = new DataColumn("stepnum", typeof(String));
            dataTable.Columns.Add(dc8);
            dataTable.Rows[0]["stepnum"] = stepnum;

            DataColumn dc13 = new DataColumn("Venturi_Acceptance1", typeof(String));
            dataTable.Columns.Add(dc13);
            dataTable.Rows[0]["Venturi_Acceptance1"] = Venturi_Acceptance1;

            DataColumn dc14 = new DataColumn("Venturi_Acceptance2", typeof(String));
            dataTable.Columns.Add(dc14);
            dataTable.Rows[0]["Venturi_Acceptance2"] = Venturi_Acceptance2;

            DataColumn dc15 = new DataColumn("Grinding_Acceptance1", typeof(String));
            dataTable.Columns.Add(dc15);
            dataTable.Rows[0]["Grinding_Acceptance1"] = Grinding_Acceptance1;

            DataColumn dc16 = new DataColumn("Grinding_Acceptance2", typeof(String));
            dataTable.Columns.Add(dc16);
            dataTable.Rows[0]["Grinding_Acceptance2"] = Grinding_Acceptance2;

            return dataTable;
        }



        public DataTable AddAdditionalColumnsCustomReactor(DataTable dataTable, string custombatchnumber, string customproductname, string custombmrpagenumber, string custombmrsection, string batchno, string batchStartDateTime, string batchEndDateTime, string Venturi_Acceptance1, string Venturi_Acceptance2, string Grinding_Acceptance1, string Grinding_Acceptance2)
        {
            string datetime = DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss");

            DataColumn dc3 = new DataColumn("custombatchnumber", typeof(String));
            dataTable.Columns.Add(dc3);
            dataTable.Rows[0]["customva"] = custombatchnumber;

            DataColumn dc4 = new DataColumn("customproductname", typeof(String));
            dataTable.Columns.Add(dc4);
            dataTable.Rows[0]["customta"] = customproductname;

            DataColumn dc6 = new DataColumn("custombmrpagenumber", typeof(String));
            dataTable.Columns.Add(dc6);
            dataTable.Rows[0]["timeInterval"] = custombmrpagenumber;

            DataColumn dc7 = new DataColumn("custombmrsection", typeof(String));
            dataTable.Columns.Add(dc7);
            dataTable.Rows[0]["tacceptance"] = custombmrsection;

            DataColumn dc8 = new DataColumn("batchno", typeof(String));
            dataTable.Columns.Add(dc8);
            dataTable.Rows[0]["stepnum"] = batchno;

            DataColumn dc9 = new DataColumn("batchStartDateTime", typeof(String));
            dataTable.Columns.Add(dc8);
            dataTable.Rows[0]["stepnum"] = batchStartDateTime;

            DataColumn dc10 = new DataColumn("batchEndDateTime", typeof(String));
            dataTable.Columns.Add(dc8);
            dataTable.Rows[0]["stepnum"] = batchEndDateTime;

            DataColumn dc13 = new DataColumn("Venturi_Acceptance1", typeof(String));
            dataTable.Columns.Add(dc13);
            dataTable.Rows[0]["Venturi_Acceptance1"] = Venturi_Acceptance1;

            DataColumn dc14 = new DataColumn("Venturi_Acceptance2", typeof(String));
            dataTable.Columns.Add(dc14);
            dataTable.Rows[0]["Venturi_Acceptance2"] = Venturi_Acceptance2;

            DataColumn dc15 = new DataColumn("Grinding_Acceptance1", typeof(String));
            dataTable.Columns.Add(dc15);
            dataTable.Rows[0]["Grinding_Acceptance1"] = Grinding_Acceptance1;

            DataColumn dc16 = new DataColumn("Grinding_Acceptance2", typeof(String));
            dataTable.Columns.Add(dc16);
            dataTable.Rows[0]["Grinding_Acceptance2"] = Grinding_Acceptance2;

            return dataTable;
        }

#endregion
        public int ConvertTimeIntervalToSeconds(string timeInterval)
        {
            switch (timeInterval)
            {
                case "1 Sec":
                    return 1;
                case "1 Min":
                    return 1;
                case "2 Min":
                    return 2;
                case "3 Min":
                    return 3;
                case "5 Min":
                    return 5;
                case "10 Min":
                    return 10;
                case "15 Min":
                    return 15;
                case "20 Min":
                    return 16;
                case "25 Min":
                    return 25;
                case "30 Min":
                    return 30;
                case "35 Min":
                    return 35;
                case "40 Min":
                    return 40;
                case "45 Min":
                    return 45;
                case "50 Min":
                    return 50;
                case "55 Min":
                    return 55;
                case "60 Min":
                    return 60;
                default:
                    return 0; // Or throw an exception for invalid input
            }
        }

        public string GetRDLCNameReactorWise(string reactorName)
        {
            var reactorReports = new Dictionary<string, string>
    {
        { "E4-GLR-01", "TrendReportE4_GLR_01" },
        { "E4-GLR-44", "TrendReportE4_GLR_44" },
        { "E4-GLR-06", "TrendReportE4_GLR_06" },
        { "E4-GLR-04", "TrendReportE4_GLR_04" },
        { "E4-SSR-07", "TrendReportE4_SSR_07" },
        { "E4-SSR-09", "TrendReportE4_SSR_09" },
        { "E4-SSR-11", "TrendReportE4_SSR_11" },
        { "E4-GLR-12", "TrendReportE4_GLR_12" },
        { "E4-GLR-13", "TrendReportE4_GLR_13" },
        { "E4-GLR-310", "TrendReportE4_GLR_310" },
        { "E4-GLR-17", "TrendReportE4_GLR_17" },
        { "E4-SSR-19", "TrendReportE4_SSR_19" },
        { "E4-GLR-28", "TrendReportE3_GLR_28" },
        { "E4-SSR-29", "TrendReportE4_SSR_29" },
        { "E4-SSR-43", "TrendReportE4_SSR_43" },
        { "E4-SSR-44", "TrendReportE4_SSR_44" },
        { "E4-GLR-23", "TrendReportE4-GLR-23" },
        { "E4-GLR-45", "TrendReportE4_GLR_45" },
        { "E4-GLR-348", "TrendReportE4_GLR_348" },
        { "E4-SSR-311", "TrendReportE4-SSR-311" },
        { "E4-SSR-24", "TrendReportE4-SSR-24" },
        { "E4-SSR-368", "TrendReportE4-SSR-368" },
        { "E4-GLR-32", "TrendReportE4-GLR-32" },
        { "E4-GLR-33", "TrendReportE4-GLR-33" },
        { "E4-GLR-36", "TrendReportE4-GLR-36" },
        { "E4-SSR-37", "TrendReportE4-SSR-37" },
        { "E4-JTM-86", "TrendReportE4_JMM_86" },
        { "E4-JTM-87", "TrendReportE4_JMM_87" },
        { "E4-ANFD-63", "TrendReportE4-ANFD-63" },
        { "E4-ANFD-375", "TrendReportE4-ANFD-375" },
        { "E4-ANFD-371", "TrendReportE4-ANFD-371" },
        { "E4-VTD-77", "TrendReportE4-VTD-77" },
        { "OGB88", "TrendReport_OGB88" },
        { "OGB73", "TrendReport_OGB73" },
        { "E4-MM-70", "TrendReport_MM70" },
        { "E4-MM-78", "TrendReport_MM78" },
        { "E4-MM-82", "TrendReport_MM82" }
    };

            if (reactorReports.TryGetValue(reactorName, out string rdlcReportname))
            {
                return rdlcReportname;
            }

            return "DefaultReportName"; // Return a default report name or handle the case where the reactor name is not found
        }

    }
}