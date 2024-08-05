using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;
using System.Data;
using System.Windows.Forms;

namespace Cipla_Bangalore_API4_091.DAL
{
    class DAL
    {
        public string connectionString = "";

        public DataTable TimeintervalFetch(string connectionString)
        {
            DataTable dt1 = new DataTable();

            try
            {
                using (OdbcConnection conn = new OdbcConnection(connectionString))
                {
                    conn.Open();
                    string query = "SELECT TimeVal, TimeInterval FROM tbl_TimeInterval";
                    using (OdbcCommand cmd = new OdbcCommand(query, conn))
                    {
                        using (OdbcDataAdapter adp = new OdbcDataAdapter(cmd))
                        {
                            adp.Fill(dt1);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return dt1;
        }

        public DataTable FetchBatchNoData(string connetionString, string queryText)
        {
            DataTable batchData = new DataTable();
            try
            {
                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    Conn.Open();
                    OdbcCommand query = new OdbcCommand(queryText, Conn);
                    OdbcDataAdapter adp = new OdbcDataAdapter(query);
                    query.CommandTimeout = 30000;
                    adp.Fill(batchData);
                    Conn.Close();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return batchData;
        }

      

        public DataTable GetProductName(string batchstrt, string batchend, string reactorName, string connetionString,string viewName)
        {
            try
            {
                DataTable dt = new DataTable();
                using (OdbcConnection Conn = new OdbcConnection(connetionString))
                {
                    int TagProductName = 0;

                    #region Product Code Tag Mapping


                    //if (reactorName == "E6-SSR-01")
                    //{
                    //    TagProductName = 8;
                    //}
                    //else if (reactorName == "E6-SSR-04")
                    //{
                    //    TagProductName = 07;
                    //}
                    //else if (reactorName == "E6-SSR-02")
                    //{
                    //    TagProductName = 08;
                    //}
                    //else if (reactorName == "E6-SSR-03")
                    //{
                    //    TagProductName = 08;
                    //}
                    //else if (reactorName == "E6-SSR-05")
                    //{
                    //    TagProductName = 08;
                    //}
                    //else if (reactorName == "E6-SSR-06")
                    //{
                    //    TagProductName = 08;
                    //}
                    ////Dryer name
                    //else if (DryerName == "E6-ANFD-113")
                    //{
                    //    TagProductName = 08;
                    //}
                    //else if (DryerName == "E6-ANFD-114")
                    //{
                    //    TagProductName = 08;
                    //}

                    //else if (SolventName == "E6-SSR-99/MMA")
                    //{
                    //    TagProductName = 03;
                    //}
                    //else if (SolventName == "E6-SSR-03-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-GLR-08-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-GLR-09-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-GLR-10-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-SSR-11-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-SSR-99-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}

                    //else if (SolventName == "E6-DT-66-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-DT-62-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-DT-63-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-DT-65-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-DT-64-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    //else if (SolventName == "E6-DT-76-SOLVENT")
                    //{
                    //    TagProductName = 05;
                    //}
                    #endregion

                    #region Product Code Tag Mapping

                    // Define dictionaries for mapping names to TagProductName
                    var reactorTagMapping = new Dictionary<string, int>
                    {
                        { "E6-SSR-01", 8 },
                        { "E6-SSR-04", 7 },
                        { "E6-SSR-02", 8 },
                        { "E6-SSR-03", 8 },
                        { "E6-SSR-05", 8 },
                        { "E6-SSR-06", 8 }
                    };

                    var dryerTagMapping = new Dictionary<string, int>
                    {
                        { "E6-ANFD-113", 8 },
                        { "E6-ANFD-114", 8 }
                    };

                    var solventTagMapping = new Dictionary<string, int>
                    {
                        { "E6-SSR-99/MMA", 3 },
                        { "E6-SSR-03-SOLVENT", 5 },
                        { "E6-GLR-08-SOLVENT", 5 },
                        { "E6-GLR-09-SOLVENT", 5 },
                        { "E6-GLR-10-SOLVENT", 5 },
                        { "E6-SSR-11-SOLVENT", 5 },
                        { "E6-SSR-99-SOLVENT", 5 },
                        { "E6-DT-66-SOLVENT", 5 },
                        { "E6-DT-62-SOLVENT", 5 },
                        { "E6-DT-63-SOLVENT", 5 },
                        { "E6-DT-65-SOLVENT", 5 },
                        { "E6-DT-64-SOLVENT", 5 },
                        { "E6-DT-76-SOLVENT", 5 }
                    };

                    // Check mappings for ReactorName, DryerName, and SolventName
                    if (reactorTagMapping.TryGetValue(reactorName, out TagProductName) ||
                        dryerTagMapping.TryGetValue(reactorName, out TagProductName) ||
                        solventTagMapping.TryGetValue(reactorName, out TagProductName))
                    {
                        // TagProductName is successfully set by one of the mappings
                    }
                    #endregion


                    Conn.Open();

                    OdbcCommand cmd = new OdbcCommand("SELECT * FROM StringTable WHERE DateAndTime BETWEEN '" + batchstrt + "' AND '" + batchend + "' and TagIndex=" + TagProductName + "", Conn);  // Fetch product name

                    OdbcDataAdapter adap = new OdbcDataAdapter(cmd);
                    
                    adap.Fill(dt);
                    
                }
                return dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return new DataTable();
            }
        }
    }
}
