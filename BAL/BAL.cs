using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.Odbc;
using System.Data;
using System.Windows.Forms;
using Cipla_Bangalore_API4_091.DAL;

namespace Cipla_Bangalore_API4_091.BAL
{
    class BAL
    {
        private readonly DAL.DAL DAL;
        public BAL()
        {
            DAL = new DAL.DAL();
        }

        public DataTable GetTimeIntervals(string connectionString)
        {
            return DAL.TimeintervalFetch(connectionString);
        }
        public DataTable FetchBatchNoData(string connectionString, string queryText)
        {
            return DAL.FetchBatchNoData(connectionString, queryText);
        }
        public DataTable GetBatchStartEndDateTime(string connectionString, string viewName, string batchNo, string fromDateTime, string toDateTime)
        {
            string queryText = $"SELECT Min(DateAndTime) as BatchStartDateTime, Max(DateAndTime) as BatchEndDateTime FROM {viewName} WHERE BatchNo='{batchNo}' AND DateAndTime BETWEEN '{fromDateTime}' AND '{toDateTime}'";
            return DAL.FetchBatchNoData(connectionString, queryText);
        }
    
        public DataTable GetProductName(string btstrt, string btend, string reactorOrDryerName, string connectionString,string viewName)
        {
            return DAL.GetProductName(btstrt, btend, reactorOrDryerName, connectionString, viewName);
        }

        public int ConvertTimeIntervalToSeconds(string timeInterval)
        {
            switch (timeInterval)
            {
                case "1 Sec":
                    return 1;
                case "1 Min":
                    return 60;
                case "2 Min":
                    return 120;
                case "3 Min":
                    return 180;
                case "5 Min":
                    return 300;
                case "10 Min":
                    return 600;
                case "15 Min":
                    return 900;
                case "20 Min":
                    return 1200;
                case "25 Min":
                    return 1500;
                case "30 Min":
                    return 1800;
                case "35 Min":
                    return 2100;
                case "40 Min":
                    return 2400;
                case "45 Min":
                    return 2700;
                case "50 Min":
                    return 3000;
                case "55 Min":
                    return 3300;
                case "60 Min":
                    return 3600;
                default:
                    return 0; // Or throw an exception for invalid input
            }
        }


        #region add coloumns

        
        public DataTable FetchBatchData(string connectionString, string queryText)
        {
            DataTable batchData = new DataTable();
            try
            {
                using (OdbcConnection conn = new OdbcConnection(connectionString))
                {
                    conn.Open();
                    OdbcCommand query = new OdbcCommand(queryText, conn);
                    OdbcDataAdapter adp = new OdbcDataAdapter(query);
                    query.CommandTimeout = 60000;
                    adp.Fill(batchData);
                    conn.Close();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            return batchData;
        }
     


        #endregion

    }
}

