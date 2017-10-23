#define DEBUG

using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Oracle.DataAccess.Client;
using System.Configuration;

namespace ERPAppAddition.ERPAddition
{
    public class GV
    {
        private static string GetConnStr()
        {
#if !DEBUG
            return ConfigurationManager.ConnectionStrings["MES_NDMES_MESMGR"].ConnectionString;
#else
            return ConfigurationManager.ConnectionStrings["MES_NDTMES_MESMGR"].ConnectionString;
#endif
        }

        public static OracleConnection gOraCon = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_NDMES_MESMGR"].ConnectionString);
        public static OracleConnection gOraCo2 = new OracleConnection(ConfigurationManager.ConnectionStrings["MES_RPTMIT_RPTMIT"].ConnectionString);
        public static OracleCommand gOraCmd;
        public static OracleDataReader gOraDR;

        //public static string gStrPageTitle;
    }
}