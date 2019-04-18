using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;

namespace Tutorial.SqlConn
{
    class DBUtils
    {
        public static SqlConnection GetDBConnection()
        {
            string path = @"resources\databasepath.txt";
            string dbp,s = "";
            int n = 0;

            using (StreamReader sr = new StreamReader(path))
            {
                dbp = (sr.ReadToEnd());
            }
            n = dbp.IndexOf(";");
            s = dbp.Substring(0, n);
            dbp = dbp.Remove(0, n+1);

            string datasource = @"" + s;
            string database = dbp;

            // string datasource = @"DESKTOP-588FCHK\SQLEXPRESS";
            // string database = "OutputTableDB";
            //string username = "";
            //string password = "";

            return DBSQLServerUtils.GetDBConnection(datasource, database);
        }
    }

}