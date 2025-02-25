using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;

namespace BiometricServer
{
    public static class SQLight
    {
        private static  SQLiteConnection connection;

        public static bool Initialize()
        {
            connection = new SQLiteConnection("Data Source=db_S_BIOMETRIC_V1.0.db; Version = 3; New = True; Compress = True; ");

            try
            {
                connection.Open();
            }
            catch (Exception e)
            {
                LogManager.DefaultLogger.Error("Exception on SQLiteConnection Database opening:\r\n" + e.ToString());
                return false;
            }

            return true;
        }

        //----------------------------------------------------------------
        public static SQLiteConnection Connection
        //----------------------------------------------------------------
        {
            get
            {
                return connection;
            }
        }

        //----------------------------------------------------------------
        public static void Close()
        //----------------------------------------------------------------
        {
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();
                connection.Dispose();
                connection = null;
            }
        }

    }
}
