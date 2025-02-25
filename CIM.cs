using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using Cimetrix.Value;
using EMSERVICELib;
using System.Collections;
using System.Runtime.InteropServices;
using VALUELib;

namespace BiometricServer
{
    public static class CIM
    {
        private static CMyCimetrix _CMyCimetrix;
        private static SQLiteConnection conn;
        private static Hashtable dataTypes = new Hashtable();
        private static Hashtable DV_SV_EC_Types = new Hashtable();
        public const int _PrimaryConnection = 1;


        //----------------------------------------------------------------
        public static void INitializeCIM(CMyCimetrix _myCMyCimetrix, SQLiteConnection connection)
        //----------------------------------------------------------------
        {
            conn = connection;
            _CMyCimetrix = _myCMyCimetrix;

        }

        //----------------------------------------------------------------
        public static string GetDataTypeFromId(int varid)
        //----------------------------------------------------------------
        {
            try
            {
                string datatype = (string)dataTypes[varid];
                return datatype;
            }
            catch (Exception)
            {
                LogManager.DefaultLogger.Info("Datatype for varid " + varid .ToString() + "not found");
            }

            return String.Empty;
        }

        //----------------------------------------------------------------
        public static string Get_DV_SV_EC_TypeFromId(int varid)
        //----------------------------------------------------------------
        {
            try
            {
                string type = (string)DV_SV_EC_Types[varid];
                return type;
            }
            catch (Exception)
            {
                LogManager.DefaultLogger.Info("Get_DV_SV_EC_TypeFromId for varid " + varid.ToString() + "not found");
            }

            return String.Empty;
        }

        /// <summary>
        /// This function is to be called after object _CMyCimetrix is initialized
        /// and before the GEM interface is enabled during final initialization. 
        /// 
        /// This function demonstrates how to add more variables, collection events
        /// to the GEM interface at runtime. 
        /// </summary>
        public static void DefineDynamicGEMItems()
        {
            string[] dvList = new string[0];
            VALUELib.CxValueObject svidList = new VALUELib.CxValueObject();
            VALUELib.CxValueObject ecidList = new VALUELib.CxValueObject();


            VALUELib.CxValueObject emptyI1 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyI1.SetDataType(0, 0, VALUELib.ValueType.I1);

            VALUELib.CxValueObject emptyI2 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyI2.SetDataType(0, 0, VALUELib.ValueType.I2);

            VALUELib.CxValueObject emptyI4 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyI4.SetDataType(0, 0, VALUELib.ValueType.I4);

            VALUELib.CxValueObject emptyI8 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyI8.SetDataType(0, 0, VALUELib.ValueType.I8);

            VALUELib.CxValueObject emptyU1 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyU1.SetDataType(0, 0, VALUELib.ValueType.U1);

            VALUELib.CxValueObject emptyU2 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyU2.SetDataType(0, 0, VALUELib.ValueType.U2);

            VALUELib.CxValueObject emptyU4 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyU4.SetDataType(0, 0, VALUELib.ValueType.U4);

            VALUELib.CxValueObject emptyU8 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyU8.SetDataType(0, 0, VALUELib.ValueType.U8);

            VALUELib.CxValueObject emptyF4 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyF4.SetDataType(0, 0, VALUELib.ValueType.F4);

            VALUELib.CxValueObject emptyF8 = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyF8.SetDataType(0, 0, VALUELib.ValueType.F8);

            VALUELib.CxValueObject emptyA = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyA.SetDataType(0, 0, VALUELib.ValueType.A);

            VALUELib.CxValueObject emptyW = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyW.SetDataType(0, 0, VALUELib.ValueType.W);

            VALUELib.CxValueObject emptyBo = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyBo.SetDataType(0, 0, VALUELib.ValueType.Bo);

            VALUELib.CxValueObject emptyBi = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyBi.SetDataType(0, 0, VALUELib.ValueType.Bi);

            VALUELib.CxValueObject emptyL = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyL.SetDataType(0, 0, VALUELib.ValueType.L);

            VALUELib.CxValueObject emptyJ = _CMyCimetrix._CxClientClerk.CreateValueObject();
            emptyJ.SetDataType(0, 0, VALUELib.ValueType.J);

            VALUELib.CxValueObject emptyValue = _CMyCimetrix._CxClientClerk.CreateValueObject();


            // Register DV values
            //--------------------------------------------------------------------------------------
            LogManager.DefaultLogger.Info("Register DV values from table DV -> begin");

            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT varid, name, datatype, description FROM DV WHERE varid > 7000";
            SQLiteDataReader rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                Int32 varid = rdr.GetInt32(0);
                string name = rdr.GetString(1);
                string datatype = rdr.GetString(2);
                string description = rdr.GetString(3);

                dataTypes.Add(varid, datatype);
                DV_SV_EC_Types.Add(varid, "DV");

                switch (datatype)
                {
                    case "I1":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.I1, -1, "", emptyValue, emptyValue, emptyI1, false, false);
                        break;

                    case "I2":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.I2, -1, "", emptyValue, emptyValue, emptyI2, false, false);
                        break;

                    case "I4":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.I4, -1, "", emptyValue, emptyValue, emptyI4, false, false);
                        break;

                    case "I8":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.I8, -1, "", emptyValue, emptyValue, emptyI8, false, false);
                        break;

                    case "U1":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.U1, -1, "", emptyValue, emptyValue, emptyU1, false, false);
                        break;

                    case "U2":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.U2, -1, "", emptyValue, emptyValue, emptyU2, false, false);
                        break;

                    case "U4":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.U4, -1, "", emptyValue, emptyValue, emptyU4, false, false);
                        break;

                    case "U8":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.U8, -1, "", emptyValue, emptyValue, emptyU8, false, false);
                        break;

                    case "F4":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.F4, -1, "", emptyValue, emptyValue, emptyF4, false, false);
                        break;

                    case "F8":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.F8, -1, "", emptyValue, emptyValue, emptyF8, false, false);
                        break;

                    case "A":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.A, -1, "", emptyValue, emptyValue, emptyA, false, false);
                        break;

                    case "W":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.W, -1, "", emptyValue, emptyValue, emptyW, false, false);
                        break;

                    case "Bo":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.Bo, -1, "", emptyValue, emptyValue, emptyBo, false, false);
                        break;

                    case "Bi":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.Bi, -1, "", emptyValue, emptyValue, emptyBi, false, false);
                        break;

                    case "L":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.L, -1, "", emptyValue, emptyValue, emptyL, false, false);
                        break;

                    case "J":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.DV, VALUELib.ValueType.J, -1, "", emptyValue, emptyValue, emptyJ, false, false);
                        break;
                }
            }
            LogManager.DefaultLogger.Info("Register DV values from table DV -> end");

            cmd.Dispose();
            rdr.Dispose();

            // Register EC values
            //--------------------------------------------------------------------------------------
            LogManager.DefaultLogger.Info("Register EC values from table EC -> begin");

            cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT varid, name, datatype, description FROM EC";
            rdr = cmd.ExecuteReader();
            
            Int32 varid_EC;
            string name_EC;

            try
            {
                while (rdr.Read())
                {
                    varid_EC = rdr.GetInt32(0);
                    name_EC = rdr.GetString(1);
                    string datatype = rdr.GetString(2);
                    string description = rdr.GetString(3);

                    dataTypes.Add(varid_EC, datatype);
                    DV_SV_EC_Types.Add(varid_EC, "EC");

                    switch (datatype)
                    {
                        case "I1":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.I1, -1, "", emptyValue, emptyValue, emptyI1, false, false);
                            break;

                        case "I2":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.I2, -1, "", emptyValue, emptyValue, emptyI2, false, false);
                            break;

                        case "I4":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.I4, -1, "", emptyValue, emptyValue, emptyI4, false, false);
                            break;

                        case "I8":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.I8, -1, "", emptyValue, emptyValue, emptyI8, false, false);
                            break;

                        case "U1":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.U1, -1, "", emptyValue, emptyValue, emptyU1, false, false);
                            break;

                        case "U2":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.U2, -1, "", emptyValue, emptyValue, emptyU2, false, false);
                            break;

                        case "U4":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.U4, -1, "", emptyValue, emptyValue, emptyU4, false, false);
                            break;

                        case "U8":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.U8, -1, "", emptyValue, emptyValue, emptyU8, false, false);
                            break;

                        case "F4":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.F4, -1, "", emptyValue, emptyValue, emptyF4, false, false);
                            break;

                        case "F8":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.F8, -1, "", emptyValue, emptyValue, emptyF8, false, false);
                            break;

                        case "A":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.A, -1, "", emptyValue, emptyValue, emptyA, false, false);
                            break;

                        case "W":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.W, -1, "", emptyValue, emptyValue, emptyW, false, false);
                            break;

                        case "Bo":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.Bo, -1, "", emptyValue, emptyValue, emptyBo, false, false);
                            break;

                        case "Bi":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.Bi, -1, "", emptyValue, emptyValue, emptyBi, false, false);
                            break;

                        case "L":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.L, -1, "", emptyValue, emptyValue, emptyL, false, false);
                            break;

                        case "J":
                            _CMyCimetrix.DefineVariable(name_EC, varid_EC, description, VarType.EC, VALUELib.ValueType.J, -1, "", emptyValue, emptyValue, emptyJ, false, false);
                            break;
                    }

                    ecidList.AppendValueU4(0, varid_EC);

                    _CMyCimetrix._CxClientClerk.RegisterByteStreamValueChangedHandler(1, varid_EC, _CMyCimetrix);
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Info("Exception on Register EC values:\r\n" + ex.ToString());
            }



            LogManager.DefaultLogger.Info("Register EC values from table EC -> end");

            cmd.Dispose();
            rdr.Dispose();

            LogManager.DefaultLogger.Info("_CxClientClerk.RegisterGetValuesToByteBufferHandler for ECs -> begin");
            _CMyCimetrix._CxClientClerk.RegisterGetValuesToByteBufferHandler(0, ecidList, _CMyCimetrix);
            LogManager.DefaultLogger.Info("_CxClientClerk.RegisterGetValuesToByteBufferHandler for ECs -> end");

            // Register SV values
            //--------------------------------------------------------------------------------------
            LogManager.DefaultLogger.Info("Register SV values from table SV -> begin");

            cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT varid, name, datatype, description FROM SV WHERE varid > 4000";
            rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                Int32 varid = rdr.GetInt32(0);
                string name = rdr.GetString(1);
                string datatype = rdr.GetString(2);
                string description = rdr.GetString(3);

                dataTypes.Add(varid, datatype);
                DV_SV_EC_Types.Add(varid, "SV");

                switch (datatype)
                {
                    case "I1":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.I1, -1, "", emptyValue, emptyValue, emptyI1, false, false);
                        break;

                    case "I2":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.I2, -1, "", emptyValue, emptyValue, emptyI2, false, false);
                        break;

                    case "I4":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.I4, -1, "", emptyValue, emptyValue, emptyI4, false, false);
                        break;

                    case "I8":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.I8, -1, "", emptyValue, emptyValue, emptyI8, false, false);
                        break;

                    case "U1":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.U1, -1, "", emptyValue, emptyValue, emptyU1, false, false);
                        break;

                    case "U2":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.U2, -1, "", emptyValue, emptyValue, emptyU2, false, false);
                        break;

                    case "U4":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.U4, -1, "", emptyValue, emptyValue, emptyU4, false, false);
                        break;

                    case "U8":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.U8, -1, "", emptyValue, emptyValue, emptyU8, false, false);
                        break;

                    case "F4":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.F4, -1, "", emptyValue, emptyValue, emptyF4, false, false);
                        break;

                    case "F8":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.F8, -1, "", emptyValue, emptyValue, emptyF8, false, false);
                        break;

                    case "A":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.A, -1, "", emptyValue, emptyValue, emptyA, false, false);
                        break;

                    case "W":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.W, -1, "", emptyValue, emptyValue, emptyW, false, false);
                        break;

                    case "Bo":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.Bo, -1, "", emptyValue, emptyValue, emptyBo, false, false);
                        break;

                    case "Bi":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.Bi, -1, "", emptyValue, emptyValue, emptyBi, false, false);
                        break;

                    case "L":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.L, -1, "", emptyValue, emptyValue, emptyL, false, false);
                        break;

                    case "J":
                        _CMyCimetrix.DefineVariable(name, varid, description, VarType.SV, VALUELib.ValueType.J, -1, "", emptyValue, emptyValue, emptyJ, false, false);
                        break;
                }

                svidList.AppendValueU4(0, varid);
            }

            LogManager.DefaultLogger.Info("Register SV values from table SV -> end");

            cmd.Dispose();
            rdr.Dispose();

            LogManager.DefaultLogger.Info("_CxClientClerk.RegisterGetValuesToByteBufferHandler for SVs -> begin");
           
            _CMyCimetrix._CxClientClerk.RegisterGetValuesToByteBufferHandler(0, svidList, _CMyCimetrix);
            
            LogManager.DefaultLogger.Info("_CxClientClerk.RegisterGetValuesToByteBufferHandler for SVs -> end");

            // Register Collection Events
            //--------------------------------------------------------------------------------------
            LogManager.DefaultLogger.Info("Register CE events from table Events -> begin");

            cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT CEID, name, description FROM CE WHERE CEID > 8000";
            rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                Int32 CEID = rdr.GetInt32(0);
                string name = rdr.GetString(1);
                string description = rdr.GetString(2);

                _CMyCimetrix.DefineCollectionEvent(name, CEID, description, dvList);
            }
            LogManager.DefaultLogger.Info("Register CE events from table Events -> end");

            cmd.Dispose();
            rdr.Dispose();

            /**
            
            // Register Alarms 
            //--------------------------------------------------------------------------------------
            LogManager.DefaultLogger.Info("Register Alarms from table AL -> begin");

            cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT alarmid, name, description, categorie, CEID_ON, CEID_OFF FROM AL";
            rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                Int32 alarmid = rdr.GetInt32(0);
                string name = rdr.GetString(1);
                string description = rdr.GetString(2);
                Int32 categorie = rdr.GetInt32(3);
                Int32 CEID_ON = rdr.GetInt32(4);
                Int32 CEID_OFF = rdr.GetInt32(5);

                // Register Collection event for alarm on/off
                _CMyCimetrix.DefineCollectionEvent(name + "SET", CEID_ON, name + " SET collection event", dvList);
                _CMyCimetrix.DefineCollectionEvent(name + "CLEAR", CEID_OFF, name + " CLEAR collection event", dvList);

                // Register Alarm
                _CMyCimetrix.DefineAlarm(name, alarmid, description, description, CEID_ON, CEID_OFF, categorie);
            }
            LogManager.DefaultLogger.Info("Register Alarms from table AL -> end");

            cmd.Dispose();
            rdr.Dispose();

            **/
        }

        //----------------------------------------------------------------
        public static void UnregisterSVs()
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Info("UnRegister SV values -> begin");

            VALUELib.CxValueObject svidList = new VALUELib.CxValueObject();

            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT varid FROM SV WHERE varid > 4000";
            SQLiteDataReader rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                Int32 varid = rdr.GetInt32(0);
                svidList.AppendValueU4(0, varid);
            }

            _CMyCimetrix._CxClientClerk.UnregisterGetValuesToByteBufferHandler(0, svidList);

            cmd.Dispose();
            rdr.Dispose();
            
            LogManager.DefaultLogger.Info("UnRegister SV values -> end");
        }

        //----------------------------------------------------------------
        public static void UnregisterECs()
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Info("UnRegister EC values -> begin");

            VALUELib.CxValueObject svidList = new VALUELib.CxValueObject();

            SQLiteCommand cmd = conn.CreateCommand();
            cmd.CommandText = "SELECT varid FROM EC";
            SQLiteDataReader rdr = cmd.ExecuteReader();

            while (rdr.Read())
            {
                Int32 varid = rdr.GetInt32(0);
                _CMyCimetrix._CxClientClerk.UnregisterByteStreamValueChangedHandler(1, varid);
            }

            cmd.Dispose();
            rdr.Dispose();

            LogManager.DefaultLogger.Info("UnRegister EC values -> end");
        }

        //----------------------------------------------------------------
        public static void AlarmSET(int id, string text)
        //----------------------------------------------------------------
        {
            _CMyCimetrix.AlarmSET(id, text);
        }

        //----------------------------------------------------------------
        public static void AlarmSET(string name, string text)
        //----------------------------------------------------------------
        {
            _CMyCimetrix.AlarmSET(name, text);
        }

        //----------------------------------------------------------------
        public static void AlarmCLEAR(int id)
        //----------------------------------------------------------------
        {
            _CMyCimetrix.AlarmCLEAR(id);
        }

        //----------------------------------------------------------------
        public static void AlarmCLEAR(string name, string text)
        //----------------------------------------------------------------
        {
            _CMyCimetrix.ClearAlarm(name, text);
        }

        //----------------------------------------------------------------
        public static void SendCollectionEventWithData(string[] variableNames, CxValue[] variableValues, int CEID)
        //----------------------------------------------------------------
        {
            _CMyCimetrix.SendCollectionEventWithData(0, variableNames, variableValues, CEID);
        }

        //----------------------------------------------------------------
        public static void SendCollectionEventWithData(string[] variableNames, CxValue[] variableValues, string eventName)
        //----------------------------------------------------------------
        {
            _CMyCimetrix.SendCollectionEventWithData(0, variableNames, variableValues, eventName);
        }

        //----------------------------------------------------------------
        public static void CommunicationStateEnable()
        //----------------------------------------------------------------
        {
            _CMyCimetrix.GEMStateCommunicationStateEnable(_PrimaryConnection, true);
        }

        //----------------------------------------------------------------
        public static void CommunicationStateDisable()
        //----------------------------------------------------------------
        {
            _CMyCimetrix.GEMStateCommunicationStateEnable(_PrimaryConnection, false);
        }

        //----------------------------------------------------------------
        public static void Offline()
        //----------------------------------------------------------------
        {
            _CMyCimetrix.GEMStateControlStateOnline(_PrimaryConnection, false);
        }

        //----------------------------------------------------------------
        public static void Online()
        //----------------------------------------------------------------
        {
            _CMyCimetrix.GEMStateControlStateOnline(_PrimaryConnection, true);
        }

        //----------------------------------------------------------------
        public static void Local()
        //----------------------------------------------------------------
        {
            _CMyCimetrix.GEMStateControlStateRemote(_PrimaryConnection, false);
        }

        //----------------------------------------------------------------
        public static void Remote()
        //----------------------------------------------------------------
        {
            _CMyCimetrix.GEMStateControlStateRemote(_PrimaryConnection, true);
        }




    }
}
