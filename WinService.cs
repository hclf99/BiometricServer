using Cimetrix.Value;
using EMSERVICELib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SQLite;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Topshelf;

namespace BiometricServer
{
    public class WinService
    {
        public CMyCimetrix _CMyCimetrix;

        public const int _PrimaryConnection = 1;

        private SQLiteConnection conn;

        private CS1 cs1 = new CS1();
        private CS2 cs2 = new CS2();
        private CS3 cs3 = new CS3();
        private CS4 cs4 = new CS4();

        private Mutex mutex = new Mutex();
        private Mutex ethernetMutex = new Mutex();

        private readonly object _lockValue = new object();

        private string EquipID = "636-360";
        private UInt64 CmdSeqID = 0;
        private UInt64 SeqID = 0;

        public enum MyProcessState
        {
            SETUP = 0,
            IDLE = 1,
            EXECUTING = 2,
            PAUSE = 3
        };

        public MyProcessState _MyProcessState;


        // Processing Thread members 
        ThreadStart _ProcessingThreadStart;
        Thread _ProcessingThread;
        ManualResetEvent _EventSTART;
        ManualResetEvent _EventSTOP;
        bool _ProcessingThreadExit;


        //----------------------------------------------------------------
        public void Start()
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Directory = ReadSetting("LOGGERDIRECTORY"); ;

            if (SQLight.Initialize() == false)
            {
                LogManager.DefaultLogger.Info("SQLight: Database konnte nicht geöffnet werden");
                return;
            }

            conn = SQLight.Connection;

            cs1.Init(conn, ethernetMutex);
            cs2.Init(conn, ethernetMutex);
            cs3.Init(conn, ethernetMutex);
            cs4.Init(conn, ethernetMutex);

            var t = new Thread(InitializationMTA) { Name = "CimetrixInitialization" };
            t.SetApartmentState(ApartmentState.MTA);
            t.Start();

            int port1 = int.Parse(ReadSetting("SERVER_PORT1"));
            int port2 = int.Parse(ReadSetting("SERVER_PORT2"));
            int port3 = int.Parse(ReadSetting("SERVER_PORT3"));
            int port4 = int.Parse(ReadSetting("SERVER_PORT4"));

            // Activate TCP/IP server channel
            try
            {
                _ = cs1.StartListener();
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Info("TCPIP processor exception on C1 \r\n" + ex.Message);
            }

            // Activate TCP/IP client channel
            try
            {
                _ = cs1.StartClient();
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Info("TCPIP client exception\r\n" + ex.Message);
            }

            if (port2 != 0)
            {
                try
                {
                    _ = cs2.StartListener();
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Info("TCPIP processor exception on C2 \r\n" + ex.Message);
                }
            }


            if (port3 != 0)
            {
                try
                {
                    _ = cs3.StartListener();
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Info("TCPIP processor exception on C3 \r\n" + ex.Message);
                }
            }

            if (port4 != 0)
            {
                try
                {
                    _ = cs4.StartListener();
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Info("TCPIP processor exception on C4 \r\n" + ex.Message);
                }
            }

            LogManager.DefaultLogger.Info("==============================================================");
            LogManager.DefaultLogger.Info("Biometric Server (2.1 BETA) wurde gestartet");
            LogManager.DefaultLogger.Info("SQLight: Database Verbindung geöffnet");
            LogManager.DefaultLogger.Info("CimetrixInitialization wurde gestarted");
            LogManager.DefaultLogger.Info("TCP/IP Processor wurde gestartet");
        }

        //----------------------------------------------------------------
        public void Stop()
        //----------------------------------------------------------------
        {
            // Kill the processing thread
            _ProcessingThreadExit = true;

            if (_EventSTART != null)
                _EventSTART.Set();

            if (_EventSTOP != null)
                _EventSTOP.Set();

            _CMyCimetrix.OnRemoteCommand -= CMyCIMConnect_RemoteCommandHandler;
            _CMyCimetrix.OnGEMStateChange -= CMyCIMConnect_GEMStateChangeHandler;
            _CMyCimetrix.OnGetValueCallback -= CMyCIMConnect_GetValueCallbackHandler;
            _CMyCimetrix.OnTerminalService -= CMyCIMConnect_TerminalServiceHandler;
            _CMyCimetrix.OnLog -= CMyCIMConnect_LogHandler;
            _CMyCimetrix.OnCommEnableDisableSwitchHandler -= CMyCIMConnect_OnCommEnableDisableSwitchHandler;
            _CMyCimetrix.OnCtrlOfflineOnlineSwitchHandler -= CMyCIMConnect_CtrlOfflineOnlineSwitchHandler;
            _CMyCimetrix.OnCtrlLocalRemoteSwitchHandler -= CMyCIMConnect_CtrlLocalRemoteSwitchHandler;
            _CMyCimetrix.OnInitializationComplete -= _CMyCimetrix_onInitializationComplete;
            _CMyCimetrix.OnInitializationComplete -= _CMyCimetrix_onInitializationComplete;

            try
            {
                if (_CMyCimetrix._CxClientClerk == null)
                    return;

                CIM.UnregisterSVs();
                CIM.UnregisterECs();

                // Unregister for CommandCalled callbacks
                Log("_CxClientClerk.UnregisterCommandHandler JobCreate");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("JobCreate");

                Log("_CxClientClerk.UnregisterCommandHandler RESUME");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("RESUME");

                Log("_CxClientClerk.UnregisterCommandHandler PAUSE");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("PAUSE");

                Log("_CxClientClerk.UnregisterCommandHandler START");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("START");

                Log("_CxClientClerk.UnregisterCommandHandler STOP");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("STOP");

                Log("_CxClientClerk.UnregisterCommandHandler ABORT");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("ABORT");


                Log("_CxClientClerk.UnregisterCommandHandler PPSELECT");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("PPSELECT");

                Log("_CxClientClerk.UnregisterCommandHandler ExecuteRemoteCommand");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("ExecuteRemoteCommand");

                Log("_CxClientClerk.UnregisterCommandHandler GetProducts");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetProducts");

                Log("_CxClientClerk.UnregisterCommandHandler SelectProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("SelectProduct");

                Log("_CxClientClerk.UnregisterCommandHandler DownloadProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("DownloadProduct");

                Log("_CxClientClerk.UnregisterCommandHandler UploadProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("UploadProduct");

                Log("_CxClientClerk.UnregisterCommandHandler SetTerminalMessage");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("SetTerminalMessage");

                Log("_CxClientClerk.UnregisterCommandHandler GetUsers");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetUsers");

                Log("_CxClientClerk.UnregisterCommandHandler GetLoggedInUsers");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetLoggedInUsers");

                Log("_CxClientClerk.UnregisterCommandHandler CreateLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("CreateLot");

                Log("_CxClientClerk.UnregisterCommandHandler GetLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetLot");

                Log("_CxClientClerk.UnregisterCommandHandler GetLots");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetLot");

                Log("_CxClientClerk.UnregisterCommandHandler UpdateLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("UpdateLot");

                Log("_CxClientClerk.UnregisterCommandHandler DeleteLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("DeleteLot");

                Log("_CxClientClerk.UnregisterCommandHandler RenameProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("RenameProduct");

                Log("_CxClientClerk.UnregisterCommandHandler GetModuleProcessStates");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetModuleProcessStates");
            }
            catch (COMException ex)
            {
                Log(@"COMException unregistering CIMConnect handlers:\r\n" + ex.ToString());
            }
            finally
            {
                if (_CMyCimetrix._CxEMService != null)
                    _CMyCimetrix.Shutdown();

                // shutdown TCP/IP
                cs1.StopAndWait();
                cs2.StopAndWait();
                cs3.StopAndWait();
                cs4.StopAndWait();

                SQLight.Close();

                LogManager.StopAllAndWait();
            }
        }

        //----------------------------------------------------------------
        public void Shutdown()
        //----------------------------------------------------------------
        {
            // Kill the processing thread
            _ProcessingThreadExit = true;

            if (_EventSTART != null)
                _EventSTART.Set();

            if (_EventSTOP != null)
                _EventSTOP.Set();

            _CMyCimetrix.OnRemoteCommand -= CMyCIMConnect_RemoteCommandHandler;
            _CMyCimetrix.OnGEMStateChange -= CMyCIMConnect_GEMStateChangeHandler;
            _CMyCimetrix.OnGetValueCallback -= CMyCIMConnect_GetValueCallbackHandler;
            _CMyCimetrix.OnTerminalService -= CMyCIMConnect_TerminalServiceHandler;
            _CMyCimetrix.OnLog -= CMyCIMConnect_LogHandler;
            _CMyCimetrix.OnCommEnableDisableSwitchHandler -= CMyCIMConnect_OnCommEnableDisableSwitchHandler;
            _CMyCimetrix.OnCtrlOfflineOnlineSwitchHandler -= CMyCIMConnect_CtrlOfflineOnlineSwitchHandler;
            _CMyCimetrix.OnCtrlLocalRemoteSwitchHandler -= CMyCIMConnect_CtrlLocalRemoteSwitchHandler;
            _CMyCimetrix.OnInitializationComplete -= _CMyCimetrix_onInitializationComplete;
            _CMyCimetrix.OnInitializationComplete -= _CMyCimetrix_onInitializationComplete;

            try
            {
                if (_CMyCimetrix._CxClientClerk == null)
                    return;

                CIM.UnregisterSVs();
                CIM.UnregisterECs();

                // Unregister for CommandCalled callbacks
                Log("_CxClientClerk.UnregisterCommandHandler JobCreate");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("JobCreate");

                Log("_CxClientClerk.UnregisterCommandHandler RESUME");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("RESUME");

                Log("_CxClientClerk.UnregisterCommandHandler PAUSE");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("PAUSE");

                Log("_CxClientClerk.UnregisterCommandHandler START");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("START");

                Log("_CxClientClerk.UnregisterCommandHandler STOP");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("STOP");

                Log("_CxClientClerk.UnregisterCommandHandler ABORT");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("ABORT");


                Log("_CxClientClerk.UnregisterCommandHandler PPSELECT");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("PPSELECT");

                Log("_CxClientClerk.UnregisterCommandHandler ExecuteRemoteCommand");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("ExecuteRemoteCommand");

                Log("_CxClientClerk.UnregisterCommandHandler GetProducts");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetProducts");

                Log("_CxClientClerk.UnregisterCommandHandler SelectProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("SelectProduct");

                Log("_CxClientClerk.UnregisterCommandHandler DownloadProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("DownloadProduct");

                Log("_CxClientClerk.UnregisterCommandHandler UploadProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("UploadProduct");

                Log("_CxClientClerk.UnregisterCommandHandler SetTerminalMessage");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("SetTerminalMessage");

                Log("_CxClientClerk.UnregisterCommandHandler GetUsers");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetUsers");

                Log("_CxClientClerk.UnregisterCommandHandler GetLoggedInUsers");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetLoggedInUsers");

                Log("_CxClientClerk.UnregisterCommandHandler CreateLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("CreateLot");

                Log("_CxClientClerk.UnregisterCommandHandler GetLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetLot");

                Log("_CxClientClerk.UnregisterCommandHandler GetLots");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetLot");

                Log("_CxClientClerk.UnregisterCommandHandler UpdateLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("UpdateLot");

                Log("_CxClientClerk.UnregisterCommandHandler DeleteLot");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("DeleteLot");

                Log("_CxClientClerk.UnregisterCommandHandler RenameProduct");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("RenameProduct");

                Log("_CxClientClerk.UnregisterCommandHandler GetModuleProcessStates");
                _CMyCimetrix._CxClientClerk.UnregisterCommandHandler("GetModuleProcessStates");
            }
            catch (COMException ex)
            {
                Log(@"COMException unregistering CIMConnect handlers:\r\n" + ex.ToString());
            }
            finally
            {
                if (_CMyCimetrix._CxEMService != null)
                    _CMyCimetrix.Shutdown();

                // shutdown TCP/IP
                cs1.StopAndWait();
                cs2.StopAndWait();
                cs3.StopAndWait();
                cs4.StopAndWait();

                SQLight.Close();

                LogManager.StopAllAndWait();
            }
        }

        //----------------------------------------------------------------
        private void InitializeGUI()
        //----------------------------------------------------------------
        {
            // For the switches, just use Connection 1 
            Int64 switchValue = 0;

            // Communication Switch 
            _CMyCimetrix.GetVariableValue(1, "CommEnableSwitch", ref switchValue);
            DisplayCommSwitch((CMyCimetrix.CommEnableDisableEnum)switchValue);

            // Control Switch
            _CMyCimetrix.GetVariableValue(1, "CtrlOnlineSwitch", ref switchValue);
            DisplayOfflineOnlineSwitch((CMyCimetrix.ControlOfflineOnlineEnum)switchValue);

            // Remote Control Switch 
            _CMyCimetrix.GetVariableValue(1, "ControlStateSwitch", ref switchValue);
            DisplayLocalRemoteSwitch((CMyCimetrix.ControlLocalRemoteEnum)switchValue);

            // GEM Host Communication Status 
            _CMyCimetrix.GetVariableValue(1, "CommState", ref switchValue);
            DisplayGEMHostCommState(switchValue);

            // GEM Host Control Status 
            _CMyCimetrix.GetVariableValue(1, "CONTROLSTATE", ref switchValue);
            DisplayGEMHostControlState(switchValue);
        }

        //----------------------------------------------------------------
        private void Log(String message)
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Info(message);
        }

        //----------------------------------------------------------------
        private void LogSafe(string message)
        //----------------------------------------------------------------
        {
            CMyCIMConnect_LogHandler(null, new CMyCimetrix.LogArgs(message));
        }

        #region "GUI Display"

        //----------------------------------------------------------------
        private void DisplayCommSwitch(CMyCimetrix.CommEnableDisableEnum value)
        //----------------------------------------------------------------
        {
            switch (value)
            {
                case CMyCimetrix.CommEnableDisableEnum.Disable:
                    LogManager.DefaultLogger.Info("CommSwitch -> Disabled");
                    break;

                case CMyCimetrix.CommEnableDisableEnum.Enable:
                    LogManager.DefaultLogger.Info("CommSwitch -> Enabled");
                    break;
            }
        }

        //----------------------------------------------------------------
        private void DisplayOfflineOnlineSwitch(CMyCimetrix.ControlOfflineOnlineEnum value)
        //----------------------------------------------------------------
        {
            switch (value)
            {
                case CMyCimetrix.ControlOfflineOnlineEnum.Offline:
                    LogManager.DefaultLogger.Info("OnlineOffline -> Offline");
                    break;

                case CMyCimetrix.ControlOfflineOnlineEnum.Online:
                    LogManager.DefaultLogger.Info("OnlineOffline -> Online");
                    break;
            }
        }

        //----------------------------------------------------------------
        private void DisplayLocalRemoteSwitch(CMyCimetrix.ControlLocalRemoteEnum value)
        //----------------------------------------------------------------
        {
            switch (value)
            {
                case CMyCimetrix.ControlLocalRemoteEnum.Local:
                    LogManager.DefaultLogger.Info("LocalRemote -> Local");
                    break;

                case CMyCimetrix.ControlLocalRemoteEnum.Remote:
                    LogManager.DefaultLogger.Info("LocalRemote -> Remote");
                    break;
            }
        }

        //----------------------------------------------------------------
        private void DisplayGEMHostCommState(Int64 value)
        //----------------------------------------------------------------
        {
            EMSERVICELib.CommunicationState _communcationState = (EMSERVICELib.CommunicationState)value;
            switch (_communcationState)
            {
                case EMSERVICELib.CommunicationState.Disabled:
                    LogManager.DefaultLogger.Info("GEMHostCommState -> Disabled");
                    break;
                case EMSERVICELib.CommunicationState.Communicating:
                    LogManager.DefaultLogger.Info("GEMHostCommState -> Communicating");
                    break;
                case EMSERVICELib.CommunicationState.WaitCRAOrCRFromHost:
                    LogManager.DefaultLogger.Info("GEMHostCommState -> Not Communicating");
                    break;
                case EMSERVICELib.CommunicationState.WaitDelayOrCRFromHost:
                    LogManager.DefaultLogger.Info("GEMHostCommState -> Not Communicating");
                    break;
                default:
                    LogManager.DefaultLogger.Info("Unexpected value in DisplayGEMHostCommState = " + value);
                    break;
            }
        }

        //----------------------------------------------------------------
        private void DisplayGEMHostControlState(Int64 value)
        //----------------------------------------------------------------
        {
            EMSERVICELib.ControlState _controlState = (EMSERVICELib.ControlState)value;
            switch (_controlState)
            {
                case EMSERVICELib.ControlState.EqOffline:
                    LogManager.DefaultLogger.Info("GEMHostControlState -> Equipment Offline");
                    break;
                case EMSERVICELib.ControlState.AttemptOnline:
                    LogManager.DefaultLogger.Info("GEMHostControlState -> Attempt Online");
                    break;
                case EMSERVICELib.ControlState.HostOffline:
                    LogManager.DefaultLogger.Info("GEMHostControlState -> Host Offline");
                    break;
                case EMSERVICELib.ControlState.OnlineLocal:
                    LogManager.DefaultLogger.Info("GEMHostControlState -> Online Local");
                    break;
                case EMSERVICELib.ControlState.OnlineRemote:
                    LogManager.DefaultLogger.Info("GEMHostControlState -> Online Remote");
                    break;
                default:
                    LogManager.DefaultLogger.Error("Unexpected value in DisplayGEMHostControlState = " + value);
                    break;
            }
        }

        //----------------------------------------------------------------
        private void DisplayGEMProcessState(long value)
        //----------------------------------------------------------------
        {
            MyProcessState _MyProcessState;
            _MyProcessState = (MyProcessState)value;
            LogManager.DefaultLogger.Info("ProcessState -> " + _MyProcessState.ToString());
        }

        #endregion "GUI Display"

        #region "Event handlers"

        //----------------------------------------------------------------
        void _CMyCimetrix_onInitializationComplete(bool bSuccess, string errMessage)
        //----------------------------------------------------------------
        {
            if (bSuccess)
            {
                InitializeGUI();

                _ProcessingThreadStart = ProcessingThread;
                _ProcessingThread = new Thread(_ProcessingThreadStart);
                _ProcessingThreadExit = false;
                _ProcessingThread.Start();
            }
            else
            {
                LogManager.DefaultLogger.Info("TInitialization failed for the following reason:\r\n" + errMessage);
            }
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_SetECValueCallbackHandler(object sender, CMyCimetrix.ValueChangedCallbackArgs e)
        //----------------------------------------------------------------
        {
            lock (_lockValue)
            {
                LogManager.DefaultLogger.Info("CMyCIMConnect_SetECValueCallbackHandler variable: " + e.VariableName + " id: " + e.VariableID + " value:" + e.Value.ToString());

                CmdSeqID++;
                SeqID++;

                try
                {
                    string datatype = CIM.GetDataTypeFromId(e.VariableID);
                    string strValue = String.Empty;
                    int DataTypeID = 0;
                    string DataTypeStr = "NULL";

                    switch (datatype)
                    {
                        case "I1":
                            I1Value b = (I1Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = b.Value.ToString();
                            DataTypeID = 10;
                            DataTypeStr = "char";
                            break;

                        case "I2":
                            I2Value i2 = (I2Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = i2.Value.ToString();
                            DataTypeID = 8;
                            DataTypeStr = "short";
                            break;

                        case "I4":
                            I4Value i4 = (I4Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = i4.Value.ToString();
                            DataTypeID = 2;
                            DataTypeStr = "int";
                            break;

                        case "I8":
                            I8Value l = (I8Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = l.Value.ToString();
                            DataTypeID = 6;
                            DataTypeStr = "long long";
                            break;

                        case "U1":
                            U1Value ub = (U1Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = ub.Value.ToString();
                            DataTypeID = 9;
                            DataTypeStr = "unsigned char";
                            break;

                        case "U2":
                            U2Value us = (U2Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = us.Value.ToString();
                            DataTypeID = 7;
                            DataTypeStr = "unsigned short";
                            break;

                        case "U4":
                            U4Value ui4 = (U4Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = ui4.Value.ToString();
                            DataTypeID = 1;
                            DataTypeStr = "unsigned int";
                            break;

                        case "U8":
                            U8Value ul = (U8Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = ul.Value.ToString();
                            DataTypeID = 5;
                            DataTypeStr = "unsigned long long";
                            break;

                        case "F4":
                            F4Value f = (F4Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = f.Value.ToString();
                            DataTypeID = 13;
                            DataTypeStr = "float";
                            break;

                        case "F8":
                            F8Value d = (F8Value)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = d.Value.ToString();
                            DataTypeID = 11;
                            DataTypeStr = "double";
                            break;

                        case "A":
                            string s = CxValue.CreateFromByteBuffer((byte[])e.Value).ToString();
                            strValue = s;
                            DataTypeID = 15;
                            DataTypeStr = "string";
                            break;

                        case "Bo":
                            BoValue bo = (BoValue)CxValue.CreateFromByteBuffer((byte[])e.Value);
                            strValue = bo.ToString();
                            DataTypeID = 14;
                            DataTypeStr = "bool";
                            break;
                    }

                    using (var sw = new StringWriter())
                    {
                        using (var xw = XmlWriter.Create(sw))
                        {
                            xw.WriteStartElement("Cmd");

                            xw.WriteAttributeString("ID", "SetVariables");
                            xw.WriteAttributeString("EquipID", EquipID);
                            xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                            xw.WriteAttributeString("SeqID", SeqID.ToString());

                            xw.WriteStartElement("Variable");
                            xw.WriteAttributeString("ID", e.VariableID.ToString());
                            xw.WriteAttributeString("Name", e.VariableName);
                            xw.WriteAttributeString("Type", "CE");
                            xw.WriteAttributeString("UnitID", "0");
                            xw.WriteAttributeString("Unit", "NULL");
                            xw.WriteAttributeString("DataTypeID", DataTypeID.ToString());
                            xw.WriteAttributeString("DataType", DataTypeStr);
                            xw.WriteValue(strValue);
                            xw.WriteEndElement();

                            xw.WriteEndElement();
                        }

                        string txt = sw.ToString();
                        txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                        LogManager.DefaultLogger.Info("CMyCIMConnect_SetECValueCallbackHandler");

                        string Result = "no Result";
                        int Error = 0;
                        string sql = "";

                        try
                        {
                            // daten in db einfügen
                            SQLiteCommand cmd2 = conn.CreateCommand();
                            sql = "INSERT INTO Variables VALUES('SetVariables','" + EquipID.ToString() + "'," + CmdSeqID.ToString() + ",0," + SeqID.ToString() + ","
                                                    + e.VariableID.ToString() + ",'"
                                                    + e.VariableName + "','"
                                                    + "CE',"
                                                    + "0"
                                                    + ",'"
                                                    + "NULL"
                                                    + "',"
                                                    + DataTypeID
                                                    + ",'"
                                                    + DataTypeStr
                                                    + "','"
                                                    + "SECSHOST"
                                                    + "','"
                                                    + strValue
                                                    + "','"
                                                    + Result
                                                    + "',"
                                                    + Error
                                                    + ",'"
                                                    + DateTime.Now.ToString()
                                                    + "')";
                            cmd2.CommandText = sql;
                            cmd2.ExecuteNonQuery();
                            cmd2.Dispose();
                        }
                        catch (Exception ex)
                        {
                            LogManager.DefaultLogger.Info("CMyCIMConnect_SetECValueCallbackHandler -> DB Error:\r\n" + ex.ToString());
                            LogManager.DefaultLogger.Info("SQL:\r\n" + sql);
                        }

                        cs1.SendCommand(txt);
                    }
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Error("CMyCIMConnect_SetECValueCallbackHandler Exception:\r\n" + ex.Message);
                    return;
                }
            }
        }

        /// <summary>
        /// This callback allows CIMConnect to query the current value, so that the application
        /// does not have to continuously cache this value in CIMConnect. A "lock" is not
        /// used since the example values are type long. 
        /// RegisterGetValueToByteBufferHandler registers the variables for this callback. 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //----------------------------------------------------------------
        void CMyCIMConnect_GetValueCallbackHandler(object sender, CMyCimetrix.GetValueCallbackArgs e)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                CmdSeqID++;
                SeqID++;

                LogManager.DefaultLogger.Info("GetVariables -> GetValueCallbackHandler with variable: " + e.VariableName + " id: " + e.VariableID);

                string type = CIM.Get_DV_SV_EC_TypeFromId(e.VariableID);

                String request = String.Empty;

                if (type == "SV" || type == "EC")
                {
                    // Prepare to request the value from Equipment

                    using (var sw2 = new StringWriter())
                    {
                        using (var xw = XmlWriter.Create(sw2))
                        {
                            xw.WriteStartElement("Cmd");

                            xw.WriteAttributeString("ID", "GetVariables");
                            xw.WriteAttributeString("EquipID", EquipID);
                            xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                            xw.WriteAttributeString("SeqID", SeqID.ToString());

                            xw.WriteStartElement("Variable");
                            xw.WriteAttributeString("ID", e.VariableID.ToString());
                            xw.WriteAttributeString("Name", e.VariableName);
                            xw.WriteEndElement();

                            xw.WriteEndElement();
                        }

                        string txt = sw2.ToString();
                        txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                        string Result = "no Result";
                        int Error = 0;
                        string sql = "";

                        try
                        {
                            // daten in db einfügen
                            SQLiteCommand cmd2 = conn.CreateCommand();
                            sql = "INSERT INTO Variables VALUES('GetVariables','" + EquipID.ToString() + "'," + CmdSeqID.ToString() + ",0," + SeqID.ToString() + ","
                                                    + e.VariableID.ToString() + ",'"
                                                    + e.VariableName + "','"
                                                    + type + "',"
                                                    + "0"
                                                    + ",'"
                                                    + "NULL"
                                                    + "',"
                                                    + "0"
                                                    + ",'"
                                                    + "NULL"
                                                    + "','"
                                                    + "NULL"
                                                    + "','"
                                                    + "NULL"
                                                    + "','"
                                                    + Result
                                                    + "',"
                                                    + Error
                                                    + ",'"
                                                    + DateTime.Now.ToString()
                                                    + "')";

                            cmd2.CommandText = sql;
                            cmd2.ExecuteNonQuery();
                            cmd2.Dispose();
                        }
                        catch (Exception ex)
                        {
                            LogManager.DefaultLogger.Info("CMyCIMConnect_GetValueCallbackHandler -> DB Error:\r\n" + ex.ToString());
                            LogManager.DefaultLogger.Info("SQL:\r\n" + sql);
                        }

                        cs1.SendCommand(txt);
                        Thread.Sleep(300);
                    }
                }
                else if (type == "DV")
                {
                    LogManager.DefaultLogger.Info("Unexpected call for DV value");
                    throw new Exception("CMyCIMConnect_GetValueCallbackHandler -> Unexpected call for DV value");

                }
                else
                {
                    LogManager.DefaultLogger.Info("SECS Variable type not found in Hash Table");
                    throw new Exception("CMyCIMConnect_GetValueCallbackHandler -> SECS Variable type not found in Hash Table");
                }

                LogManager.DefaultLogger.Info("Before checking for received data");
                int count = 0;

                try
                {
                    // Hopefully data already available
                    SQLiteCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT count(*) FROM Variables WHERE VARIABLEID = " + e.VariableID.ToString();
                    SQLiteDataReader rdr = cmd.ExecuteReader();
                    rdr.Read();
                    count = rdr.GetInt32(0);
                    cmd.Dispose();
                    rdr.Dispose();

                    LogManager.DefaultLogger.Info("After checking for received data with count = " + count.ToString());
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Info("Exception on SELECT count(*) FROM Variables:\r\n" + ex.ToString());
                }

                Stopwatch sw = new Stopwatch();
                sw.Start();

                try
                {
                    while (true)
                    {
                        Thread.Sleep(100);

                        SQLiteCommand cmd = conn.CreateCommand();
                        cmd.CommandText = "SELECT count(*) FROM Variables WHERE Action = 'GetVariablesResponse' AND VARIABLEID = " + e.VariableID.ToString();
                        SQLiteDataReader rdr = cmd.ExecuteReader();
                        rdr.Read();
                        count = rdr.GetInt32(0);
                        cmd.Dispose();
                        rdr.Dispose();

                        if (count > 0)
                        {
                            break;
                        }

                        if (sw.ElapsedMilliseconds > 1500) throw new TimeoutException();
                    }
                }
                catch (TimeoutException)
                {
                    LogManager.DefaultLogger.Error("Request GetVariables did not receive the value on time ->  used a default val instead (0)");
                }

                string VALUE = String.Empty;
                string datatype = CIM.GetDataTypeFromId(e.VariableID);

                LogManager.DefaultLogger.Info("Found datatype = " + datatype);

                if (count > 0)
                {
                    // Now get the Data
                    SQLiteCommand cmd = conn.CreateCommand();
                    cmd.CommandText = "SELECT VALUE FROM Variables WHERE Action = 'GetVariablesResponse' AND VARIABLEID = " + e.VariableID.ToString() + " ORDER BY SeqID DESC";
                    SQLiteDataReader rdr = cmd.ExecuteReader();
                    rdr.Read();
                    VALUE = rdr.GetString(0);
                    cmd.Dispose();
                    rdr.Dispose();

                    if (VALUE == "NULL")
                    {
                        VALUE = "0";
                    }
                }
                else
                {
                    VALUE = "0";
                }

                try
                {
                    switch (datatype)
                    {
                        case "I1":
                            sbyte b = sbyte.Parse(VALUE);
                            e.VariableValue = new I1Value(b);
                            break;

                        case "I2":
                            short s = short.Parse(VALUE);
                            e.VariableValue = new I2Value(s);
                            break;

                        case "I4":
                            int i = int.Parse(VALUE);
                            e.VariableValue = new I4Value(i);
                            break;

                        case "I8":
                            long l = long.Parse(VALUE);
                            e.VariableValue = new I8Value(l);
                            break;

                        case "U1":
                            byte ub = byte.Parse(VALUE);
                            e.VariableValue = new U1Value(ub);
                            break;

                        case "U2":
                            ushort us = ushort.Parse(VALUE);
                            e.VariableValue = new U2Value(us);
                            break;

                        case "U4":
                            uint ui = uint.Parse(VALUE);
                            e.VariableValue = new U4Value(ui);
                            break;

                        case "U8":
                            ulong ul = ulong.Parse(VALUE);
                            e.VariableValue = new U8Value(ul);
                            break;

                        case "F4":
                            float f = float.Parse(VALUE);
                            e.VariableValue = new F4Value(f);
                            break;

                        case "F8":
                            double d = double.Parse(VALUE);
                            e.VariableValue = new F8Value(d);
                            break;

                        case "A":
                            e.VariableValue = new AValue(VALUE);
                            break;

                        case "Bo":
                            bool bo = bool.Parse(VALUE);
                            e.VariableValue = new BoValue(bo);
                            break;

                        default:
                            int i2 = int.Parse(VALUE);
                            e.VariableValue = new I4Value(i2);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Error("Exception:");
                    LogManager.DefaultLogger.Error(ex.Message);
                }
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_GEMStateChangeHandler(object sender, CMyCimetrix.GEMStateChangeArgs e)
        //----------------------------------------------------------------
        {
            // In this example, the connection 1 states only are displayed.
            // Process state machine uses connection 0
            if (e.Connection > 1)
                return;

            switch (e.StateMachine)
            {
                case StateMachine.smCommunications:
                    DisplayGEMHostCommState(e.State);
                    break;

                case StateMachine.smControl:
                    DisplayGEMHostControlState(e.State);
                    break;

                case StateMachine.smProcess:
                    DisplayGEMProcessState(e.State);
                    break;
            }
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_OnCommEnableDisableSwitchHandler(object sender, CMyCimetrix.CommEnableDisableSwitchArgs e)
        //----------------------------------------------------------------
        {
            if (e.ConnectionId != 1)
                return;

            DisplayCommSwitch(e.Setting);
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_CtrlOfflineOnlineSwitchHandler(object sender, CMyCimetrix.ControlOfflineOnlineSwitchArgs e)
        //----------------------------------------------------------------
        {
            if (e.ConnectionId != 1)
                return;

            DisplayOfflineOnlineSwitch(e.Setting);
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_CtrlLocalRemoteSwitchHandler(object sender, CMyCimetrix.ControlLocalRemoteSwitchArgs e)
        //----------------------------------------------------------------
        {
            if (e.ConnectionId != 1)
                return;

            DisplayLocalRemoteSwitch(e.Setting);
        }

        const int _ControlConnection = 1;
        //----------------------------------------------------------------
        void CMyCIMConnect_RemoteCommandHandler(object sender, CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            LogSafe("CMyCIMConnect_RemoteCommandHandler " + e.Command);

            // Only let one connection be the 'controlling' connection. 
            if (e.Connection != _ControlConnection)
            {
                LogSafe("Remote Command " + e.Command + " rejected because connection " + e.Connection + " does not support remote commands.");
                return;
            }

            // Only allow remote commands while in the Remote Control State
            if (_CMyCimetrix._ControlState[e.Connection - 1] != ControlState.OnlineRemote)
            {
                LogSafe("Remote Command " + e.Command + " rejected because the Control State is not remote");
                e.CommandResult = CommandResults.cmdRejected;
                return;
            }

            // Check Parameters
            int argumentCount;
            if (e.ArgumentValues == null)
                argumentCount = 0;
            else
                e.ArgumentNames.ItemCount(0, out argumentCount);

            switch (e.Command)
            {
                case "START":
                    if (argumentCount != 2)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because it requires 2 arguments (Modul ID + Name)");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }
                    HandleStart(e);
                    break;

                case "STOP":
                    if (argumentCount != 2)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because it requires 2 arguments (Modul ID + Name)");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }
                    HandleStop(e);
                    break;

                case "PAUSE":
                    if (argumentCount != 2)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because it requires 2 arguments (Modul ID + Name)");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }
                    HandlePause(e);
                    break;

                case "RESUME":
                    if (argumentCount != 2)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because it requires 2 arguments (Modul ID + Name)");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }
                    HandleResume(e);
                    break;

                case "ABORT":
                    if (argumentCount != 2)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because it requires 2 arguments (Modul ID + Name)");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }
                    HandleAbort(e);
                    break;

                case "PPSELECT":
                    // Validate the number of requirements 
                    if (argumentCount != 1)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because it supports one and only one required argument, PPID. ");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }

                    // Validate that the argument is PPID 
                    string argumentName;
                    e.ArgumentNames.GetValueAscii(0, 1, out argumentName);
                    if (argumentName != "PPID")
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because only PPID argument is allowed. ");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }

                    // Validate the argument value type ( must be ASCII) 
                    VALUELib.ValueType valueType;
                    e.ArgumentValues.GetDataType(0, 1, out valueType);

                    if (valueType != VALUELib.ValueType.A)
                    {
                        LogSafe("Remote Command " + e.Command + " rejected because the value for PPID must be ASCII. ");
                        e.CommandResult = CommandResults.cmdParamInvalid;
                        return;
                    }
                    break;
            }

            switch (e.Command)
            {
                case "START":
                case "RESUME":
                    _EventSTART.Set();
                    e.CommandResult = CommandResults.cmdPerformed;
                    return;

                case "STOP":
                case "PAUSE":
                case "ABORT":
                    _EventSTOP.Set();
                    e.CommandResult = CommandResults.cmdPerformed;
                    return;

                case "PPSELECT":
                    string ppid;
                    e.ArgumentValues.GetValueAscii(0, 1, out ppid);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "JobCreate":
                    if (HandleJobCreate(e) == false)
                    {
                        e.CommandResult = CommandResults.cmdCannotPerform;
                    }
                    else
                    {
                        e.CommandResult = CommandResults.cmdPerformed;
                    }
                    break;

                case "ExecuteRemoteCommand":
                    HandleExecuteRemoteCommand(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "GetProducts":
                    HandleGetProducts(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "SelectProduct":
                    HandleSelectProduct(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "DownloadProduct":
                    HandleDownloadProduct(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "UploadProduct":
                    HandleUploadProduct(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "SetTerminalMessage":
                    HandleSetTerminalMessage(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "GetUsers":
                    HandleGetUsers(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "GetLoggedInUsers":
                    HandleGetLoggedInUsers(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "CreateLot":
                    HandleCreateLot(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "GetLot":
                    HandleGetLot(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "GetLots":
                    HandleGetLots(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "UpdateLot":
                    HandleUpdateLot(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "DeleteLot":
                    HandleDeleteLot(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "RenameProduct":
                    HandleRenameProduct(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;

                case "GetModuleProcessStates":
                    HandleGetModuleProcessStates(e);
                    e.CommandResult = CommandResults.cmdPerformed;
                    break;
            }
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_LogHandler(object sender, CMyCimetrix.LogArgs e)
        //----------------------------------------------------------------
        {
            Log(e.TextMessage);
        }

        //----------------------------------------------------------------
        void CMyCIMConnect_TerminalServiceHandler(object sender, CMyCimetrix.TerminalServiceArgs e)
        //----------------------------------------------------------------
        {
        }

        #endregion "Event handlers"

        //----------------------------------------------------------------
        public void UpdateProcessState(MyProcessState state)
        //----------------------------------------------------------------
        {
            try
            {
                string stateName;
                string eventName;

                switch (state)
                {
                    case MyProcessState.SETUP:
                        stateName = "SETUP";
                        eventName = "ProcessStateToSETUP";
                        break;

                    case MyProcessState.EXECUTING:
                        stateName = "EXECUTING";
                        eventName = "ProcessStateToEXECUTING";
                        _CMyCimetrix.ProcessingStarted();

                        break;

                    case MyProcessState.IDLE:
                        stateName = "IDLE";
                        eventName = "ProcessStateToIDLE";
                        _CMyCimetrix.ProcessingCompleted();
                        break;

                    case MyProcessState.PAUSE:
                        stateName = "PAUSE";
                        eventName = "ProcessStateTPAUSE";
                        break;

                    default:
                        return;
                }

                LogSafe("_CxClientClerk.ProcessingStateChange " + stateName);
                _CMyCimetrix._CxClientClerk.ProcessingStateChange((int)state, ref stateName, -1, ref eventName);
                _MyProcessState = state;
            }
            catch (Exception e)
            {
                _CMyCimetrix.HandleException("Exception calling ::ProcessingStateChange " + state, e);
            }
        }

        //----------------------------------------------------------------
        public void InitializationMTA()
        //----------------------------------------------------------------
        {
            try
            {
                LogSafe("Initializing Cimetrix library in an MTA thread. ");
                _CMyCimetrix = new CMyCimetrix();
                _CMyCimetrix.OnLog += CMyCIMConnect_LogHandler;
                _CMyCimetrix.OnTerminalService += CMyCIMConnect_TerminalServiceHandler;
                _CMyCimetrix.OnRemoteCommand += CMyCIMConnect_RemoteCommandHandler;
                _CMyCimetrix.OnGEMStateChange += CMyCIMConnect_GEMStateChangeHandler;
                _CMyCimetrix.OnGetValueCallback += CMyCIMConnect_GetValueCallbackHandler;
                _CMyCimetrix.OnCommEnableDisableSwitchHandler += CMyCIMConnect_OnCommEnableDisableSwitchHandler;
                _CMyCimetrix.OnCtrlOfflineOnlineSwitchHandler += CMyCIMConnect_CtrlOfflineOnlineSwitchHandler;
                _CMyCimetrix.OnCtrlLocalRemoteSwitchHandler += CMyCIMConnect_CtrlLocalRemoteSwitchHandler;
                _CMyCimetrix.OnInitializationComplete += _CMyCimetrix_onInitializationComplete;
                _CMyCimetrix.OnValueChangedCallback += CMyCIMConnect_SetECValueCallbackHandler;


                _CMyCimetrix.Initialize("Biometric.epj");

                UpdateProcessState(MyProcessState.SETUP);

                _EventSTART = new ManualResetEvent(false);
                _EventSTOP = new ManualResetEvent(false);

                CIM.INitializeCIM(_CMyCimetrix, conn);

                CIM.DefineDynamicGEMItems();

                // Register for GEM Remote Commands
                LogSafe("_CxClientClerk.RegisterCommandHandler JobCreate");
                var commandDescription = "Create the job.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("JobCreate", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler PPSELECT");
                commandDescription = "Select Recipe for execution with a PPID argument.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("PPSELECT", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler START");
                commandDescription = "Start processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("START", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler STOP");
                commandDescription = "Stop processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("STOP", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler PAUSE");
                commandDescription = "PAUSE processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("PAUSE", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler RESUME");
                commandDescription = "RESUME processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("RESUME", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler ABORT");
                commandDescription = "Start processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("ABORT", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler ExecuteRemoteCommand");
                commandDescription = "ExecuteRemoteCommand processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("ExecuteRemoteCommand", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler GetProducts");
                commandDescription = "GetProducts processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("GetProducts", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler SelectProduct");
                commandDescription = "SelectProduct processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("SelectProduct", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler DownloadProduct");
                commandDescription = "DownloadProduct processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("DownloadProduct", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler UploadProduct");
                commandDescription = "UploadProduct processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("UploadProduct", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler SetTerminalMessage");
                commandDescription = "SetTerminalMessage processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("SetTerminalMessage", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler GetUsers");
                commandDescription = "GetUsers processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("GetUsers", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler GetLoggedInUsers");
                commandDescription = "GetLoggedInUsers processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("GetLoggedInUsers", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler CreateLot");
                commandDescription = "CreateLot processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("CreateLot", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler GetLot");
                commandDescription = "GetLot processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("GetLot", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler GetLots");
                commandDescription = "GetLots processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("GetLots", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler UpdateLot");
                commandDescription = "UpdateLot processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("UpdateLot", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler DeleteLot");
                commandDescription = "DeleteLot processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("DeleteLot", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler RenameProduct");
                commandDescription = "RenameProduct processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("RenameProduct", ref commandDescription, _CMyCimetrix);

                LogSafe("_CxClientClerk.RegisterCommandHandler GetModuleProcessStates");
                commandDescription = "GetModuleProcessStates processing.";
                _CMyCimetrix._CxClientClerk.RegisterCommandHandler("GetModuleProcessStates", ref commandDescription, _CMyCimetrix);

                _CMyCimetrix.InitializeFinal();
                _CMyCimetrix.GEMStateControlStateRemote(_PrimaryConnection, true);
            }
            catch (Exception ex)
            {
                LogSafe("Initialization Failed on: " + ex.Message);
            }
        }

        //----------------------------------------------------------------
        public void ProcessingThread()
        //----------------------------------------------------------------
        {
            // Process loop
            while (!_ProcessingThreadExit)
            {
                _EventSTART.Reset();
                UpdateProcessState(MyProcessState.IDLE);

                _EventSTART.WaitOne();

                if (_ProcessingThreadExit)
                {
                    break;
                }
                _EventSTOP.Reset();

                // Continue executing until stopped. 
                UpdateProcessState(MyProcessState.EXECUTING);


                while (!_EventSTOP.WaitOne(500))
                {
                    if (_ProcessingThreadExit)
                    {
                        break;
                    }
                }

            }
        }

        //----------------------------------------------------------------
        private String ReadSetting(string key)
        //----------------------------------------------------------------
        {
            var appSettings = ConfigurationManager.AppSettings;

            try
            {
                string result = appSettings[key] ?? "Not Found";
                LogManager.DefaultLogger.Info("Parameter read: " + key + " -> " + result);
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                LogManager.DefaultLogger.Error("Error reading app setting: " + key);
                return "nil";
            }
        }

        //----------------------------------------------------------------
        public bool HandleJobCreate(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string ppid, txt, TESTORDER, JOBID, QUANTITY, ASSEMBLYLOTID;
            int Handle1 = 0;
            int count = 0;

            e.ArgumentValues.GetValueAscii(0, 2, out JOBID);
            e.ArgumentValues.GetValueAscii(0, 3, out QUANTITY);
            e.ArgumentValues.GetValueAscii(0, 4, out ppid);
            e.ArgumentValues.GetValueAscii(0, 5, out ASSEMBLYLOTID);

            e.ArgumentValues.ListAt(0, 6, out Handle1);

            e.ArgumentValues.ItemCount(Handle1, out count);

            e.ArgumentValues.GetValueAscii(Handle1, 1, out txt);
            TESTORDER = txt + "\r\n";

            try
            {
                for (int i = 2; i <= count; i++)
                {
                    e.ArgumentValues.GetValueAscii(Handle1, i, out txt);

                    if (string.IsNullOrEmpty(txt))
                    {
                        txt = string.Empty;
                    }

                    TESTORDER += txt + "\r\n";
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleJobCreate Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }

            LogManager.DefaultLogger.Info("HandleJobCreate:");
            LogManager.DefaultLogger.Info("JOBID: " + JOBID);
            LogManager.DefaultLogger.Info("QUANTITY: " + QUANTITY);
            LogManager.DefaultLogger.Info("ppid: " + ppid);
            LogManager.DefaultLogger.Info("ASSEMBLYLOTID: " + ASSEMBLYLOTID);
            LogManager.DefaultLogger.Info("TESTORDER:\r\n" + TESTORDER);

            if (CreateLotFromJob(JOBID, QUANTITY, ppid, ASSEMBLYLOTID, TESTORDER) == false)
            {
                return false;

            }
            else
            {
                return true;
            }
        }

        //----------------------------------------------------------------
        public void HandleExecuteRemoteCommand(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string modulid, modulname, id, name;

            e.ArgumentValues.GetValueAscii(0, 1, out modulid);
            e.ArgumentValues.GetValueAscii(0, 2, out modulname);
            e.ArgumentValues.GetValueAscii(0, 3, out id);
            e.ArgumentValues.GetValueAscii(0, 4, out name);

            LogManager.DefaultLogger.Info("HandleExecuteRemoteCommand received: " + modulid + "/" + modulname + "/" + id + "/" + name);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "ExecuteRemoteCommand");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("RemoteCmd");
                    xw.WriteAttributeString("ID", id);
                    xw.WriteAttributeString("Name", name);
                    xw.WriteAttributeString("ModuleID", modulid);
                    xw.WriteAttributeString("ModuleName", modulname);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send ExecuteRemoteCommand");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleGetProducts(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Info("HandleGetProducts received: ");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "GetProducts");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteValue("0");

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send GetProducts");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleSelectProduct(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string name;

            e.ArgumentValues.GetValueAscii(0, 1, out name);

            LogManager.DefaultLogger.Info("HandleSelectProduct received: " + name);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "SelectProduct");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Product");
                    xw.WriteValue(name);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send SelectProduct");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleDownloadProduct(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string name;

            e.ArgumentValues.GetValueAscii(0, 1, out name);

            LogManager.DefaultLogger.Info("HandleDownloadProduct received: " + name);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "DownloadProduct");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Product");
                    xw.WriteAttributeString("Name", name);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send DownloadProduct");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleUploadProduct(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string name, xmldata;

            e.ArgumentValues.GetValueAscii(0, 1, out name);
            e.ArgumentValues.GetValueAscii(0, 2, out xmldata);

            LogManager.DefaultLogger.Info("HandleUploadProduct received: " + name + "\r\n" + xmldata);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "UploadProduct");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Product");
                    xw.WriteAttributeString("Name", name);
                    xw.WriteValue(xmldata);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send UploadProduct");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleRenameProduct(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string Old, New;

            e.ArgumentValues.GetValueAscii(0, 1, out Old);
            e.ArgumentValues.GetValueAscii(0, 2, out New);

            LogManager.DefaultLogger.Info("HandleRenameProduct received: " + Old + "/" + New);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "RenameProduct");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Product");
                    xw.WriteAttributeString("Name", Old);
                    xw.WriteValue(New);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send RenameProduct");
                cs1.SendCommand(txt);
            }
        }


        //----------------------------------------------------------------
        public void HandleSetTerminalMessage(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string Titel, Text, Type, Options;

            e.ArgumentValues.GetValueAscii(0, 1, out Titel);
            e.ArgumentValues.GetValueAscii(0, 2, out Text);
            e.ArgumentValues.GetValueAscii(0, 3, out Type);
            e.ArgumentValues.GetValueAscii(0, 4, out Options);

            LogManager.DefaultLogger.Info("HandleSetTerminalMessage received: " + Titel + "/" + Text + "/" + Type + "/" + Options);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "SetTerminalMessage");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("TerminalMessage");

                    xw.WriteStartElement("MessageTitle");
                    xw.WriteValue(Titel);
                    xw.WriteEndElement();

                    xw.WriteStartElement("MessageText");
                    xw.WriteValue(Text);
                    xw.WriteEndElement();

                    xw.WriteStartElement("MessageType");
                    xw.WriteValue(Type);
                    xw.WriteEndElement();

                    xw.WriteStartElement("MessageCode");
                    xw.WriteValue("1");
                    xw.WriteEndElement();

                    xw.WriteStartElement("MessageOptions");
                    xw.WriteValue(Options);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Stop");
                    xw.WriteValue("false");
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send SetTerminalMessage");
                cs1.SendCommand(txt);
            }
        }


        //----------------------------------------------------------------
        public void HandleGetModuleProcessStates(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string ModulName1, ModulName2 = "no", ModulID1, ModulID2 = "0";
            int itemcount = 0;

            e.ArgumentValues.ItemCount(0, out itemcount);

            e.ArgumentValues.GetValueAscii(0, 1, out ModulName1);
            e.ArgumentValues.GetValueAscii(0, 2, out ModulID1);

            if (itemcount > 2)
            {
                e.ArgumentValues.GetValueAscii(0, 3, out ModulName2);
                e.ArgumentValues.GetValueAscii(0, 4, out ModulID2);
            }

            LogManager.DefaultLogger.Info("HandleGetModuleProcessStates received: " + ModulName1 + "/" + ModulID1 + "/" + ModulName2 + "/" + ModulID2);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "GetModuleProcessStates");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("ModuleState");
                    xw.WriteAttributeString("ID", ModulID1);
                    xw.WriteAttributeString("Name", ModulName1);
                    xw.WriteEndElement();

                    if (itemcount > 2)
                    {
                        xw.WriteStartElement("ModuleState");
                        xw.WriteAttributeString("ID", ModulID2);
                        xw.WriteAttributeString("Name", ModulName2);
                        xw.WriteEndElement();
                    }

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send SetTerminalMessage");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleGetUsers(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Info("HandleGetUsers received: ");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "GetUsers");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteValue("0");
                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send GetUsers");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleGetLoggedInUsers(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            LogManager.DefaultLogger.Info("HandleGetLoggedInUsers received: ");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "GetLoggedInUsers");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteValue("0");
                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send GetLoggedInUsers");
                cs1.SendCommand(txt);
            }
        }


        //----------------------------------------------------------------
        public void HandleCreateLot(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string Name, Count, Product, CustomName, CustomData;

            e.ArgumentValues.GetValueAscii(0, 1, out Name);
            e.ArgumentValues.GetValueAscii(0, 2, out Count);
            e.ArgumentValues.GetValueAscii(0, 3, out Product);
            e.ArgumentValues.GetValueAscii(0, 4, out CustomName);
            e.ArgumentValues.GetValueAscii(0, 5, out CustomData);

            if (string.IsNullOrEmpty(CustomName))
            {
                CustomName = "---";

            }

            if (string.IsNullOrEmpty(CustomData))
            {
                CustomData = "---";

            }

            LogManager.DefaultLogger.Info("HandleCreateLot received: " + Name + "/" + Count + "/" + Product);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "CreateLot");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Lot");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(Name);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Count");
                    xw.WriteValue(Count);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Product");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(Product);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteStartElement("CustomDataList");

                    xw.WriteStartElement("CustomData");
                    xw.WriteAttributeString("Name", CustomName);
                    xw.WriteValue(CustomData);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send CreateLot");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public bool CreateLotFromJob(string Name, string Count, string Product, string CustomName, string CustomData)
        //----------------------------------------------------------------
        {
            if (string.IsNullOrEmpty(CustomName))
            {
                CustomName = "---";

            }

            if (string.IsNullOrEmpty(CustomData))
            {
                CustomData = "---";

            }

            LogManager.DefaultLogger.Info("HandleCreateLot received: " + Name + "/" + Count + "/" + Product);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "CreateLot");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Lot");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(Name);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Count");
                    xw.WriteValue(Count);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Product");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(Product);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteStartElement("CustomDataList");

                    xw.WriteStartElement("CustomData");
                    xw.WriteAttributeString("Name", CustomName);
                    xw.WriteValue(CustomData);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send CreateLot");

                if (cs1.SendCommand(txt) == false)
                    return false;
            }

            return true;
        }


        //----------------------------------------------------------------
        public void HandleUpdateLot(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string Name, Count, Product;

            e.ArgumentValues.GetValueAscii(0, 1, out Name);
            e.ArgumentValues.GetValueAscii(0, 2, out Count);
            e.ArgumentValues.GetValueAscii(0, 3, out Product);

            LogManager.DefaultLogger.Info("HandleUpdateLot received: " + Name + "/" + Count + "/" + Product);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "UpdateLot");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Lot");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(Name);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Count");
                    xw.WriteValue(Count);
                    xw.WriteEndElement();

                    xw.WriteStartElement("Product");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(Product);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send UpdateLot");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleGetLot(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string name;

            e.ArgumentValues.GetValueAscii(0, 1, out name);

            LogManager.DefaultLogger.Info("HandleGetLot received: " + name);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "GetLot");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Lot");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(name);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send GetLot");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleGetLots(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {

            LogManager.DefaultLogger.Info("HandleGetLots received:");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "GetLots");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteValue(" ");
                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send GetLots");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleDeleteLot(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string name;

            e.ArgumentValues.GetValueAscii(0, 1, out name);

            LogManager.DefaultLogger.Info("HandleDeleteLot received: " + name);

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "DeleteLot");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("Lot");

                    xw.WriteStartElement("Name");
                    xw.WriteValue(name);
                    xw.WriteEndElement();

                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                LogManager.DefaultLogger.Info("Send DeleteLot");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleStart(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string modulid, modulname;

            if (e == null)
            {
                modulid = "99";
                modulname = "Testmodul";
            }
            else
            {
                e.ArgumentValues.GetValueAscii(0, 1, out modulid);
                e.ArgumentValues.GetValueAscii(0, 2, out modulname);
            }

            LogManager.DefaultLogger.Info("HandleStart received");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "ExecuteRemoteCommand");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("RemoteCmd");
                    xw.WriteAttributeString("ID", "1");
                    xw.WriteAttributeString("Name", "Start");
                    xw.WriteAttributeString("ModuleID", modulid);
                    xw.WriteAttributeString("ModuleName", modulname);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleStop(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string modulid, modulname;

            e.ArgumentValues.GetValueAscii(0, 1, out modulid);
            e.ArgumentValues.GetValueAscii(0, 2, out modulname);

            LogManager.DefaultLogger.Info("HandleStop received ");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "ExecuteRemoteCommand");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("RemoteCmd");
                    xw.WriteAttributeString("ID", "2");
                    xw.WriteAttributeString("Name", "Stop");
                    xw.WriteAttributeString("ModuleID", modulid);
                    xw.WriteAttributeString("ModuleName", modulname);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandlePause(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string modulid, modulname;

            e.ArgumentValues.GetValueAscii(0, 1, out modulid);
            e.ArgumentValues.GetValueAscii(0, 2, out modulname);

            LogManager.DefaultLogger.Info("HandlePause received");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "ExecuteRemoteCommand");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("RemoteCmd");
                    xw.WriteAttributeString("ID", "3");
                    xw.WriteAttributeString("Name", "Pause");
                    xw.WriteAttributeString("ModuleID", modulid);
                    xw.WriteAttributeString("ModuleName", modulname);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleResume(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string modulid, modulname;

            e.ArgumentValues.GetValueAscii(0, 1, out modulid);
            e.ArgumentValues.GetValueAscii(0, 2, out modulname);

            LogManager.DefaultLogger.Info("HandleResume received ");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "ExecuteRemoteCommand");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("RemoteCmd");
                    xw.WriteAttributeString("ID", "4");
                    xw.WriteAttributeString("Name", "Resume");
                    xw.WriteAttributeString("ModuleID", modulid);
                    xw.WriteAttributeString("ModuleName", modulname);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                cs1.SendCommand(txt);
            }
        }

        //----------------------------------------------------------------
        public void HandleAbort(CMyCimetrix.RemoteCommandArgs e)
        //----------------------------------------------------------------
        {
            string modulid, modulname;

            e.ArgumentValues.GetValueAscii(0, 1, out modulid);
            e.ArgumentValues.GetValueAscii(0, 2, out modulname);


            LogManager.DefaultLogger.Info("HandleAbort received");

            CmdSeqID++;
            SeqID++;

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    xw.WriteStartElement("Cmd");

                    xw.WriteAttributeString("ID", "ExecuteRemoteCommand");
                    xw.WriteAttributeString("EquipID", EquipID);
                    xw.WriteAttributeString("CmdSeqID", CmdSeqID.ToString());
                    xw.WriteAttributeString("SeqID", SeqID.ToString());

                    xw.WriteStartElement("RemoteCmd");
                    xw.WriteAttributeString("ID", "5");
                    xw.WriteAttributeString("Name", "Abort");
                    xw.WriteAttributeString("ModuleID", modulid);
                    xw.WriteAttributeString("ModuleName", modulname);
                    xw.WriteEndElement();

                    xw.WriteEndElement();
                }

                string txt = sw.ToString();
                txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                cs1.SendCommand(txt);
            }
        }
    }

    //----------------------------------------------------------------
    internal static class ConfigureService
    //----------------------------------------------------------------
    {
        internal static void Configure()
        {
            HostFactory.Run(configure =>
            {
                configure.Service<WinService>(service =>
                {
                    service.ConstructUsing(s => new WinService());
                    service.WhenStarted(s => s.Start());
                    service.WhenStopped(s => s.Stop());
                    service.WhenShutdown(s => s.Shutdown());
                });

                configure.RunAsLocalSystem();
                configure.StartAutomatically();
                configure.EnableShutdown();
                configure.SetServiceName("Biometric");
                configure.SetDisplayName("Biometric");
                configure.SetDescription("Service for Biometric machine");
            });
        }
    }

}
