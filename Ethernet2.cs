using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using System.Threading;
using System.Data.SQLite;
using System.Configuration;
using System.Xml.Linq;
using Cimetrix.Value;
using System.IO;
using System.Xml;
using System.Linq;
using System.Globalization;
using System.Net;
using System.Runtime.Remoting.Contexts;
using System.Security.AccessControl;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;
using System.Reflection;
using VALUELib;

namespace BiometricServer
{
    public class CS2
    {
        private readonly object lockMessageSend = new object();
        private ConcurrentQueue<string> S_Request = new ConcurrentQueue<string>();
        private ConcurrentQueue<string> Ack_Answer = new ConcurrentQueue<string>();
        private Thread WorkerThread;

        private Mutex mutex;

        public bool IsStopped { get; private set; } = false;
        private SQLiteConnection conn;

        private TcpListener tcpListener = null;

        //----------------------------------------------------------------
        public void Init(SQLiteConnection connnection, Mutex ethernetMutex)
        //----------------------------------------------------------------
        {
            conn = connnection;
            mutex = ethernetMutex;
        }

        //=================================================================
        // TCPIP Server part BEGIN

        //----------------------------------------------------------------
        public void StopListener()
        //----------------------------------------------------------------
        {
            // TCP/IP SERVER CHANNEL = Event processing from DB-Matik

            if (tcpListener != null)
            {
                LogManager.DefaultLogger.Info("TCPIP -> Server2] stopping Listener 2");
                tcpListener.Stop();
            }
        }

        //----------------------------------------------------------------
        public async Task StartListener()
        //----------------------------------------------------------------
        {
            // TCP/IP SERVER CHANNEL = Event processing from DB-Matik

            WorkerThread = new Thread(new ThreadStart(Worker));
            WorkerThread.Start();

            int port = int.Parse(ReadSetting("SERVER_PORT2"));

            LogManager.DefaultLogger.Info("TCPIP -> Server2] starting Listener on Port: " + port.ToString());

            tcpListener = new TcpListener(IPAddress.Any, port);
            
            try
            {
                tcpListener.Start();
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("TCPIP -> Server2]  Exception tcpListener.Start():\r\n" + ex.ToString());
            }

            while (!IsStopped)
            {
                try
                {
                    var tcpClient = await tcpListener.AcceptTcpClientAsync();
                    LogManager.DefaultLogger.Info("TCPIP -> Server2] DB-Matik Client has connected");

                    var task = StartHandleConnectionAsync(tcpClient);

                    // if already faulted, re-throw any error on the calling context
                    if (task.IsFaulted)
                    {
                        LogManager.DefaultLogger.Error("TCPIP -> Server2] if (task.IsFaulted) entered");
                        await task;
                    }
                }
                catch(Exception)
                {
                    //LogManager.DefaultLogger.Info("TCPIP -> Server2] Exception2:\r\n" + ex.ToString());
                }
            }
        }

        //----------------------------------------------------------------
        private async Task StartHandleConnectionAsync(TcpClient tcpClient)
        //----------------------------------------------------------------
        {
            // TCP/IP SERVER CHANNEL = Event processing from DB-Matik

            // start the new connection task
            var connectionTask = HandleConnectionAsync(tcpClient, S_Request, Ack_Answer);

            try
            {
                await connectionTask;
                // we may be on another thread after "await"
            }
            catch (Exception ex)
            {
                // log the error
                LogManager.DefaultLogger.Error("TCPIP -> Server2] Exception1:\r\n" + ex.ToString());

                // stop worker 
                IsStopped = true;
            }
            finally
            {
                tcpClient.Close();
                tcpListener.Stop();
                IsStopped = false;

                Thread.Sleep(2000);

                // start again
                _ = StartListener();
            }
        }

        //----------------------------------------------------------------
        private static Task HandleConnectionAsync(TcpClient tcpClient, ConcurrentQueue<string> S_Request, ConcurrentQueue<string> Ack_Answer)
        //----------------------------------------------------------------
        {
            // TCP/IP SERVER CHANNEL = Event processing from DB-Matik

            return Task.Run(async () =>
            {
                using (var networkStream = tcpClient.GetStream())
                {
                    networkStream.Flush();

                    while (true)
                    {
                        var buffer = new byte[65535];

                        var byteCount = await networkStream.ReadAsync(buffer, 0, buffer.Length);
                        string request = Encoding.UTF8.GetString(buffer, 0, byteCount);

                        if (byteCount > 2)
                        {
                            LogManager.DefaultLogger.Info("TCPIP -> Server2] Client wrote:\r\n" + request);

                            S_Request.Enqueue(request);
                            Ack_Answer.Enqueue(request);

                            // lookup end of XML data comming in

                            if (request.Contains("</Evt>"))
                            {
                                Thread.Sleep(100);
                                var serverResponseBytes = Encoding.UTF8.GetBytes(PrepareEvtAck(Ack_Answer));
                                await networkStream.WriteAsync(serverResponseBytes, 0, serverResponseBytes.Length);
                                LogManager.DefaultLogger.Info("TCPIP -> Server2] ACK Response has been written");
                            }
                         }
                    }

                }
            });
        }

        //=================================================================
        // TCPIP Server part END



        //----------------------------------------------------------------
        private void HandleSetVariablesResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items2 = requestXml.Descendants()
               .Select(node => node.Value.ToString())
               .ToArray();

                string Value = items2[1];
                string Result = items2[2];
                string Error = items2[3];
                string TimeStamp = items2[4];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                var items = (from x in xmlFile.Descendants("Variable")
                             select new
                             {
                                 ID = x.Attribute("ID").Value,
                                 Name = x.Attribute("Name").Value,
                                 Type = "CE",
                                 UnitID = "0",
                                 Unit = "NULL",
                                 DataTypeID = "0",
                                 DataType = "NULL",
                             }).ToArray(); ;


                foreach (var item in items)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string Type = item.Type;
                    string UnitID = item.UnitID;
                    string Unit = item.Unit;
                    string DataTypeID = item.DataTypeID;
                    string DataType = item.DataType;
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleSetVariablesResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }


        //----------------------------------------------------------------
        private void HandleGetVariablesResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = 
                (from x in xmlFile.Descendants("Variable")
                    select new
                    {
                        ID = x.Attribute("ID").Value,
                        Name = x.Attribute("Name").Value,
                        Type = x.Attribute("Type").Value,
                        UnitID = x.Attribute("UnitID").Value,
                        Unit = x.Attribute("Unit").Value,
                        DataTypeID = x.Attribute("DataTypeID").Value,
                        DataType = x.Attribute("DataType").Value,
                        Value = x.Value
                    }).ToArray(); ;

                foreach (var item in items)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string Type = item.Type;
                    string UnitID = item.UnitID;
                    string Unit = item.Unit;
                    string DataTypeID = item.DataTypeID;
                    string DataType = item.DataType;
                    string Value = item.Value;
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetVariablesResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleVariableChanged_old(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items =
                (from x in xmlFile.Descendants("Variable")
                 select new
                 {
                     ID = x.Attribute("ID").Value,
                     Name = x.Attribute("Name").Value,
                     Type = x.Attribute("Type").Value,
                     UnitID = x.Attribute("UnitID").Value,
                     Unit = x.Attribute("Unit").Value,
                     DataTypeID = x.Attribute("DataTypeID").Value,
                     DataType = x.Attribute("DataType").Value,
                     User = x.Attribute("User").Value,
                     Value = x.Value
                 }).ToArray(); ;

                foreach (var item in items)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string Type = item.Type;
                    string UnitID = item.UnitID;
                    string Unit = item.Unit;
                    string DataTypeID = item.DataTypeID;
                    string DataType = item.DataType;
                    string User = item.User;
                    string Value = item.Value;
                }

                var dataNames = new string[] { "VariableChangedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "VariableChanged");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleVariableChanged Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleGetModuleProcessStatesResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("ModuleState")
                             select new
                             {
                                 ID = x.Attribute("ID").Value,
                                 Name = x.Attribute("Name").Value,
                                 Value = x.Value
                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string Value = item.Value;
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetModuleProcessStatesResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleModuleProcessStateChanged(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;

                byte ProcessState = 0, PreviousProcessState = 0;
                string ProcessStateString = "unknown";

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("ModuleState")
                             select new
                             {
                                 ID = x.Attribute("ID").Value,
                                 Name = x.Attribute("Name").Value,
                                 Value = x.Value
                             }).ToArray();

                var items2 = (from x in xmlFile.Descendants("PreviousModuleState")
                              select new
                              {
                                  ID = x.Attribute("ID").Value,
                                  Name = x.Attribute("Name").Value,
                                  Value = x.Value
                              }).ToArray();

                foreach (var item in items)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string Value = item.Value;

                    ProcessState = byte.Parse(item.Value);

                    switch (ProcessState)
                    {
                        case 0:
                            ProcessStateString = "Init";
                            break;

                        case 1:
                            ProcessStateString = "Loaded";
                            break;

                        case 2:
                            ProcessStateString = "Ready";
                            break;

                        case 3:
                            ProcessStateString = "Standby";
                            break;

                        case 4:
                            ProcessStateString = "NoOperation";
                            break;

                        case 5:
                            ProcessStateString = "Setup";
                            break;

                        case 6:
                            ProcessStateString = "Down";
                            break;

                        case 7:
                            ProcessStateString = "Running";
                            break;

                        default:
                            ProcessStateString = "Unknown";
                            break;
                    }
                }

                foreach (var item in items2)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string Value = item.Value;

                    PreviousProcessState = byte.Parse(item.Value);
                }

                var dataNames = new string[] { "PreviousProcessState", "ProcessState", "ProcessStateString" };
                var dataValues = new CxValue[] { new U1Value(PreviousProcessState), new U1Value(ProcessState), new AValue(ProcessStateString) };

                // Den 11er CE mal raussenden
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ProcessingStateChange");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetModuleProcessStateChanged Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleAlarmSet(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;

                // Linq parser
                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string ID = items[1];
                string Text = items[2];
                string ModuleID = items[3];
                string ModuleName = items[4];
                string TimeStamp = items[5];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                // Alarm daten aus DB holen für Cimetrix 
                SQLiteCommand cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT name FROM AL WHERE alarmid = " + ID.ToString();
                SQLiteDataReader rdr = cmd.ExecuteReader();

                if(rdr.HasRows)
                {
                    rdr.Read();
                    string name = rdr.GetString(0);
                    string txt = ModuleName + ": " + Text;
                    CIM.AlarmSET(name, txt);

                    // Daten in File speichern
                    string txtsave = "ALSET," + name + "," + txt + "\r\n";
                    string DIRECTORY = ReadSetting("ALARMDIRECTORY"); ;

                    string filePath = DIRECTORY + "\\" + "Alarms_" + DateTime.Now.ToString("dd-MM-yyyy") + ".log";
                    File.AppendAllText(filePath, txtsave);
                }
                rdr.Dispose();
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleAlarmSet Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleAlarmClear(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;

                // Linq parser
                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string ID = items[1];
                string Text = items[2];
                string ModuleID = items[3];
                string ModuleName = items[4];
                string Response = items[5];
                string TimeStamp = items[6];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                // Alarm daten aus DB holen für Cimetrix 
                SQLiteCommand cmd = conn.CreateCommand();
                cmd.CommandText = "SELECT name FROM AL WHERE alarmid = " + ID.ToString();
                SQLiteDataReader rdr = cmd.ExecuteReader();

                if (rdr.HasRows)
                {
                    rdr.Read();
                    string name = rdr.GetString(0);
                    string txt = ModuleName + ": " + Response;
                    CIM.AlarmCLEAR(name, txt);

                    // Daten in File speichern
                    string txtsave = "ALCLEAR," + name + "," + txt + "\r\n";
                    string DIRECTORY = ReadSetting("ALARMDIRECTORY"); ;

                    string filePath = DIRECTORY + "\\" + "Alarms_" + DateTime.Now.ToString("dd-MM-yyyy") + ".log";
                    File.AppendAllText(filePath, txtsave);
                }
                rdr.Dispose();
                cmd.Dispose();
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleAlarmClear Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleMaterialReceived(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();


                string MaterialId = items[1];
                string MaterialName = items[2];
                string ModuleID = items[3];
                string ModuleName = items[4];
                string TimeStamp = items[5];

                var dataNames = new string[] { "MaterialReceivedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "MaterialReceived");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleMaterialReceived Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleMaterialProcessed(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();


                string MaterialId = items[1];
                string MaterialName = items[2];
                string ModuleID = items[3];
                string ModuleName = items[4];
                string TimeStamp = items[5];

                var dataNames = new string[] { "MaterialProcessedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "MaterialProcessed");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleMaterialProcessed Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleMaterialRemoved(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();


                string MaterialId = items[1];
                string MaterialName = items[2];
                string ModuleID = items[3];
                string ModuleName = items[4];
                string TimeStamp = items[5];

                var dataNames = new string[] { "MaterialRemovedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "MaterialRemoved");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleMaterialRemoved Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleExecuteRemoteCommandResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var CmdSeqID = requestXml.Attribute("CmdSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Result = items[0];
                string ERRORCODE = items[1];
                string TimeStamp = items[2];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                var items2 = (from x in xmlFile.Descendants("RemoteCmd")
                              select new
                              {
                                  ID = x.Attribute("ID").Value,
                                  Name = x.Attribute("Name").Value,
                                  ModuleID = x.Attribute("ModuleID").Value,
                                  ModuleName = x.Attribute("ModuleName").Value,
                                  Value = x.Value
                              }).ToArray(); ;

                foreach (var item in items2)
                {
                    string ID = item.ID;
                    string Name = item.Name;
                    string ModuleID = item.ModuleID;
                    string ModuleName = item.ModuleName;
                    string Value = item.Value;
                }

                var dataNames = new string[] { "ExecuteRemoteCommandResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ExecuteRemoteCommandResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("ExecuteRemoteCommandResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleUserLoggedIn(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items2 = requestXml.Descendants()
               .Select(node => node.Value.ToString())
               .ToArray();
                string username = items2[0];
                string TimeStamp = items2[1];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("User")
                             select new
                             {
                                 AccessLevel = x.Attribute("AccessLevel").Value,
                                 Default = x.Attribute("Default").Value,
                                 Deleteable = x.Attribute("Deleteable").Value,
                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string AccessLevel = item.AccessLevel;
                    string Default = item.Default;
                    string Deleteable = item.Deleteable;
                }

                var dataNames = new string[] { "UserLoggedInData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "UserLoggedIn");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleUserLoggedIn Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleUserLoggedOut(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items2 = requestXml.Descendants()
               .Select(node => node.Value.ToString())
               .ToArray();
                string username = items2[0];
                string TimeStamp = items2[1];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("User")
                             select new
                             {
                                 AccessLevel = x.Attribute("AccessLevel").Value,
                                 Default = x.Attribute("Default").Value,
                                 Deleteable = x.Attribute("Deleteable").Value,
                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string AccessLevel = item.AccessLevel;
                    string Default = item.Default;
                    string Deleteable = item.Deleteable;
                }

                var dataNames = new string[] { "UserLoggedOutData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "UserLoggedOut");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleUserLoggedOut Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleUserCreated(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items2 = requestXml.Descendants()
               .Select(node => node.Value.ToString())
               .ToArray();
                string username = items2[0];
                string TimeStamp = items2[1];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("User")
                             select new
                             {
                                 AccessLevel = x.Attribute("AccessLevel").Value,
                                 Default = x.Attribute("Default").Value,
                                 Deleteable = x.Attribute("Deleteable").Value,
                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string AccessLevel = item.AccessLevel;
                    string Default = item.Default;
                    string Deleteable = item.Deleteable;
                }

                var dataNames = new string[] { "UserCreatedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "UserCreated");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleUserCreated Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleUserDeleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items2 = requestXml.Descendants()
               .Select(node => node.Value.ToString())
               .ToArray();
                string username = items2[0];
                string TimeStamp = items2[1];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("User")
                             select new
                             {
                                 AccessLevel = x.Attribute("AccessLevel").Value,
                                 Default = x.Attribute("Default").Value,
                                 Deleteable = x.Attribute("Deleteable").Value,
                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string AccessLevel = item.AccessLevel;
                    string Default = item.Default;
                    string Deleteable = item.Deleteable;
                }

                var dataNames = new string[] { "UserDeletedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "UserDeleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleUserDeleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleControlStateChanged(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string PreviousState = items[1];
                string PreviousSubState = items[2];
                string CurrentState = items[4];
                string CurrentSubState = items[5];
                string TimeStamp = items[6];

                // Forward SECS/GEM
                if (CurrentState == "Online")
                {
                    CIM.CommunicationStateEnable();
                    Task.Delay(200).Wait();
                    CIM.Online();
                    Task.Delay(200).Wait();
                }
                else
                {
                    // Offline
                    CIM.Offline();
                    Task.Delay(200).Wait();
                    CIM.CommunicationStateDisable();
                    Task.Delay(200).Wait();
                }

                if (CurrentSubState == "Remote")
                {
                    CIM.Remote();
                }
                else
                {
                    // Local
                    CIM.Local();
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleControlStateChanged Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLotDeleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[3];
                string TimeStamp = items[5];

                var dataNames = new string[] { "LotDeletedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotDeleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotDeleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLotStarted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[3];
                string TimeStamp = items[5];

                var dataNames = new string[] { "LotStartedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotStarted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotStarted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLotCompleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                request = request.Replace("&lt;", "<");
                request = request.Replace("&gt;", ">");

                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                // the fixed first part
                //==========================
                string Name = items[1];
                string Count = items[2];
                string Product = items[3];

                string BadCount = items[6];
                string GoodCount = items[7];
                string Yield = items[8];

                string TimeStamp = items[9];
                string StartTime_dt = TimeStamp;

                TimeStamp = items[10];
                string EndTime_dt = TimeStamp;

                string AssemblyLotID = items[11];
                string AssemblyLotQty = items[13];


                // the timestamp
                //===================================================
                var items2 = requestXml.Descendants("TimeStamp")
                              .Select(node => node.Value.ToString())
                              .ToArray();

                TimeStamp = items2[0];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime TimeStamp_dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                // the variable part
                //==========================================
                var items3 = (from x in requestXml.Descendants("AssemblyLotRejects")
                              select new
                              {
                                  Value = x.Elements(),
                              }).ToArray();

                var mycxValue = new List<CxValue>();
                var mycxValue2 = new List<CxValue>();

                foreach (var item in items3)
                {
                    foreach (var z in item.Value)
                    {
                        if (z.Name.ToString().Contains("AssemblyLotRejectLossCode"))
                        {
                            mycxValue.Add(new AValue(z.Value.ToString()));
                        }
                    }
                }

                foreach (var item in items3)
                {
                    foreach (var z in item.Value)
                    {
                        if (z.Name.ToString().Contains("AssemblyLotRejectLossQty"))
                        {
                            uint i = byte.Parse(z.Value);
                            mycxValue2.Add(new U4Value(i));
                        }
                    }
                }

                var dataNames = new string[]
                {
                    "LotCompleted_Name",
                    "LotCompleted_Count",
                    "LotCompleted_Product",
                    "LotCompleted_BadCount",
                    "LotCompleted_GoodCount",
                    "LotCompleted_Yield",
                    "LotCompleted_StartTime",
                    "LotCompleted_EndTime",
                    "LotCompleted_AssemblyLotID",
                    "LotCompleted_AssemblyLotQty",
                    "LotCompleted_AssemblyLotRejects",
                    "LotCompleted_AssemblyLotRejects_Val"
                };


                var dataValues = new CxValue[]
                {
                    new AValue(Name),
                    new I4Value(Int32.Parse(Count)),
                    new AValue(Product),
                    new I4Value(Int32.Parse(BadCount)),
                    new I4Value(Int32.Parse(GoodCount)),
                    new F4Value(float.Parse(Yield)),
                    new AValue(StartTime_dt),
                    new AValue(EndTime_dt),
                    new AValue(AssemblyLotID),
                    new I4Value(Int32.Parse(AssemblyLotQty)),
                    new LValue(mycxValue),
                    new LValue(mycxValue2),
                };

                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotCompleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotCompleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLMaterialReport(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                request = request.Replace("&lt;", "<");
                request = request.Replace("&gt;", ">");

                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants("Lot")
                               .Select(node => node.ToString())
                               .ToArray();

                // the  first part (Lot)
                //==========================
                string Lot = items[0];
                var requestLot = XElement.Parse(Lot);
                var LotName = requestLot.Attribute("Name").Value;

                // now the fixed part
                //==========================
                var items2 = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string AssemblyLotID = items2[3];
                string MaterialId = items2[4];
                string MaterialConsumption = items2[5];

                // the variable part
                //==========================================
                var items3 = (from x in requestXml.Descendants("LGARejects")
                              select new
                              {
                                  Value = x.Elements(),
                              }).ToArray();

                var mycxValue = new List<CxValue>();
                var mycxValue2 = new List<CxValue>();

                foreach (var item in items3)
                {
                    foreach (var z in item.Value)
                    {
                        if (z.Name.ToString().Contains("LGARejectsLossCode"))
                        {
                            mycxValue.Add(new AValue(z.Value.ToString()));
                        }
                    }
                }

                foreach (var item in items3)
                {
                    foreach (var z in item.Value)
                    {
                        if (z.Name.ToString().Contains("LGARejectsQty"))
                        {
                            uint i = byte.Parse(z.Value);
                            mycxValue2.Add(new U4Value(i));
                        }
                    }
                }

                var dataNames = new string[]
                {
                    "MaterialReport_LotName",
                    "MaterialReport_AssemblyLotID",
                    "MaterialReport_MaterialId",
                    "MaterialReport_MaterialConsumption",
                    "MaterialReport_LGARejects",
                    "MaterialReport_LGARejects_Val",
                };

                var dataValues = new CxValue[]
                {
                    new AValue(LotName),
                    new AValue(AssemblyLotID),
                    new AValue(MaterialId),
                    new AValue(MaterialConsumption),
                    new LValue(mycxValue),
                    new LValue(mycxValue2),
                };

                CIM.SendCollectionEventWithData(dataNames, dataValues, "MaterialReport");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("MaterialReport Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLotAborted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[3];
                string Result = items[5];
                string TimeStamp = items[6];

                var dataNames = new string[] { "LotAbortedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotAborted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotAborted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLotPaused(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[3];
                string Result = items[5];
                string TimeStamp = items[6];

                var dataNames = new string[] { "LotPausedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotPaused");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotPaused Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }


        //----------------------------------------------------------------
        private void HandleLotCreated(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[3];
                string TimeStamp = items[5];

                var dataNames = new string[] { "LotCreatedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotCreated");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotCreated Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleLotResumed(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[3];
                string TimeStamp = items[5];
                TimeStamp = TimeStamp.Substring(0, 14);

                var dataNames = new string[] { "LotResumedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "LotResumed");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleLotResumed Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleProductCreated(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string TimeStamp = items[2];

                var dataNames = new string[] { "ProductCreatedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ProductCreated");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleProductCreated Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleProductSelected(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string TimeStamp = items[2];

                var dataNames = new string[] { "ProductSelectedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ProductSelected");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleProductSelected Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleProductDeleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string TimeStamp = items[2];

                var dataNames = new string[] { "ProductDeletedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ProductDeleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleProductDeleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleProductStored(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string TimeStamp = items[2];

                var dataNames = new string[] { "ProductStoredData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ProductStored");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleProductStored Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleProductDownloaded(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string TimeStamp = items[2];

                var dataNames = new string[] { "ProductDownloadedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ProductDownloaded");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleProductDownloaded Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleOperatorCommandExecuted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Command = items[1];
                string TimeStamp = items[2];

                var dataNames = new string[] { "OperatorCommandExecutedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "OperatorCommandExecuted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleOperatorCommandExecuted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleItemsProcessStarted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants("Item")
                               .Select(node => node.ToString())
                               .ToArray();

                foreach (var item in items)
                {
                    string txt = item.ToString();

                    XDocument xmlFile = XDocument.Parse(txt);

                    var items2 = (from x in xmlFile.Descendants()
                                  select new
                                  {
                                      ItemId = x.Element("ItemID"),
                                      ModuleID = x.Element("ModuleID"),
                                      ModuleName = x.Element("ModuleName"),
                                      TrackingNumber = x.Element("TrackingNumber"),
                                      ShiftRegisterPos = x.Element("ShiftRegisterPos"),
                                  }).ToArray(); ;

                    string ItemId = "NULL", ModuleID = "NULL", ModuleName = "NULL", TrackingNumber = "NULL", ShiftRegisterPos = "NULL";

                    foreach (var item2 in items2)
                    {
                        ItemId = item2.ItemId.Value;
                        ModuleID = item2.ModuleID.Value;
                        ModuleName = item2.ModuleName.Value;
                        TrackingNumber = item2.TrackingNumber.Value;
                        ShiftRegisterPos = item2.ShiftRegisterPos.Value;
                        break;
                    }
                }

                var dataNames = new string[] { "ItemsProcessStartedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ItemsProcessStarted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleItemsProcessStarted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleItemsProcessCompleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants("Item")
                               .Select(node => node.ToString())
                               .ToArray();

                foreach (var item in items)
                {
                    string txt = item.ToString();

                    XDocument xmlFile = XDocument.Parse(txt);

                    var items2 = (from x in xmlFile.Descendants()
                                  select new
                                  {
                                      ItemId = x.Element("ItemID"),
                                      ModuleID = x.Element("ModuleID"),
                                      ModuleName = x.Element("ModuleName"),
                                      TrackingNumber = x.Element("TrackingNumber"),
                                      ShiftRegisterPos = x.Element("ShiftRegisterPos"),
                                      Result = x.Element("Result"),
                                      ResultData = x.Element("ResultData"),
                                  }).ToArray(); ;

                    string ItemId = "NULL", ModuleID = "NULL", ModuleName = "NULL",
                        TrackingNumber = "NULL", ShiftRegisterPos = "NULL", Result = "NULL", ResultData = "NULL";

                    foreach (var item2 in items2)
                    {
                        ItemId = item2.ItemId.Value;
                        ModuleID = item2.ModuleID.Value;
                        ModuleName = item2.ModuleName.Value;
                        TrackingNumber = item2.TrackingNumber.Value;
                        ShiftRegisterPos = item2.ShiftRegisterPos.Value;
                        Result = item2.Result.Value;
                        ResultData = item2.ResultData.Value;
                        break;
                    }
                }

                var dataNames = new string[] { "ItemsProcessCompletedData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ItemsProcessCompleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleItemsProcessCompleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleItemProcessStarted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                string ShiftRegisterPos = "NULL";

                var items = requestXml.Descendants("Item")
                               .Select(node => node.ToString())
                               .ToArray();

                foreach (var item in items)
                {
                    string txt = item.ToString();

                    XDocument xmlFile = XDocument.Parse(txt);

                    var items2 = (from x in xmlFile.Descendants()
                                  select new
                                  {
                                      ItemId = x.Element("ItemID"),
                                      ModuleID = x.Element("ModuleID"),
                                      ModuleName = x.Element("ModuleName"),
                                      TrackingNumber = x.Element("TrackingNumber"),
                                      ShiftRegisterPos = x.Element("ShiftRegisterPos"),
                                  }).ToArray(); ;

                    string ItemId = "NULL", ModuleID = "NULL", ModuleName = "NULL", TrackingNumber = "NULL";

                    foreach (var item2 in items2)
                    {
                        ItemId = item2.ItemId.Value;
                        ModuleID = item2.ModuleID.Value;
                        ModuleName = item2.ModuleName.Value;
                        TrackingNumber = item2.TrackingNumber.Value;
                        ShiftRegisterPos = item2.ShiftRegisterPos.Value;
                        break;
                    }
                }

                var dataNames = new string[] { "ItemProcessStartedData" };
                var dataValues = new CxValue[] { new AValue(ShiftRegisterPos) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ItemProcessStarted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleItemProcessStarted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleItemProcessCompleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                string ShiftRegisterPos = "NULL";


                var items = requestXml.Descendants("Item")
                               .Select(node => node.ToString())
                               .ToArray();

                foreach (var item in items)
                {
                    string txt = item.ToString();

                    XDocument xmlFile = XDocument.Parse(txt);

                    var items2 = (from x in xmlFile.Descendants()
                                  select new
                                  {
                                      ItemId = x.Element("ItemID"),
                                      ModuleID = x.Element("ModuleID"),
                                      ModuleName = x.Element("ModuleName"),
                                      TrackingNumber = x.Element("TrackingNumber"),
                                      ShiftRegisterPos = x.Element("ShiftRegisterPos"),
                                      Result = x.Element("Result"),
                                      ResultData = x.Element("ResultData"),
                                  }).ToArray(); ;

                    string ItemId = "NULL", ModuleID = "NULL", ModuleName = "NULL",
                        TrackingNumber = "NULL", Result = "NULL", ResultData = "NULL";

                    foreach (var item2 in items2)
                    {
                        ItemId = item2.ItemId.Value;
                        ModuleID = item2.ModuleID.Value;
                        ModuleName = item2.ModuleName.Value;
                        TrackingNumber = item2.TrackingNumber.Value;
                        ShiftRegisterPos = item2.ShiftRegisterPos.Value;
                        Result = item2.Result.Value;
                        ResultData = item2.ResultData.Value;
                        break;
                    }
                }

                var dataNames = new string[] { "ItemProcessCompletedData" };
                var dataValues = new CxValue[] { new AValue(ShiftRegisterPos) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ItemProcessCompleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleItemProcessCompleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleItemCompleted(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants("ProducedUnits")
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string strUnits = items[0];


                var dataNames = new string[] { "ItemCompletedData" };
                var dataValues = new CxValue[] { new I4Value(Int32.Parse(strUnits)) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "ItemCompleted");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleItemCompleted Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleCreateLotResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Result = items[1];
                string Name = items[2];
                string Count = items[3];
                string Product = items[4];

                var dataNames = new string[] { "CreateLotResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "CreateLotResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleICreateLotResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleGetLotResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[1];
                string Count = items[2];
                string Product = items[4];

                var dataNames = new string[] { "GetLotResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "GetLotResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetLotResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleGetLotsResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items3 = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Result = items3[0];


                var items = requestXml.Descendants("Lot")
                               .Select(node => node.ToString())
                               .ToArray();

                foreach (var item in items)
                {
                    string txt = item.ToString();

                    XDocument xmlFile = XDocument.Parse(txt);

                    var items2 = (from x in xmlFile.Descendants()
                                  select new
                                  {
                                      Name = x.Element("Name"),
                                      Count = x.Element("Count"),
                                      Product = x.Element("Product"),
                                  }).ToArray(); ;

                    string Name = "NULL", Count = "NULL", Product = "NULL";

                    foreach (var item2 in items2)
                    {
                        Name = item2.Name.Value;
                        Count = item2.Count.Value;
                        Product = item2.Product.Value;
                        break;
                    }
                }

                var dataNames = new string[] { "GetLotsResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "GetLotsResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetLotsResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleDeleteLotResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
               .Select(node => node.Value.ToString())
               .ToArray();

                string Result = items[0];
                string Name = items[1];

                var dataNames = new string[] { "DeleteLotResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "DeleteLotResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleDeleteLotResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleGetProductsResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("Product")
                             select new
                             {
                                 Value = x.Value,
                                 Result = "NULL",
                                 Error = "0",
                             }).ToArray();



                // delete old Recipes in DB
                SQLiteCommand cmd = conn.CreateCommand();

                try
                {
                    cmd.CommandText = "DELETE FROM Product WHERE Action = 'GetProductsResponse'";
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                catch (Exception ex)
                {
                    LogManager.DefaultLogger.Error("HandleGetProductsResponse Exception:");
                    LogManager.DefaultLogger.Error(cmd.CommandText);
                    LogManager.DefaultLogger.Error(ex.Message);
                }

                foreach (var item in items)
                {
                    string Name = item.Value;
                    string Result = item.Result;
                    string Error = item.Error;
                    DateTime dt = DateTime.Now;
                }

                var dataNames = new string[] { "GetProductsResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "GetProductsResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetProductsResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleSelectProductResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Name = items[0];
                string Result = items[1];
                string Error = items[2];
                string TimeStamp = items[3];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                var dataNames = new string[] { "SelectProductResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "SelectProductResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleSelectProductResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleDownloadProductResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("Product")
                             select new
                             {
                                 Name = x.Attribute("Name").Value,
                                 Content = x.Value,
                                 Result = "NULL",
                                 Error = "0",
                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string Name = item.Name;
                    string Content = item.Content;
                    string Result = item.Result;
                    string Error = item.Error;
                    DateTime dt = DateTime.Now;

                    break;
                }

                var dataNames = new string[] { "DownloadProductResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "DownloadProductResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleDownloadProductResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleUploadProductResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Result = items[1];
                string Error = items[2];
                string TimeStamp = items[3];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                XDocument xmlFile = XDocument.Parse(request);

                var items2 = (from x in xmlFile.Descendants("Product")
                              select new
                              {
                                  Name = x.Attribute("Name").Value,
                              }).ToArray(); ;

                foreach (var item2 in items2)
                {
                    string Name = item2.Name;

                    break;
                }

                var dataNames = new string[] { "UploadProductResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "UploadProductResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleUploadProductResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleRenameProductResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants()
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Result = items[1];
                string Error = items[2];
                string TimeStamp = items[3];
                TimeStamp = TimeStamp.Substring(0, 14);
                DateTime dt = DateTime.ParseExact(TimeStamp, "yyyyMMddHHmmss", CultureInfo.InvariantCulture);

                XDocument xmlFile = XDocument.Parse(request);

                var items2 = (from x in xmlFile.Descendants("Product")
                              select new
                              {
                                  Name = x.Attribute("Name").Value,
                              }).ToArray(); ;

                foreach (var item2 in items2)
                {
                    string Name = item2.Name;

                    break;
                }

                var dataNames = new string[] { "RenameProductResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "RenameProductResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleRenameProductResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleSetTerminalMessageResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                var items = requestXml.Descendants("Response")
                               .Select(node => node.Value.ToString())
                               .ToArray();

                string Response = items[0];

                var dataNames = new string[] { "SetTerminalMessageResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "SetTerminalMessageResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleSetTerminalMessageResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleGetUsersResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("User")
                             select new
                             {
                                 AccessLevel = x.Attribute("AccessLevel").Value,
                                 Default = x.Attribute("Default").Value,
                                 Deleteable = x.Attribute("Deleteable").Value,
                                 LoggedIn = x.Attribute("LoggedIn").Value,
                                 username = x.Value,

                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string username = item.username;
                    string AccessLevel = item.AccessLevel;
                    string Default = item.Default;
                    string Deleteable = item.Deleteable;
                    string LoggedIn = item.LoggedIn;
                    DateTime dt = DateTime.Now;
                }

                var dataNames = new string[] { "GetUsersResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "GetUsersResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetUsersResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleGetCurrentLoggedInUserResponse(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                var SeqID = requestXml.Attribute("SeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("User")
                             select new
                             {
                                 AccessLevel = x.Attribute("AccessLevel").Value,
                                 Default = x.Attribute("Default").Value,
                                 Deleteable = x.Attribute("Deleteable").Value,
                                 LoggedIn = x.Attribute("LoggedIn").Value,
                                 username = x.Value,

                             }).ToArray(); ;

                foreach (var item in items)
                {
                    string username = item.username;
                    string AccessLevel = item.AccessLevel;
                    string Default = item.Default;
                    string Deleteable = item.Deleteable;
                    string LoggedIn = item.LoggedIn;
                    DateTime dt = DateTime.Now;
                }

                var dataNames = new string[] { "GetLoggedInUsersResponseData" };
                var dataValues = new CxValue[] { new AValue(request) };
                CIM.SendCollectionEventWithData(dataNames, dataValues, "GetLoggedInUsersResponse");
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleGetCurrentLoggedInUserResponse Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        private void HandleVariableChanged(string request)
        //----------------------------------------------------------------
        {
            mutex.WaitOne();

            string Name = string.Empty, Type = string.Empty, DataType = string.Empty, Val = string.Empty;

            LogManager.DefaultLogger.Error("Entering HandleVariableChanged");

            try
            {
                var requestXml = XElement.Parse(request);
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;

                XDocument xmlFile = XDocument.Parse(request);

                var items = (from x in xmlFile.Descendants("Variable")
                             select new
                             {
                                 Name = x.Attribute("Name").Value,
                                 Type = x.Attribute("Type").Value,
                                 DataType = x.Attribute("DataType").Value,
                                 val = x.Value,

                             }).ToArray(); ;



                foreach (var item in items)
                {
                    Name = item.Name;
                    Type = item.Type;
                    DataType = item.DataType;
                    Val = item.val;
                }

                var dataNames = new string[]
                {
                    "VariableChanged_Name",
                    "VariableChanged_ObjectType",
                    "VariableChanged_DataType",
                    "VariableChanged_Value",
                };

                var dataValues = new CxValue[]
                {
                    new AValue(Name),
                    new AValue(Type),
                    new AValue(DataType),
                    new AValue(Val),
                };

                CIM.SendCollectionEventWithData(dataNames, dataValues, "VariableChanged");

            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("HandleVariableChanged Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }
            finally
            {
                mutex.ReleaseMutex();
            }
        }

        //----------------------------------------------------------------
        public void StopAndWait()
        //----------------------------------------------------------------
        {
            IsStopped = true;

            StopListener();

            if (WorkerThread != null)
            {
                WorkerThread.Join();
            }
        }

        //----------------------------------------------------------------
        private static string PrepareEvtAck(ConcurrentQueue<string> Ack_Answer)
        //----------------------------------------------------------------
        {
            string Request, txt = String.Empty, RequestCompletet = String.Empty;

            try
            {
                while (Ack_Answer.TryDequeue(out Request))
                {
                    RequestCompletet = RequestCompletet + Request;
                }

                int count = Regex.Matches(RequestCompletet, "</Evt>").Count;

                if (count == 2)
                {
                    // remove double event
                    int index = RequestCompletet.IndexOf("<Evt ", count);
                    RequestCompletet = RequestCompletet.Substring(0, index);
                }

                // parse XML for Answer
                var requestXml = XElement.Parse(RequestCompletet);
                var ID = requestXml.Attribute("ID").Value;
                var EquipID = requestXml.Attribute("EquipID").Value;
                var EvtSeqID = requestXml.Attribute("EvtSeqID").Value;
                string TimeStamp = DateTime.Now.ToString("yyyyMMddhhmmss");

                // prepare answer string
                using (var sw = new StringWriter())
                {
                    using (var xw = XmlWriter.Create(sw))
                    {
                        xw.WriteStartElement("EvtAck");
                        xw.WriteAttributeString("ID", ID);
                        xw.WriteAttributeString("EquipID", EquipID);
                        xw.WriteAttributeString("EvtSeqID", EvtSeqID);
                        xw.WriteElementString("Result", "OK");
                        xw.WriteElementString("Error", "0");
                        xw.WriteElementString("TimeStamp", TimeStamp);
                        xw.WriteEndElement();
                    }

                    txt = sw.ToString();
                    txt = txt.Replace("<?xml version=\"1.0\" encoding=\"utf-16\"?>", "");
                    LogManager.DefaultLogger.Info("Prepared EvtAck:\r\n" + txt);
                }
            }
            catch (Exception ex)
            {
                LogManager.DefaultLogger.Error("PrepareEvtAck Exception:");
                LogManager.DefaultLogger.Error(ex.Message);
            }

            return txt;
        }

        //----------------------------------------------------------------
        private void Worker()
        //----------------------------------------------------------------
        {
            string Request, RequestCompletet = String.Empty;

            while(!IsStopped)
            {
                if(!S_Request.TryDequeue(out Request) && RequestCompletet == String.Empty)
                {
                    Thread.Sleep(100);
                    continue;
                }

                RequestCompletet = RequestCompletet + Request;

                // TODO: eventuell prüfen ob nach </Evt> nicht noch was folgt 
                // was bereits zur nächsten transaktion gehört -> wieder in die queu geben

                if (RequestCompletet.Contains("</Evt>"))
                {
                    int count = Regex.Matches(RequestCompletet, "</Evt>").Count;

                    if(count >= 2)
                    {
                        // remove double event
                        int index = RequestCompletet.IndexOf("<Evt ", 2);
                        Request = RequestCompletet.Substring(0, index);
                        RequestCompletet = RequestCompletet.Substring(index);
                    }
                    else
                    {
                        Request = RequestCompletet;
                        RequestCompletet = String.Empty;
                    }

                    LogManager.DefaultLogger.Info("[Worker2] Client wrote:\r\n" + Request);
                }
                else
                {
                    continue;
                }

                // Events below
                //----------------------------------------------------

                if (Request.Contains("AlarmSet"))
                {
                    var request = Request;
                    Task.Run(() => HandleAlarmSet(request));
                    continue;
                }

                if (Request.Contains("AlarmCleared"))
                {
                    var request = Request;
                    Task.Run(() => HandleAlarmClear(request));
                    continue;
                }

                if (Request.Contains("VariableChanged"))
                {
                    var request = Request;
                    Task.Run(() => HandleVariableChanged(request));
                    continue;
                }

                if (Request.Contains("ModuleProcessStatesChanged"))
                {
                    var request = Request;
                    Task.Run(() => HandleModuleProcessStateChanged(request));
                    continue;
                }

                if (Request.Contains("MaterialReceived"))
                {
                    var request = Request;
                    Task.Run(() => HandleMaterialReceived(request));
                    continue;
                }

                if (Request.Contains("MaterialProcessed"))
                {
                    var request = Request;
                    Task.Run(() => HandleMaterialProcessed(request));
                    continue;
                }

                if (Request.Contains("MaterialRemoved"))
                {
                    var request = Request;
                    Task.Run(() => HandleMaterialRemoved(request));
                    continue;
                }

                if (Request.Contains("UserLoggedIn"))
                {
                    var request = Request;
                    Task.Run(() => HandleUserLoggedIn(request));
                    continue;
                }

                if (Request.Contains("UserLoggedOut"))
                {
                    var request = Request;
                    Task.Run(() => HandleUserLoggedOut(request));
                    continue;
                }

                if (Request.Contains("UserCreated"))
                {
                    var request = Request;
                    Task.Run(() => HandleUserCreated(request));
                    continue;
                }

                if (Request.Contains("UserDeleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleUserDeleted(request));
                    continue;
                }

                if (Request.Contains("ControlStateChanged"))
                {
                    var request = Request;
                    Task.Run(() => HandleControlStateChanged(request));
                    continue;
                }

                if (Request.Contains("LotCreated"))
                {
                    var request = Request;
                    Task.Run(() => HandleLotCreated(request));
                    continue;
                }

                if (Request.Contains("LotDeleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleLotDeleted(request));
                    continue;
                }

                if (Request.Contains("LotStarted"))
                {
                    var request = Request;
                    Task.Run(() => HandleLotStarted(request));
                    continue;
                }

                if (Request.Contains("LotCompleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleLotCompleted(request));
                    continue;
                }

                if (Request.Contains("LotAborted "))
                {
                    var request = Request;
                    Task.Run(() => HandleLotAborted(request));
                    continue;
                }

                if (Request.Contains("LotPaused"))
                {
                    var request = Request;
                    Task.Run(() => HandleLotPaused(request));
                    continue;
                }

                if (Request.Contains("LotResumed"))
                {
                    var request = Request;
                    Task.Run(() => HandleLotResumed(request));
                    continue;
                }

                if (Request.Contains("ProductCreated"))
                {
                    var request = Request;
                    Task.Run(() => HandleProductCreated(request));
                    continue;
                }

                if (Request.Contains("ProductSelected"))
                {
                    var request = Request;
                    Task.Run(() => HandleProductSelected(request));
                    continue;
                }

                if (Request.Contains("ProductDeleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleProductDeleted(request));
                    continue;
                }

                if (Request.Contains("ProductStored"))
                {
                    var request = Request;
                    Task.Run(() => HandleProductStored(request));
                    continue;
                }

                if (Request.Contains("ProductDownloaded"))
                {
                    var request = Request;
                    Task.Run(() => HandleProductDownloaded(request));
                    continue;
                }

                if (Request.Contains("OperatorCommandExecuted"))
                {
                    var request = Request;
                    Task.Run(() => HandleOperatorCommandExecuted(request));
                    continue;
                }

                if (Request.Contains("ItemsProcessStarted"))
                {
                    var request = Request;
                    Task.Run(() => HandleItemsProcessStarted(request));
                    continue;
                }

                if (Request.Contains("ItemsProcessCompleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleItemsProcessCompleted(request));
                    continue;
                }

                if (Request.Contains("ItemProcessStarted"))
                {
                    var request = Request;
                    Task.Run(() => HandleItemProcessStarted(request));
                    continue;
                }

                if (Request.Contains("ItemProcessCompleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleItemProcessCompleted(request));
                    continue;
                }

                if (Request.Contains("ItemCompleted"))
                {
                    var request = Request;
                    Task.Run(() => HandleItemCompleted(request));
                    continue;
                }

                if (Request.Contains("MaterialReport"))
                {
                    var request = Request;
                    Task.Run(() => HandleLMaterialReport(request));
                    continue;
                }

                // Responses below
                //----------------------------------------------------

                if (Request.Contains("GetVariablesResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetVariablesResponse(request));
                    continue;
                }

                if (Request.Contains("SetVariablesResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleSetVariablesResponse(request));
                    continue;
                }

                if (Request.Contains("GetModuleProcessStatesResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetModuleProcessStatesResponse(request));
                    continue;
                }

                if (Request.Contains("ExecuteRemoteCommandResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleExecuteRemoteCommandResponse(request));
                    continue;
                }

                if (Request.Contains("CreateLotResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleCreateLotResponse(request));
                    continue;
                }

                if (Request.Contains("GetLotResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetLotResponse(request));
                    continue;
                }

                if (Request.Contains("GetLotsResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetLotsResponse(request));
                    continue;
                }

                if (Request.Contains("DeleteLotResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleDeleteLotResponse(request));
                    continue;
                }

                if (Request.Contains("GetProductsResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetProductsResponse(request));
                    continue;
                }

                if (Request.Contains("SelectProductResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleSelectProductResponse(request));
                    continue;
                }

                if (Request.Contains("DownloadProductResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleDownloadProductResponse(request));
                    continue;
                }

                if (Request.Contains("UploadProductResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleUploadProductResponse(request));
                    continue;
                }

                if (Request.Contains("RenameProductResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleRenameProductResponse(request));
                    continue;
                }

                if (Request.Contains("SetTerminalMessageResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleSetTerminalMessageResponse(request));
                    continue;
                }

                if (Request.Contains("GetUsersResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetUsersResponse(request));
                    continue;
                }

                if (Request.Contains("GetCurrentLoggedInUserResponse"))
                {
                    var request = Request;
                    Task.Run(() => HandleGetCurrentLoggedInUserResponse(request));
                    continue;
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
                string result = appSettings[key] ?? "Not Found 2";
                return result;
            }
            catch (ConfigurationErrorsException)
            {
                return "nil";
            }
        }
    }
}
