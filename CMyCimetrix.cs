using System;
using System.IO;
using System.Xml;
using System.Runtime.InteropServices;

using Cimetrix.Value;

using EMSERVICELib;

using VALUELib;

using ValueType = VALUELib.ValueType;

namespace BiometricServer
{
    /// <summary>
    ///    Summary description for MyCIMConnect.
    /// </summary>

    //CMyCIMConnect object implements EMSERVICELib.ICxEMAppCallback interface	
    public class CMyCimetrix : ICxEMAppCallback, ICxEMAppCallback6
	{
		#region Member Variables

		private const int E_FAIL = unchecked((int)0x80004005);

		private const int E_NOTIMPL = unchecked((int)0x80004001);

		public const long _NumberOfConnections = 1;

		public const int _StartupWaitTimeInMS = 15000;

		public CommunicationState[] _CommunicationState = new CommunicationState[_NumberOfConnections];

		public ControlState[] _ControlState = new ControlState[_NumberOfConnections];

		public SpoolingState[] _SpoolingState = new SpoolingState[_NumberOfConnections];

		/// <summary>
		///    Client application interface.
		/// </summary>
		public CxClientClerk _CxClientClerk;

		/// <summary>
		///    The main EMService/CIMConnect interface.
		/// </summary>
		public CxEMService _CxEMService;

		public ICxClientTool _ICxClientTool;

		public bool _initialized = false;


		public int _VID_CommEnableSwitch = 2035;
		public int _VID_CtrlOnlineSwitch = 2034;
		public int _VID_ControlStateSwitch = 2033;


		#endregion Member Variables

		#region Delegates

		public delegate void CommEnableDisableSwitchHandler(object sender, CommEnableDisableSwitchArgs e);

		public delegate void ControlLocalRemoteSwitchHandler(object sender, ControlLocalRemoteSwitchArgs e);

		public delegate void ControlOfflineOnlineSwitchHandler(object sender, ControlOfflineOnlineSwitchArgs e);




		#endregion

		#region Event Handlers
		public delegate void onInitializationStartHandler();

		public event onInitializationStartHandler OnInitializationStart;

		private void NotifyInitializationStart()
		{
			if (OnInitializationStart != null)
			{
				OnInitializationStart();
			}
		}


		public delegate void onInitializationCompleteHandler(bool bSuccess, string errMessage);

		public event onInitializationCompleteHandler OnInitializationComplete;

		private void NotifyInitializationComplete(bool bSuccess, string errMessage)
		{
			if (OnInitializationComplete != null)
			{
				OnInitializationComplete(bSuccess, errMessage);
			}
		}

		// Handle incoming Terminal Service messages. 
		public delegate void TerminalServiceHandler(object sender, TerminalServiceArgs e);

		public event TerminalServiceHandler OnTerminalService;

		private void TerminalService(string textMessage)
		{
			if (OnTerminalService != null)
			{
				var e = new TerminalServiceArgs(textMessage);
				OnTerminalService(this, e);
			}
		}

		// Handle general application logging messages. 
		public delegate void LogHandler(object sender, LogArgs e);

		public event LogHandler OnLog;

		public void Log(string textMessage)
		{
			if (OnLog != null)
			{
				var e = new LogArgs(textMessage);
				// Asynchronous fire and forget
				OnLog.BeginInvoke(this, e, null, null);
			}
		}

		// Handle reporting GEM state changes 
		public delegate void GEMStateChangeHandler(object sender, GEMStateChangeArgs e);

		public event GEMStateChangeHandler OnGEMStateChange;

		private void GEMStateChange(int connection, StateMachine statemachine, int state)
		{
			if (OnGEMStateChange != null)
			{
				var e = new GEMStateChangeArgs(connection, statemachine, state);
				// Asynchronous fire and forget
				OnGEMStateChange.BeginInvoke(this, e, null, null);
			}
		}

		public event CommEnableDisableSwitchHandler OnCommEnableDisableSwitchHandler;

		private void CommEnableDisableSwitchChange(int connId, CommEnableDisableEnum newSetting)
		{
			if (OnCommEnableDisableSwitchHandler != null)
			{
				var e = new CommEnableDisableSwitchArgs(connId, newSetting);
				// Asynchronously fire and forget
				OnCommEnableDisableSwitchHandler.BeginInvoke(this, e, null, null);
			}
		}

		public event ControlOfflineOnlineSwitchHandler OnCtrlOfflineOnlineSwitchHandler;

		private void CtrlOfflineOnlineSwitchChange(int connId, ControlOfflineOnlineEnum newSetting)
		{
			if (OnCtrlOfflineOnlineSwitchHandler != null)
			{
				var e = new ControlOfflineOnlineSwitchArgs(connId, newSetting);
				// Asynchronously fire and forget
				OnCtrlOfflineOnlineSwitchHandler.BeginInvoke(this, e, null, null);
			}
		}

		public event ControlLocalRemoteSwitchHandler OnCtrlLocalRemoteSwitchHandler;

		private void CtrlLocalRemoteSwitchChange(int connId, ControlLocalRemoteEnum newSetting)
		{
			if (OnCtrlLocalRemoteSwitchHandler != null)
			{
				var e = new ControlLocalRemoteSwitchArgs(connId, newSetting);
				// Asynchronously fire and forget
				OnCtrlLocalRemoteSwitchHandler.BeginInvoke(this, e, null, null);
			}
		}

		/// <summary>
		/// Handle Remote Commands, CommandCalled callback
		/// IMPORTANT: This is executed synchronously
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void RemoteCommandHandler(object sender, RemoteCommandArgs e);

		public event RemoteCommandHandler OnRemoteCommand;

		private void RemoteCommand(
			int connection,
			string command,
			CxValueObject argumentNames,
			CxValueObject argumentValues,
			ref CommandResults commandResult)
		{
			if (OnRemoteCommand != null)
			{
				var e = new RemoteCommandArgs(connection, command, argumentNames, argumentValues,
					ref commandResult);
				// Synchronous
				OnRemoteCommand(this, e);
				commandResult = e.CommandResult;
			}
		}

		//  
		/// <summary>
		/// Handle GetValueToByteBuffer callback. 
		/// IMPORTANT: This is executed synchronously. 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void GetValueCallbackHandler(object sender, GetValueCallbackArgs e);

		public event GetValueCallbackHandler OnGetValueCallback;

		private object GetValueCallback(string variableName, int variableID)
		{
			if (OnGetValueCallback != null)
			{
				var e = new GetValueCallbackArgs(variableName, variableID);
				OnGetValueCallback(this, e);
				return e.VariableValue.ToByteBuffer();
			}
			return null;
		}


		//  
		/// <summary>
		/// Handle VerifyValueToByteBuffer callback. 
		/// IMPORTANT: This is executed synchronously. 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void VerifyValueCallbackHandler(object sender, VerifyValueCallbackArgs e);

		public event VerifyValueCallbackHandler OnVerifyValueCallback;

		private VerifyValueResults VerifyValueCallback(string variableName, int variableID, CxValueObject variableValue)
		{
			if (OnVerifyValueCallback != null)
			{
				var e = new VerifyValueCallbackArgs(variableName, variableID, variableValue);
				OnVerifyValueCallback(this, e);
				return e.Result;
			}
			return VerifyValueResults.vvVerified;
		}

		public class VerifyValueCallbackArgs : EventArgs
		{
			public VerifyValueCallbackArgs(string variableName, int variableID, CxValueObject variableValue)
			{
				VariableName = variableName;
				VariableID = variableID;
				Result = VerifyValueResults.vvVerified;
				Value = variableValue;
			}

			public string VariableName { get; private set; }
			public int VariableID { get; private set; }
			public VerifyValueResults Result { get; set; }
			public CxValueObject Value { get; private set; }
		}

		//  
		/// <summary>
		/// Handle ValueChangedToByteBuffer callback. 
		/// IMPORTANT: This is executed asynchronously. 
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		public delegate void ValueChangedCallbackHandler(object sender, ValueChangedCallbackArgs e);

		public event ValueChangedCallbackHandler OnValueChangedCallback;

		private void ValueChangedCallback(string variableName, int variableID, object variableValue)
		{
			if (OnValueChangedCallback != null)
			{
				var e = new ValueChangedCallbackArgs(variableName, variableID, variableValue);
				OnValueChangedCallback.BeginInvoke(this, e, null, null); // fire and forget
			}
		}

		public class ValueChangedCallbackArgs : EventArgs
		{
			public ValueChangedCallbackArgs(string variableName, int variableID, object variableValue)
			{
				VariableName = variableName;
				VariableID = variableID;
				Value = variableValue;
			}

			public string VariableName { get; private set; }
			public int VariableID { get; private set; }
			public object Value { get; private set; }
		}


		#endregion

		#region Types for Event Handlers

		#region Nested type: CommEnableDisableSwitchArgs

		public class CommEnableDisableSwitchArgs : EventArgs
		{
			public CommEnableDisableSwitchArgs(int connId, CommEnableDisableEnum newSetting)
			{
				ConnectionId = connId;
				Setting = newSetting;
			}

			public int ConnectionId { get; private set; }


			public CommEnableDisableEnum Setting { get; private set; }

		}

		#endregion

		#region Nested type: CtrlLocalRemoteSwitchArgs

		public class ControlLocalRemoteSwitchArgs : EventArgs
		{

			public ControlLocalRemoteSwitchArgs(int connId, ControlLocalRemoteEnum newSetting)
			{
				ConnectionId = connId;
				Setting = newSetting;
			}

			public int ConnectionId { get; private set; }

			public ControlLocalRemoteEnum Setting { get; private set; }
		}

		#endregion

		#region Nested type: CtrlOfflineOnlineSwitchArgs

		public class ControlOfflineOnlineSwitchArgs : EventArgs
		{
			public ControlOfflineOnlineSwitchArgs(int connId, ControlOfflineOnlineEnum newSetting)
			{
				ConnectionId = connId;
				Setting = newSetting;
			}

			public int ConnectionId { get; private set; }

			public ControlOfflineOnlineEnum Setting { get; private set; }

		}

		#endregion

		#region Nested type: GEMStateChangeArgs

		public class GEMStateChangeArgs : EventArgs
		{
			public GEMStateChangeArgs(int connection, StateMachine statemachine, int state)
			{
				Connection = connection;
				StateMachine = statemachine;
				State = state;
			}

			public int Connection { get; private set; }


			public StateMachine StateMachine { get; private set; }


			public int State { get; private set; }

		}

		#endregion

		#region Nested type: GetValueCallbackArgs

		public class GetValueCallbackArgs : EventArgs
		{
			public GetValueCallbackArgs(string variableName, int variableID)
			{
				VariableName = variableName;
				VariableID = variableID;
			}

			public string VariableName { get; private set; }


			public int VariableID { get; private set; }


			public CxValue VariableValue { get; set; }
		}

		#endregion

		#region Nested type: LogArgs

		public class LogArgs : EventArgs
		{

			public LogArgs(string textMessage)
			{
				TextMessage = textMessage;
			}

			public string TextMessage { get; private set; }

		}

		#endregion

		#region Nested type: RemoteCommandArgs

		public class RemoteCommandArgs : EventArgs
		{
			public RemoteCommandArgs(
				int connection,
				string command,
				CxValueObject argumentNames,
				CxValueObject argumentValues,
				ref CommandResults commandResult)
			{
				Connection = connection;
				Command = command;
				ArgumentNames = argumentNames;
				ArgumentValues = argumentValues;
				CommandResult = commandResult;
			}

			public int Connection { get; private set; }
			public string Command { get; private set; }
			public CxValueObject ArgumentNames { get; private set; }
			public CxValueObject ArgumentValues { get; private set; }
			public CommandResults CommandResult { get; set; }
		}

		#endregion

		#region Nested type: TerminalServiceArgs

		public class TerminalServiceArgs : EventArgs
		{

			public TerminalServiceArgs(string textMessage)
			{
				TextMessage = textMessage;
			}

			public string TextMessage { get; private set; }
		}

		#endregion

		#endregion

		#region Enumerations

		public enum CommEnableDisableEnum
		{
			Disable = 0,

			Enable = 256
		}

		public enum ControlLocalRemoteEnum
		{
			Local = 0,

			Remote = 1
		}


		public enum ControlOfflineOnlineEnum
		{
			Offline = 0,

			Online = 1
		}

		#endregion


		/// <summary>
		///    Generic exception handling function to demonstrate how to get the COM HRESULT from
		///    a COM exception, and to handle the different types of COM exceptions that can be
		///    thrown.
		/// </summary>
		/// <param name="message"></param>
		/// <param name="e"></param>
		public void HandleException(string message, Exception e)
		{
			// Check for a general COM exception 
			var myCOMException = e as COMException;
			if (null != myCOMException)
			{
				Log("COMException: " + message + " HRESULT: 0x" 
					 + myCOMException.ErrorCode.ToString("X")
					 + Environment.NewLine
					 + myCOMException.Message);
			}
			else
			{
				Log("Exception: " + message + Environment.NewLine + e.Message);
			}
		}

		/// <summary>
		///    Initialize the interfaces and CIMConnect
		/// </summary>
		public void Initialize(string EPJFile)
		{
			try
			{
				NotifyInitializationStart();

				_CxEMService = new CxEMService();

				// Wait for the service to be in a RUNNING state (15 s in this sample) 
				// Optional if you want to have windows service level monitoring
				//				var myServiceController = new ServiceController { ServiceName = "EMService" };
				//				var myTimeSpan = new TimeSpan(0, 0, 0, 15);
				//				myServiceController.WaitForStatus(ServiceControllerStatus.Running, myTimeSpan);

				// Wait for EMService (in case it just started up)
				int result;
				Log("CxEMService.WaitOnEquipmentReady");
				_CxEMService.WaitOnEquipmentReady(0, _StartupWaitTimeInMS, out result);

				// Establish a connection to CIMConnect (EMService)
				var appID = 0;
				Log("CxEMService.Connect");
				_CxEMService.Connect(0, out _CxClientClerk, out appID);

				// Get the Client Tool interface
				_ICxClientTool = (ICxClientTool)_CxClientClerk;

				// Make sure that the correct EPJ file is loaded. 
				string currentProject;
				Log("ICxClientTool.GetCurrentProject");
				_ICxClientTool.GetCurrentProject(out currentProject);

				// Get the default project directory from the currentProject
				var desiredProject = currentProject.Substring(0, currentProject.LastIndexOf("\\")) + "\\" + EPJFile;

				if (File.Exists(desiredProject) == false)
				{
					throw new ApplicationException("The EPJ project file = " + desiredProject + " does not exist.");
				}

				if (currentProject != desiredProject)
				{
					Log("ICxClientTool.LoadProject " + EPJFile);
					_ICxClientTool.LoadProject(EPJFile);

					// Wait for the project load to completely. 
					Log("CxEMService.WaitOnEquipmentReady");
					_CxEMService.WaitOnEquipmentReady(0, _StartupWaitTimeInMS, out result);
				}

				// Make sure that the default EPJ file is correct.
				// This makes startup faster next time. 
				string defaultProject;
				Log("ICxClientTool.GetDefaultProjectFile");
				_ICxClientTool.GetDefaultProjectFile(out defaultProject);
				if (defaultProject != desiredProject)
				{
					Log("ICxClientTool.SetDefaultProjectFile");
					_ICxClientTool.SetDefaultProjectFile(desiredProject);
				}

				// Give CIMConnect your application's name.
				_CxClientClerk.appName = "Getting Started Biometric Application";

				// Register to receive Terminal Services
				Log("_CxClientClerk.RegisterTerminalMsgHandler");
				_CxClientClerk.RegisterTerminalMsgHandler(this);

				// Setup for each connection
				for (var host = 1; host <= _NumberOfConnections; host++)
				{
					// Register to receive State Machine changes
					Log("_CxClientClerk.RegisterStateMachineHandler " + host);
					_CxClientClerk.RegisterStateMachineHandler(host, this);

					// Get the communication state
					Int64 currentValue = 0;
					GetVariableValue(host, "CommState", ref currentValue);
					_CommunicationState[host - 1] = (CommunicationState)currentValue;

					// Get the control state
					GetVariableValue(host, "CONTROLSTATE", ref currentValue);
					_ControlState[host - 1] = (ControlState)currentValue;

					// Get the control state
					GetVariableValue(host, "SpoolState", ref currentValue);
					_SpoolingState[host - 1] = (SpoolingState)currentValue;
				}

				// Register for ValueChange for CommEnableSwitch
				_ICxClientTool.GetVarID(1, "CommEnableSwitch", out _VID_CommEnableSwitch);
				Log("_CxClientClerk.RegisterByteStreamValueChangedHandler(1, CommEnableSwitch): " + _VID_CommEnableSwitch.ToString());
				_CxClientClerk.RegisterByteStreamValueChangedHandler(1, _VID_CommEnableSwitch, this);

				// Register for ValueChange for CtrlOnlineSwitch
				_ICxClientTool.GetVarID(1, "CtrlOnlineSwitch", out _VID_CtrlOnlineSwitch);
				Log("_CxClientClerk.RegisterByteStreamValueChangedHandler(1, CtrlOnlineSwitch): " + _VID_CtrlOnlineSwitch.ToString());
				_CxClientClerk.RegisterByteStreamValueChangedHandler(1, _VID_CtrlOnlineSwitch, this);

				// Register for ValueChange for ControlStateSwitch
				_ICxClientTool.GetVarID(1, "ControlStateSwitch", out _VID_ControlStateSwitch);
				Log("_CxClientClerk.RegisterByteStreamValueChangedHandler(1, ControlStateSwitch)" + _VID_ControlStateSwitch.ToString());
				_CxClientClerk.RegisterByteStreamValueChangedHandler(1, _VID_ControlStateSwitch, this);

				// Register to receive Process State Machine changes
				Log("_CxClientClerk.RegisterStateMachineHandler 0");
				_CxClientClerk.RegisterStateMachineHandler(0, this);
			}
			catch (Exception e)
			{
				HandleException("Exception initializing communication to CIMConnect", e);
				NotifyInitializationComplete(false, "Exception initializing CIMConnect: " + e.Message);

				if (_CxEMService != null)
				{
					_CxEMService.StopService();
				}

				throw;
			}
		}

		/// <summary>
		///    Final initialization of the GEM interface. Until this is called, communication
		///    with the GEM Host is not permitted. This allows for equipment initialization to
		///    be completed as well as initialization of CIMConnect and this application.
		/// </summary>
		public void InitializeFinal()
		{
			try
			{
				Int64 defaultCommState = 0;
				for (var i = 1; i <= _NumberOfConnections; i++)
				{
					GetVariableValue(i, "DefaultCommState", ref defaultCommState);
					if (256 == defaultCommState)
					{
						_CxClientClerk.EnableComm(i);
					}
				}
				_initialized = true;
			}
			catch (Exception e)
			{
				HandleException("Exception in the final initialization of CIMConnect", e);
				NotifyInitializationComplete(false, "Exception in final initialization of CIMConnect: " + e.Message);
				throw;
			}

			NotifyInitializationComplete(true, "");
		}

		/// <summary>
		/// </summary>
		public void Shutdown()
		{
			try
			{
				_initialized = false;
				if (null != _CxClientClerk)
				{
					for (var host = 1; host <= _NumberOfConnections; host++)
					{
						Log("_CxClientClerk.DisableComm: " + host);
						_CxClientClerk.DisableComm(host);

						Log("_CxClientClerk.UnregisterStateMachineHandler: " + host);
						_CxClientClerk.UnregisterStateMachineHandler(1, this);
					}

					// Unregister ValueChange for CommEnableSwitch 
					Log("_CxClientClerk.UnregisterByteStreamValueChangedHandler(1, CommEnableSwitch)");
					_CxClientClerk.UnregisterByteStreamValueChangedHandler(1, _VID_CommEnableSwitch);

					// Unregister ValueChange for CtrlOnlineSwitch 
					Log("_CxClientClerk.UnregisterByteStreamValueChangedHandler(1, _CtrlOnlineSwitch)");
					_CxClientClerk.UnregisterByteStreamValueChangedHandler(1, _VID_CtrlOnlineSwitch);

					// Unregister ValueChange for ControlStateSwitch 
					Log("_CxClientClerk.UnregisterByteStreamValueChangedHandler(1, ControlStateSwitch)");
					_CxClientClerk.UnregisterByteStreamValueChangedHandler(1, _VID_ControlStateSwitch);

					Log("_CxClientClerk.UnregisterStateMachineHandler 0");
					_CxClientClerk.UnregisterStateMachineHandler(0, this);

					Log("_CxClientClerk.UnregisterTerminalMsgHandler");
					_CxClientClerk.UnregisterTerminalMsgHandler(this);
				}
				if (null != _CxEMService)
				{
					Log("CxEMService.StopService");
					_CxEMService.StopService();
				}
			}
			catch (Exception e)
			{
				HandleException("Exception shutting down CMyCIMConnect", e);
			}
		}

		/// <summary>
		///    Convert a SECS-II message ID into the stream number and function number.
		/// </summary>
		/// <param name="messageID">SECS-II message ID</param>
		/// <param name="s">stream number</param>
		/// <param name="f">function number</param>
		public void MsgIDtoSF(int messageID, out int s, out int f)
		{
			s = (messageID >> 8) & 0xff;
			f = messageID & 0xff;
		}

		/// <summary>
		///    Convert a stream number and function number into a SECS-II message ID;
		/// </summary>
		/// <param name="s">stream </param>
		/// <param name="f">function</param>
		/// <returns>SECS-II message ID</returns>
		public int SFtoMsgID(int s, int f)
		{
			return (s << 8) | f;
		}

		public CxValue CreateFromSMN(string xml)
		{
			return CxValue.CreateFromSMN(XmlReader.Create(new StringReader(xml)));
		}

		/// <summary>
		///    Send a terminal service text message to the host.
		/// </summary>
		/// <param name="connection"></param>
		/// <param name="message">The text message</param>
		public void SendTerminalMessage(int connection, string message)
		{
			try
			{
				Log("_CxClientClerk.SendTerminalMsg");
				_CxClientClerk.SendTerminalMsg(connection, 0, message);
			}
			catch (Exception e)
			{
				HandleException("Exception calling ::SendTerminalMsg", e);
			}
		}

		/// <summary>
		///    Send a terminal service acknowledge event to the
		///    host. Send to host 1 by default.
		/// </summary>
		/// <param name="connection"></param>
		public void SendTerminalAcknowledge(int connection)
		{
			try
			{
				var name = "MessageRecognition";
				Log("_CxClientClerk.TriggerWellKnownEvent " + name);
				_CxClientClerk.TriggerWellKnownEvent(connection, name);
			}
			catch (Exception e)
			{
				HandleException("Exception calling ::TriggerWellKnownEvent", e);
			}
		}

		public void ProcessingStarted()
		{
			try
			{
				var name = "ProcessingStarted";
				Log("_CxClientClerk.TriggerWellKnownEvent " + name);
				_CxClientClerk.TriggerWellKnownEvent(1, name);
			}
			catch (Exception e)
			{
				HandleException("Exception calling ::TriggerWellKnownEvent", e);
			}
		}

		public void ProcessingStopped()
		{
			try
			{
				var name = "ProcessingStopped";
				Log("_CxClientClerk.TriggerWellKnownEvent " + name);
				_CxClientClerk.TriggerWellKnownEvent(1, name);
			}
			catch (Exception e)
			{
				HandleException("Exception calling ::TriggerWellKnownEvent", e);
			}
		}

		public void ProcessingCompleted()
		{
			try
			{
				var name = "ProcessingCompleted";
				Log("_CxClientClerk.TriggerWellKnownEvent " + name);
				_CxClientClerk.TriggerWellKnownEvent(1, name);
			}
			catch (Exception e)
			{
				HandleException("Exception calling ::TriggerWellKnownEvent", e);
			}
		}


		/// <summary>
		/// Define a collection event in the GEM interface. This adds
		/// the collection event to the GEM interface, but not to the EPJ
		/// file. This should be done only before the GEM interface is enabled. 
		/// This function assumes that the name and ID are unique, and not already in 
		/// use by an existing definition previously added to CIMConnect either through
		/// the EPJ file or through this API function. 
		/// </summary>
		/// <param name="name"></param>
		/// <param name="id"></param>
		/// <param name="description"></param>
		/// <param name="DVList"></param>
		public void DefineCollectionEvent(string name, int id, string description, string[] DVList)
		{
			try
			{
				_ICxClientTool.CreateEvent(0, id, name, ref description);
			}
			catch (Exception e)
			{
				HandleException("Exception CreateEvent " + name, e);
			}
			
			if(DVList.GetLength(0) > 0)
            {
				try
				{
					VALUELib.CxValueObject dvNameList = _CxClientClerk.CreateValueObject();
					foreach (var dv in DVList)
					{
						dvNameList.AppendValueAscii(0, dv);
					}
					object dvNameListObject;
					dvNameList.CopyToByteStream(0, 0, out dvNameListObject);

					_ICxClientTool.SetEventAssociatedDVNameList(0, id, dvNameListObject);
				}
				catch (Exception e)
				{
					HandleException("Exception SetEventAssociatedDVNameList " + name, e);
				}
			}
		}

		/// <summary>
		/// Define a status variable, data variable or equipment constant in the GEM 
		/// interface. This adds the variable to the GEM interface, but not to the EPJ
		/// file. This should be done only before the GEM interface is enabled. 
		/// This function assumes that the name and ID are unique, and not already in 
		/// use by an existing definition previously added to CIMConnect either through
		/// the EPJ file or through this API function. 
		/// </summary>
		/// <param name="name"></param>
		/// <param name="id"></param>
		/// <param name="description"></param>
		/// <param name="variableType"></param>
		/// <param name="valueType"></param>
		/// <param name="eventID"></param>
		/// <param name="unit"></param>
		/// <param name="maximum"></param>
		/// <param name="minimum"></param>
		/// <param name="initial"></param>
		/// <param name="isPrivate"></param>
		/// <param name="isPersitent"></param>
		public void DefineVariable(string name, int id, string description,	VarType variableType, ValueType valueType,	int eventID, string unit, CxValueObject maximum, CxValueObject minimum, CxValueObject initial,
			bool isPrivate, bool isPersitent)
		{
			try
			{
				_ICxClientTool.CreateVariable(0, id, variableType, valueType, eventID,
					name, ref description, ref unit, maximum, minimum, initial,
					ref isPrivate, ref isPersitent);
			}
			catch (Exception e)
			{
				HandleException("Exception CreateVariable " + name, e);
			}
		}

		/// <summary>
		/// Define an alarm in the GEM interface. This adds
		/// the alarm to the GEM interface, but not to the EPJ
		/// file. This should be done only before the GEM interface is enabled. 
		/// This function assumes that the name and ID are unique, and not already in 
		/// use by an existing definition previously added to CIMConnect either through
		/// the EPJ file or through this API function. 
		/// </summary>
		/// <param name="name"></param>
		/// <param name="id"></param>
		/// <param name="description"></param>
		/// <param name="defaultText"></param>
		/// <param name="setCEID"></param>
		/// <param name="clearCEID"></param>
		/// <param name="alarmCode"></param>
		public void DefineAlarm(string name, int id, string description, string defaultText, int setCEID, int clearCEID, int alarmCode)
		{
			try
			{
				_ICxClientTool.CreateAlarm2(id, defaultText, ref name, ref description,	setCEID, clearCEID, alarmCode);
			}
			catch (Exception e)
			{
				HandleException("Exception CreateAlarm2 " + name, e);
			}
		}

		/// <summary>
		/// Send a collection event to all host connections
		/// </summary>
		/// <param name="eventName">Collection Event name</param>
		public void SendCollectionEvent(string eventName)
		{
			try
			{
				var eventID = -1;
				Log("_CxClientClerk.TriggerEvent " + eventName);
				_CxClientClerk.TriggerEvent(0, ref eventID, ref eventName);
			}
			catch (Exception e)
			{
				HandleException("Exception SendCollectionEvent " + eventName, e);
			}
		}

		/// <summary>
		/// Trigger a collection event while updating the values for variables. This is
		/// ideal for updating data variables associated with the event.
		/// </summary>
		/// <param name="connection"></param>
		/// <param name="variableNames"></param>
		/// <param name="variableValues"></param>
		/// <param name="eventName"></param>
		public void SendCollectionEventWithData(int connection,	string[] variableNames,	CxValue[] variableValues, string eventName)
		{
			try
			{
				var eventID = -1;
				object varIDs = null;
				object varNames = variableNames;
				
				var varValues = new object[variableValues.Length];
				
				for (var i = 0; i < variableValues.Length; i++)
				{
					varValues[i] = variableValues[i].ToByteBuffer();
				}
				object results;

				Log("_CxClientClerk.SetValuesTriggerEvent " + eventName);
				
				_CxClientClerk.SetValuesTriggerEvent
				(
					connection,
					VarType.varANY,
					ref varIDs,
					ref varNames,
					varValues,
					ref eventID,
					ref eventName,
					out results
				);
			}
			catch (Exception e)
			{
				HandleException("Exception in SendCollectionEventWithData " + eventName, e);
			}
		}

		/// <summary>
		/// Trigger a collection event while updating the values for variables. This is
		/// ideal for updating data variables associated with the event.
		/// </summary>
		/// <param name="connection"></param>
		/// <param name="variableNames"></param>
		/// <param name="variableValues"></param>
		/// <param CEID="eventId"></param>
		public void SendCollectionEventWithData(int connection,	string[] variableNames,	CxValue[] variableValues, int CEID)
		{
			try
			{
				var eventID = CEID;
				object varIDs = null;
				object varNames = variableNames;
				var varValues = new object[variableValues.Length];
				string eventName = String.Empty; ;

				for (var i = 0; i < variableValues.Length; i++)
				{
					varValues[i] = variableValues[i].ToByteBuffer();
				}
				object results;

				Log("_CxClientClerk.SetValuesTriggerEvent " + CEID.ToString());

				_CxClientClerk.SetValuesTriggerEvent(
					connection,
					VarType.varANY,
					ref varIDs,
					ref varNames,
					varValues,
					ref eventID,
					ref eventName,
					out results);
			}
			catch (Exception e)
			{
				HandleException("Exception in SendCollectionEventWithData " + CEID.ToString(), e);
			}
        }



		/// <summary>
		/// 
		/// </summary>
		/// <param name="alarmName"></param>
		/// <param name="text"></param>
		/// <param name="connId"></param>
		/// <param name="varNames"></param>
		/// <param name="varVals"></param>
		public void SetAlarm(string alarmName, string text, int connId = 0,
										  string[] varNames = null, CxValue[] varVals = null)
		{
			object results;
			object varIDs = null;
			object objVarNames = varNames == null ? new string[0] : varNames;
			var objVarVals = varVals == null ? new object[0] : new object[varVals.Length];

			if (varNames != null)
			{
				for (var i = 0; i < varVals.Length; i++)
				{
					objVarVals[i] = varVals[i].ToByteBuffer();
				}
			}

			_CxClientClerk.SetAlarm2(-1, alarmName, 0, text, connId, ref varIDs, ref objVarNames, objVarVals, out results);
		}

		/// <summary>
		/// 
		/// </summary>
		/// <param name="alarmName"></param>
		/// <param name="text"></param>
		/// <param name="connId"></param>
		/// <param name="varNames"></param>
		/// <param name="varVals"></param>
		public void ClearAlarm(string alarmName, string text, int connId = 0,  string[] varNames = null, CxValue[] varVals = null)
		{
			object results;
			object varIDs = null;
			object objVarNames = varNames == null ? new string[0] : varNames;
			var objVarVals = varVals == null ? new object[0] : new object[varVals.Length];

			if (varNames != null)
			{
				for (var i = 0; i < varVals.Length; i++)
				{
					objVarVals[i] = varVals[i].ToByteBuffer();
				}
			}

			try
            {
				_CxClientClerk.ClearAlarm2(-1, alarmName, 0, text, connId, ref varIDs, ref objVarNames, objVarVals, out results);
			}
			catch (Exception e)
			{
				HandleException("Exception ClearAlarm " + alarmName, e);
			}
		}

		/// <summary>
		/// Change an alarm to the SET state.
		/// </summary>
		/// <param name="name">Alarm name</param>
		/// <param name="text">Text about the alarm.</param>
		public void AlarmSET(string name, string text)
		{
			try
			{
				var id = -1;
				if (text.Length == 0)
				{
					Log("_CxClientClerk.SetAlarm " + name);
					_CxClientClerk.SetAlarm(ref id, ref name);
				}
				else
				{
					Log("_CxClientClerk.SetAlarmAndText " + name);
					_CxClientClerk.SetAlarmAndText(id, name, text);
				}
			}
			catch (Exception e)
			{
				HandleException("Exception AlarmSET " + name, e);
			}
		}

		/// <summary>
		/// Change an alarm to the SET state.
		/// </summary>
		/// <param id="alarm id">Alarm name</param>
		public void AlarmSET(int id, string text)
		{
			try
			{
				Log("_CxClientClerk.SetAlarm (id) " + id.ToString());
				_CxClientClerk.SetAlarm(ref id, ref text);
			}
			catch (Exception e)
			{
				HandleException("Exception AlarmSET (id) " + id.ToString(), e);
			}
		}


		/// <summary>
		/// Change the state of an alarm to CLEAR
		/// </summary>
		/// <param name="name">Alarm name</param>
		public void AlarmCLEAR(string name)
		{
			try
			{
				var id = -1;
				Log("_CxClientClerk.ClearAlarm " + name);
				_CxClientClerk.ClearAlarm(ref id, ref name);
			}
			catch (Exception e)
			{
				HandleException("Exception AlarmCLEAR " + name, e);
			}
		}

		/// <summary>
		/// Change the state of an alarm to CLEAR
		/// </summary>
		/// <param id="alarm id">Alarm name</param>
		public void AlarmCLEAR(int id)
		{
			string name = String.Empty;

			try
			{
				Log("_CxClientClerk.ClearAlarm (id) " + id.ToString());
				_CxClientClerk.ClearAlarm(ref id, ref name);
			}
			catch (Exception e)
			{
				HandleException("Exception AlarmCLEAR (id) " + id.ToString(), e);
			}
		}


		/// <summary>
		/// Cache the value of a variable, any type, in CIMConnect.
		/// </summary>
		/// <param name="connection">CIMConnect connection number. 0 if a common variable.</param>
		/// <param name="name">variable name</param>
		/// <param name="myCxValue">The variable's new value. </param>
		/// <returns></returns>
		public void CacheVariable(int connection, string name, CxValue myCxValue)
		{
			try
			{
				VariableResults results;
				var id = -1;
				Log("_CxClientClerk.SetValueFromByteBuffer " + name);
				_CxClientClerk.SetValueFromByteBuffer(
																	 connection,
					VarType.varANY,
					ref id,
					ref name,
					myCxValue.ToByteBuffer(),
					out results);
			}
			catch (Exception e)
			{
				var sValue = string.Empty;
				if (myCxValue != null)
				{
					sValue = myCxValue.ToString();
				}
				HandleException("Exception calling ::SetValueFromByteBuffer " + name + ": " + sValue, e);
			}
		}

		/// <summary>
		///    Cache the value of a variable of type Unsigned 4 Byte Integer
		///    The ICxValue interface is used to allow setting the value object
		///    directly as an unsigned integer since the CxValueObject::SetValueU4
		///    function requires the value to be specified as a signed integer.
		/// </summary>
		/// <param name="name">variable name</param>
		/// <param name="value">variable value</param>
		/// <returns></returns>
		public void CacheVariable(string name, uint value)
		{
			try
			{
				CxValue myCxValue = new U4Value(value);
				CacheVariable(0, name, myCxValue);
			}
			catch (Exception e)
			{
				HandleException("Exception CacheVariable " + name + ": " + value, e);
			}
		}

		/// <summary>
		///    Cache the value of a variable of type 8 byte floating point
		/// </summary>
		/// <param name="name">variable name</param>
		/// <param name="value">variable value</param>
		/// <returns></returns>
		public void CacheVariable(string name, double value)
		{
			try
			{
				CxValue myCxValue = new F8Value(value);
				CacheVariable(0, name, myCxValue);
			}
			catch (Exception e)
			{
				HandleException("Exception calling CacheVariable " + name + ": " + value, e);
			}
		}

		/// <summary>
		///    Cache the value of a variable of type ASCII
		/// </summary>
		/// <param name="name">variable name</param>
		/// <param name="value">variable value</param>
		/// <returns></returns>
		public void CacheVariable(string name, string value)
		{
			try
			{
				CxValue myCxValue = new AValue(value);
				CacheVariable(0, name, myCxValue);
			}
			catch (Exception e)
			{
				HandleException("Exception calling CacheVariable " + name + ": " + value, e);
			}
		}

		/// <summary>
		///    Cache the value of multiple variables, any type, in CIMConnect.
		/// </summary>
		/// <param name="connection">CIMConnect connection number. 0 if a common variable.</param>
		/// <param name="variableNames"></param>
		/// <param name="variableValues"></param>
		public void CacheVariables(int connection, string[] variableNames, CxValue[] variableValues)
		{
			try
			{
				object varIDs = null;
				object varNames = variableNames;
				var varValues = new object[variableValues.Length];
				for (var i = 0; i < variableValues.Length; i++)
				{
					varValues[i] = variableValues[i].ToByteBuffer();
				}
				object results;

				Log("_CxClientClerk.SetValuesFromByteBuffer ");
				_CxClientClerk.SetValuesFromByteBuffer(connection,
					VarType.varANY,
					ref varIDs,
					ref varNames,
					varValues,
					out results);
			}
			catch (Exception e)
			{
				HandleException("Exception in CacheVariables ", e);
			}
		}

		/// <summary>
		///    Get the value of a Status Variable, Data Variable or Equipment Constant from CIMConnect. The value
		///    must be type integer (I1, I2, I4, U1, U2, U4, Binary), and not be of type array for this function
		///    to work.
		/// </summary>
		/// <param name="connection"></param>
		/// <param name="name"></param>
		/// <param name="value"></param>
		public bool GetVariableValue(int connection, string name, ref Int64 value)
		{
			try
			{
				VariableResults result;
				var id = -1;
				object oValue;

				Log("_CxClientClerk.GetValueToByteBuffer " + name);
				_CxClientClerk.GetValueToByteBuffer(connection, ref id, ref name, out oValue, out result);

				var myCxValue = CxValue.CreateFromByteBuffer((byte[])oValue);
				value = CxValue.GetValueInt(myCxValue);
			}
			catch (Exception e)
			{
				HandleException("Exception in GetVariableValue " + name, e);
			}
			return true;
		}

		/// <summary>
		///    Change the GEM Control state to ONLINE or OFFLINE
		/// </summary>
		/// <param name="connection">Connection #</param>
		/// <param name="state">if true, go ONLINE. if false, go OFFLINE</param>
		public void GEMStateControlStateOnline(int connection, bool state)
		{
			try
			{
				uint switchValue = 0;
				if (state)
				{
					Log("_CxClientClerk.GoOnline " + connection);
					_CxClientClerk.GoOnline(connection);

					switchValue = (uint)ControlOfflineOnlineEnum.Online;
				}
				else
				{
					Log("_CxClientClerk.GoOffline " + connection);
					_CxClientClerk.GoOffline(connection);
					switchValue = (uint)ControlOfflineOnlineEnum.Offline;
				}

				VariableResults result;
				var newValue = new U4Value(switchValue);
				int id = -1;
				string name = "CtrlOnlineSwitch";
				Log("_CxClientClerk.SetValueFromByteBuffer " + connection + "CtrlOnlineSwitch " + switchValue);
				_CxClientClerk.SetValueFromByteBuffer(connection, VarType.varANY, ref id, ref name, newValue.ToByteBuffer(), out result);
			}
			catch (Exception e)
			{
				HandleException("Exception in GEMStateControlStateOnline ", e);
			}
		}

		/// <summary>
		///    Change the GEM Control State to REMOTE or LOCAL
		/// </summary>
		/// <param name="connection">Connection #</param>
		/// <param name="state">if true, go REMOTE. if false, go LOCAL</param>
		public void GEMStateControlStateRemote(int connection, bool state)
		{
			try
			{
				if (state)
				{
					Log("_CxClientClerk.GoRemote " + connection);
					_CxClientClerk.GoRemote(connection);
				}
				else
				{
					Log("_CxClientClerk.GoLocal " + connection);
					_CxClientClerk.GoLocal(connection);
				}
			}
			catch (Exception e)
			{
				HandleException("Exception in GEMStateControlStateRemote ", e);
			}
		}

		/// <summary>
		///    Change the GEM Communication State to ENABLE or DISABLE communication
		/// </summary>
		/// <param name="connection">Connection #</param>
		/// <param name="state">if true, ENABLE. if false, DISABLE</param>
		public void GEMStateCommunicationStateEnable(int connection, bool state)
		{
			try
			{
				if (state)
				{
					Log("_CxClientClerk.EnableComm " + connection);
					_CxClientClerk.EnableComm(connection);
				}
				else
				{
					Log("_CxClientClerk.DisableComm " + connection);
					_CxClientClerk.DisableComm(connection);
				}
			}
			catch (Exception e)
			{
				HandleException("Exception in GEMStateCommunicationStateEnable ", e);
			}
		}



		#region ICxEMAppCallback Members

		void ICxEMAppCallback.AlarmChanged(int alarmID, int state)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.AsyncMsgError(int connectionID, int msgID, int TID, int ErrorID, int lExtra)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.CallbackReplaced(int connectionID, Callbacks callback, int itemID, string itemName) { }

		/// <summary>
		/// </summary>
		/// <param name="connectionID"></param>
		/// <param name="command"></param>
		/// <param name="paramNames"></param>
		/// <param name="paramValues"></param>
		/// <param name="cmdResult"></param>
		void ICxEMAppCallback.CommandCalled(
			int connectionID,
			string command,
			CxValueObject paramNames,
			CxValueObject paramValues,
			ref CommandResults cmdResult)
		{
			try
			{
				Log("CommandCalled " + connectionID + " command: " + command);

				// By default, assume it will be rejected. 
				cmdResult = CommandResults.cmdCannotPerform;

				// Pass the command on to the Form for processing
				RemoteCommand(connectionID, command, paramNames, paramValues, ref cmdResult);
				Log("Remote Command Handled: " + command + " result = " + cmdResult);
			}
			catch (Exception ex)
			{
				Log("Remote Command " + command + " rejected because an exception occurred: " + ex.Message);
				cmdResult = CommandResults.cmdRejected;
			}
		}

		void ICxEMAppCallback.EventTriggered(int connectionID, int eventID)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.GetValue(int connectionID, int varid, string name, CxValueObject value)
		{
			value.SetValueU4(0,0,99);
		}

		/// <summary>
		/// </summary>
		/// <param name="connectionID"></param>
		/// <param name="varid"></param>
		/// <param name="name"></param>
		/// <param name="pValueBuffer"></param>
		void ICxEMAppCallback.GetValueToByteBuffer(int connectionID, int varid, string name, ref object pValueBuffer)
		{
			try
			{
				pValueBuffer = GetValueCallback(name, varid);
			}
			catch (Exception ex)
			{
				Log(ex.Message);
			}
		}

		void ICxEMAppCallback.GetValues(int connectionID, object ids, object names, object values)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.GetValuesToByteBuffer(int connectionID, object varIDs, object names, ref object pValueBuffers)
		{
			try
			{
				int[] idList = (int[])varIDs;
				string[] nameList = (string[])names;

				object[] valueList = new object[idList.Length];
				for (int i = 0; i < idList.Length; i++)
				{
					// Query from the application one at a time, using the same delegate event as
					// the GetValueToByteBuffer. 
					valueList[i] = GetValueCallback(nameList[i], idList[i]);
				}
				pValueBuffers = valueList;
			}
			catch (Exception ex)
			{
				Log(ex.Message);
			}
		}

		void ICxEMAppCallback.HostPPLoadInqAck(int connectionID, RecipeGrant result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.HostPPSendAck(int connectionID, RecipeAck result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.HostTermMsgAck(int connectionID, TerminalMsgResults result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.MessageReceived(
			int connectionID,
			int msgID,
			CxValueObject msg,
			bool replyExpected,
			int TID,
			ref int replyMsgID,
			CxValueObject reply,
			ref MessageResults result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPDRequest(int connectionID, CxValueObject names)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPData(int connectionID, CxValueObject recipe, int result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPDataFile(int connectionID, string filename, int result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPDelete(int connectionID, CxValueObject names, ref RecipeAck result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPLoadInquire(int connectionID, string name, int length, ref RecipeGrant result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPRequest(int connectionID, string name, CxValueObject recipe, ref RecipeAck result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPSend(int connectionID, CxValueObject recipe, ref RecipeAck result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPSendFile(int connectionID, string filename, ref RecipeAck result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.PPSendFileVerify(int connectionID, string filename, ref RecipeAck result)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.SetValue(int connectionID, int varid, string name, CxValueObject value)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.SetValueFromByteBuffer(int connectionID, int varid, string name, object ValueBuffer)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.SetValues(int connectionID, object ids, object names, object values)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.SetValuesFromByteBuffer(int connectionID, object varIDs, object names, object ValueBuffers)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.StateChange(int connectionID, StateMachine machine, int state)
		{
			try
			{
				if (connectionID > _NumberOfConnections)
				{
					return;
				}

				switch (machine)
				{
					case StateMachine.smCommunications:
						_CommunicationState[connectionID - 1] = (CommunicationState)state;
						break;
					case StateMachine.smControl:
						_ControlState[connectionID - 1] = (ControlState)state;
						break;
					case StateMachine.smSpooling:
						_SpoolingState[connectionID - 1] = (SpoolingState)state;
						break;
				}
				GEMStateChange(connectionID, machine, state);
			}
			catch (Exception e)
			{
				Log(e.Message);
			}
		}

		void ICxEMAppCallback.TerminalMsgRcvd(
			int connectionID,
			int terminalID,
			CxValueObject lines,
			ref TerminalMsgResults result)
		{
			try
			{
				Log("EMSERVICELib.ICxEMAppCallback.TerminalMsgRcvd");
				if (result.Equals(null))
				{
					return;
				}

				result = TerminalMsgResults.tRejected;

				// Setup my string to display the terminal message
				var mymessage = "From host #" + connectionID + ":" + Environment.NewLine;
				string terminal;

				// Get the data type (expecting S10F3 or S10F5 format)
				ValueType datatype;
				lines.GetDataType(0, 0, out datatype);
				if (datatype == ValueType.valueANY)
				{
					Log("ERROR: MyCIMConnect: TerminalMsgRcvd GetDataType failed.");
					return;
				}
				switch (datatype)
				{
					case ValueType.L:
						{
							int numberLines;
							var tmpMessage = "";
							lines.ItemCount(0, out numberLines);
							for (var i = 0; i < numberLines; i++)
							{
								if (i > 0)
								{
									tmpMessage += Environment.NewLine;
								}
								lines.GetValueAscii(0, i + 1, out terminal);
								if (terminal != null)
								{
									tmpMessage += terminal;
								}
							}
							if (tmpMessage.Length == 0)
							{
								mymessage = ""; // empty text message from host. 
							}
							else
							{
								mymessage += tmpMessage;
							}
						}
						break;
					case ValueType.A:
						{
							lines.GetValueAscii(0, 0, out terminal);
							if (terminal == null)
							{
								mymessage = ""; // empty text message from host 
							}
							else
							{
								mymessage += terminal;
							}
						}
						break;
					default:
						{
							Log("ERROR: MyCIMConnectCallbacks: TerminalMsgRcvd Unexpected data type " + datatype);
							return;
						}
				}

				// Display the Terminal Service message
				TerminalService(mymessage);

				// Accept the Terminal Service message
				result = TerminalMsgResults.tAccepted;
			}
			catch (Exception e)
			{
				Log(e.Message);
			}
		}

		void ICxEMAppCallback.ValueChanged(int connectionID, int varid, string name, CxValueObject value)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.ValuesChanged(int connectionID, object ids, object names, object values)
		{
			throw new NotImplementedException();
		}

		void ICxEMAppCallback.VerifyValue(
			int connectionID,
			int varid,
			string name,
			CxValueObject value,
			ref VerifyValueResults result)
		{
			try
			{
				result = VerifyValueCallback(name, varid, value);
			}
			catch (Exception e)
			{
				Log(e.Message);
			}
		}

		void ICxEMAppCallback.VerifyValues(int connectionID, object ids, object names, object values, ref object results)
		{
			throw new NotImplementedException();
		}

		#endregion


		void ICxEMAppCallback6.ValueChanged(int connectionID, int variableID, string variableName, object valueBuffer)
		{
			if (variableID == _VID_CommEnableSwitch)
			{
				U4Value value = (U4Value)CxValue.CreateFromByteBuffer((byte[])valueBuffer);
				CommEnableDisableSwitchChange(connectionID, (CommEnableDisableEnum)value.Value);
			}
			else if (variableID == _VID_ControlStateSwitch)
			{
				U4Value value = (U4Value)CxValue.CreateFromByteBuffer((byte[])valueBuffer);
				CtrlLocalRemoteSwitchChange(connectionID, (ControlLocalRemoteEnum)value.Value);
			}
			else if (variableID == _VID_CtrlOnlineSwitch)
			{
				U4Value value = (U4Value)CxValue.CreateFromByteBuffer((byte[])valueBuffer);
				CtrlOfflineOnlineSwitchChange(connectionID, (ControlOfflineOnlineEnum)value.Value);
			}
			else
			{
				ValueChangedCallback(variableName, variableID, valueBuffer);
			}
		}
	}
}