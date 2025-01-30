winClaas = CreateObject("WinClass")
*****************************************
*loUser = winClaas.GetUserInfo(5)
*****************************************
*Local loAllUsers(1)
*winClaas.GetLoggedUsers(@loAllUsers)
*****************************************
*MessageBox(winClaas.GetUserInfoFromSessionID(5,"state"))
*****************************************
*Local loAllSessions(1)
*winClaas.GetAllSessions(@loAllSessions)
*****************************************
*MessageBox(winClaas.GetConsoleSessionID())
*****************************************
LOCAL loAllProcesses(1)
winClaas.GetSystemProcesses(@loAllProcesses, "vfp8.exe")
******************************************
*LOCAL aWindowHandles(1)
*MESSAGEBOX(winClaas.GetWindowByProcessID(19580, @aWindowHandles))


Set Step On 

DEFINE CLASS WinClass AS Custom

	Null 								= Chr(0) 	&&& Null Character
	ByteSize 							= 4			&&& Byte Size of the pointer
	Pointer								= ""		&&& Pointer Size
	#DEFINE WTS_CURRENT_SERVER_HANDLE 	0
	#DEFINE WTS_UserName 				5
	#DEFINE WTS_WinStationName			6
	#DEFINE WTS_DomainName				7
	
	&& WTSEnumerateSessions
	#Define PWTS_SESSION_INFO 			12			&&& Size of Array in Bytes
	#DEFINE WTS_Active					0 			&&& The session is currently active
	#DEFINE WTS_Connected				1			&&& The session is connected to the client
	#DEFINE WTS_ConnectQuery			2			&&& The session is in the process of connecting to the client
	#DEFINE WTS_Shadow 					3			&&& The session is in a shadowing state
	#DEFINE WTS_Disconnected			4			&&& The session is disconnected from the client
	#DEFINE WTS_Idle					5			&&& The session is currently idle
	#DEFINE WTS_Listen					6			&&& The session is listening for a connection
	#DEFINE WTS_Reset					7			&&& The session is in the process of resetting
	#DEFINE WTS_Down					8			&&& The session is currently down
	#DEFINE WTS_Init					9			&&& The session is initializing
	
	#DEFINE GWL_STYLE 					-16			&&& Type of Style [Visible or not Visible]
	#DEFINE WS_VISIBLE 					0x10000000 	&&& Windows Visible Style
	#DEFINE GW_HWNDNEXT 				2			&&& Next Windows

	
	PROCEDURE Init
		* Definir
		This.Pointer = Replicate(This.Null, This.ByteSize)
	ENDPROC

	PROCEDURE Destroy
		* Definir
	ENDPROC
	
*	PROCEDURE Error(nError, cMethod, nLine)
		* Definir
*	EndProc
	
	*********************************************************************************
	
	*** <summary>
	*** Returns an array with the windows information
	*** </summary>
	*** <param name="nProcessID">Process ID of the EXE</param>
	*** <param name="aWindowHandles">Array to sabe the informations</param>
	*** <remarks></remarks>
	FUNCTION GetWindowByProcessID
	    LPARAMETERS nProcessID, aWindowHandles
	        
		LOCAL hWnd, nWindowProcessID, nCurrentProcessID, lnX, lcCaption
		STORE 0 TO hWnd, nWindowProcessID, nCurrentProcessID, lnX
		lcCaption = ""
		
	    aWindowHandles[1] = 0  
	    hWnd = This.GetNextWindow(hWnd) 

	    DO WHILE hWnd > 0
	        
	        nWindowProcessID = This.GetWindowThreadProcessId(hWnd)

			IF nWindowProcessID <> nProcessID
	        	hWnd = This.GetNextWindow(hWnd)
	        	LOOP 
	        ENDIF
	        
	        && Check if windows is visible
	        IF NOT This.IsVisibleWindow(hWnd)
	        	hWnd = This.GetNextWindow(hWnd)
	        	LOOP 
	        ENDIF
	        
	        lcCaption = This.GetWindowText(hWnd)
	        
	        IF Empty(lcCaption)
	        	hWnd = This.GetNextWindow(hWnd)
	        	LOOP 
	        ENDIF
	        
	        lnX = lnX + 1
			DIMENSION aWindowHandles[lnX]
	        aWindowHandles[lnX] = CREATEOBJECT("empty")
	        ADDPROPERTY(aWindowHandles[lnX],"Handle",hWnd)
	        ADDPROPERTY(aWindowHandles[lnX],"Caption", lcCaption)
	        hWnd = This.GetNextWindow(hWnd)
	    ENDDO
	    
	    RETURN aWindowHandles
	ENDFUNC

	*** <summary>
	*** Returns Next Window
	*** </summary>
	*** <param name="hWndCurrent">Current window handle</param>
	*** <remarks></remarks>
	FUNCTION GetNextWindow
	    LPARAMETERS hWndCurrent
	    
	    LOCAL lnHandle
	    lnHandle = 0
	    
	    IF hWndCurrent = 0
	        lnHandle = This.FindWindow(NULL, NULL)  && Start with the first window
	    ELSE
	        lnHandle = This.GetWindow(hWndCurrent, GW_HWNDNEXT)
	    ENDIF
	    
	    RETURN lnHandle
	ENDFUNC

	*** <summary>
	*** Checks if the window is Visible
	*** </summary>
	*** <param name="pln_Handle">Window Handle</param>
	*** <remarks></remarks>
	FUNCTION IsVisibleWindow
		LPARAMETERS pln_Handle
		
		LOCAL lnHandle
		lnHandle = pln_Handle
		RETURN BITAND(This.GetWindowLong(lnHandle, GWL_STYLE), WS_VISIBLE) = WS_VISIBLE
	ENDFUNC
	
	*** <summary>
	*** Get System Processes 
	*** </summary>
	*** <param name="plA_Processes">Pointer to Array</param>
	*** <param name="plc_Name">Process name</param>
	*** <remarks></remarks>
	FUNCTION GetSystemProcesses
		LPARAMETERS plA_Processes, plc_Name
		
		LOCAL loLocator, loWMI, lnX, lcName, plcWhere, LOCCPWRSHL
		lcName		= IIF(TYPE('plc_Name') == 'C', ALLTRIM(plc_Name),'')
		plcWhere	= ''
		loLocator 	= CREATEOBJECT('WBEMScripting.SWBEMLocator')
		loWMI		= loLocator.ConnectServer() 
		loWMI.Security_.ImpersonationLevel = 3  		&& Impersonate
		
		IF NOT EMPTY(lcName)
			plcWhere = TEXTMERGE("WHERE Name = '<<lcName>>'")
		ENDIF
		
		LOCCPWRSHL = ""
		TEXT TO LOCCPWRSHL NOSHOW TEXTMERGE PRETEXT 2 
			SELECT * FROM Win32_Process <<plcWhere>>
		ENDTEXT 
		loProcesses	= loWMI.ExecQuery(LOCCPWRSHL)
		
		IF loProcesses.Count <= 0
			RETURN loProcesses.Count 
		ENDIF
		
		Dimension plA_Processes[loProcesses.Count]

		lnX = 1
		FOR EACH loProcess in loProcesses
			plA_Processes[lnX] = CREATEOBJECT("empty")
			ADDPROPERTY(plA_Processes[lnX],"CreationClassName"			, loProcess.CreationClassName								)
			ADDPROPERTY(plA_Processes[lnX],"Caption"					, loProcess.Caption											)
			ADDPROPERTY(plA_Processes[lnX],"CommandLine"				, loProcess.CommandLine										)
			ADDPROPERTY(plA_Processes[lnX],"CreationDate"				, loProcess.CreationDate									)
			ADDPROPERTY(plA_Processes[lnX],"CSCreationClassName"		, loProcess.CSCreationClassName								)
			ADDPROPERTY(plA_Processes[lnX],"CSName"						, loProcess.CSName											)
			ADDPROPERTY(plA_Processes[lnX],"Description"				, loProcess.Description										)
			ADDPROPERTY(plA_Processes[lnX],"ExecutablePath"				, loProcess.ExecutablePath									)
			ADDPROPERTY(plA_Processes[lnX],"ExecutionState"				, loProcess.ExecutionState									)
			ADDPROPERTY(plA_Processes[lnX],"Handle"						, loProcess.Handle											)
			ADDPROPERTY(plA_Processes[lnX],"HandleCount"				, loProcess.HandleCount										)
			ADDPROPERTY(plA_Processes[lnX],"InstallDate"				, loProcess.InstallDate										)
			ADDPROPERTY(plA_Processes[lnX],"KernelModeTime"				, loProcess.KernelModeTime									)
			ADDPROPERTY(plA_Processes[lnX],"MaximumWorkingSetSize"		, loProcess.MaximumWorkingSetSize							)
			ADDPROPERTY(plA_Processes[lnX],"MinimumWorkingSetSize"		, loProcess.MinimumWorkingSetSize							)
			ADDPROPERTY(plA_Processes[lnX],"Name"						, loProcess.Name											)
			ADDPROPERTY(plA_Processes[lnX],"OSCreationClassName"		, loProcess.OSCreationClassName								)
			ADDPROPERTY(plA_Processes[lnX],"OSName"						, loProcess.OSName											)
			ADDPROPERTY(plA_Processes[lnX],"OtherOperationCount"		, loProcess.OtherOperationCount								)
			ADDPROPERTY(plA_Processes[lnX],"OtherTransferCount"			, loProcess.OtherTransferCount								)
			ADDPROPERTY(plA_Processes[lnX],"PageFaults"					, loProcess.PageFaults										)
			ADDPROPERTY(plA_Processes[lnX],"PageFileUsage"				, loProcess.PageFileUsage									)
			ADDPROPERTY(plA_Processes[lnX],"ParentProcessID"			, loProcess.ParentProcessID									)
			ADDPROPERTY(plA_Processes[lnX],"PeakPageFileUsage"			, loProcess.PeakPageFileUsage								)
			ADDPROPERTY(plA_Processes[lnX],"PeakVirtualSize"			, loProcess.PeakVirtualSize									)
			ADDPROPERTY(plA_Processes[lnX],"PeakWorkingSetSize"			, loProcess.PeakWorkingSetSize								)
			ADDPROPERTY(plA_Processes[lnX],"Priority"					, loProcess.Priority										)
			ADDPROPERTY(plA_Processes[lnX],"PrivatePageCount"			, loProcess.PrivatePageCount								)
			ADDPROPERTY(plA_Processes[lnX],"ProcessID"					, loProcess.ProcessId										)
			ADDPROPERTY(plA_Processes[lnX],"QuotaNonPagedPoolUsage"		, loProcess.QuotaNonPagedPoolUsage							)
			ADDPROPERTY(plA_Processes[lnX],"QuotaPagedPoolUsage"		, loProcess.QuotaPagedPoolUsage								)
			ADDPROPERTY(plA_Processes[lnX],"QuotaPeakNonPagedPoolUsage"	, loProcess.QuotaPeakNonPagedPoolUsage						)
			ADDPROPERTY(plA_Processes[lnX],"QuotaPeakPagedPoolUsage"	, loProcess.QuotaPeakPagedPoolUsage							)
			ADDPROPERTY(plA_Processes[lnX],"ReadOperationCount"			, loProcess.ReadOperationCount								)
			ADDPROPERTY(plA_Processes[lnX],"ReadTransferCount"			, loProcess.ReadTransferCount								)
			ADDPROPERTY(plA_Processes[lnX],"SessionId"					, loProcess.SessionId										)
			ADDPROPERTY(plA_Processes[lnX],"Status"						, loProcess.Status											)
			ADDPROPERTY(plA_Processes[lnX],"TerminationDate"			, loProcess.TerminationDate									)
			ADDPROPERTY(plA_Processes[lnX],"ThreadCount"				, loProcess.ThreadCount										)
			ADDPROPERTY(plA_Processes[lnX],"UserModeTime"				, loProcess.UserModeTime									)
			ADDPROPERTY(plA_Processes[lnX],"VirtualSize"				, loProcess.VirtualSize										)
			ADDPROPERTY(plA_Processes[lnX],"WindowsVersion"				, loProcess.WindowsVersion									)
			ADDPROPERTY(plA_Processes[lnX],"WorkingSetSize"				, loProcess.WorkingSetSize									)
			ADDPROPERTY(plA_Processes[lnX],"WriteOperationCount"		, loProcess.WriteOperationCount								)
			ADDPROPERTY(plA_Processes[lnX],"WriteTransferCount"			, loProcess.WriteTransferCount								)
			ADDPROPERTY(plA_Processes[lnX],"UserName"					, This.GetUserInfoFromSessionID(loProcess.SessionId,"User")	)
			ADDPROPERTY(plA_Processes[lnX],"WindowInfo[1]"				, .null.													)
			
			If Lower(plA_Processes[lnX].UserName) == Lower(GetEnv("username"))
				Local laWindow(1)
				This.GetWindowByProcessID(plA_Processes[lnX].ProcessID, @laWindow)
				
				Dimension plA_Processes[lnX].WindowInfo[Alen(laWindow)]
				
				Local lnY
				lnY = 1
				For lnY = 1 to Alen(laWindow)
					plA_Processes[lnX].WindowInfo[lnY] = laWindow[lnY]
				EndFor 
				
			EndIf
			
		    lnX = lnX + 1
		NEXT
		
		RETURN loProcesses.Count
	ENDFUNC
	
	
	*** <summary>
	*** Get Logged Users
	*** </summary>
	*** <param name="plA_AllUsers">Pointer to Array</param>
	*** <remarks></remarks>
	FUNCTION GetLoggedUsers
		Lparameters plA_AllUsers
		Local loLocator, loWMI, loProcesses, lnX
		
		loLocator 	= CREATEOBJECT('WBEMScripting.SWBEMLocator')
		loWMI		= loLocator.ConnectServer() 
		loWMI.Security_.ImpersonationLevel = 3  		&& Impersonate
		 
		 
		TEXT TO PRVCPWRSHL NOSHOW TEXTMERGE PRETEXT 2 
			SELECT * FROM Win32_Process WHERE Name = 'winlogon.exe'
		ENDTEXT 
		loProcesses	= loWMI.ExecQuery(PRVCPWRSHL)
		
		IF loProcesses.Count == 0
			Return .f.
		ENDIF
		
		Dimension plA_AllUsers[loProcesses.Count]
		lnX = 1
		
		FOR EACH loProcess in loProcesses	
			plA_AllUsers[lnX] = This.GetUserInfo(loProcess.sessionID)
			lnX = lnX + 1
		ENDFOR
	
		RETURN loProcesses.Count
	Endfunc
	
	*** <summary>
	*** Get user current Session ID
	*** </summary>
	*** <remarks></remarks>
	FUNCTION GetConsoleSessionID 
	    LOCAL nSessionID
	    nSessionID = This.WTSGetActiveConsoleSessionId()
	    RETURN nSessionID
	ENDFUNC
	
	*** <summary>
	*** Get All Sessions ID in the Server
	*** </summary>
	*** <param name="plA_AllSessions">Pointer to Array</param>
	*** <remarks></remarks>
	FUNCTION GetAllSessions 
		Lparameters plA_AllSessions
	   
	    LOCAL lcInfo, count, index, sessionId, lA_AllSessions
	    
	    * Call WTSEnumerateSessions to retrieve session information
	    Dimension plA_AllSessions[1]
	    Count = This.WTSEnumerateSessions(WTS_CURRENT_SERVER_HANDLE, 0, 1, @lA_AllSessions)
	    
	    If count == 0
	    	Return count
	    EndIf
	    
	    Dimension plA_AllSessions[count]
	    
	    index = 0
	    For index = 1 to count
			
			plA_AllSessions[index] = CreateObject("empty")
			AddProperty(plA_AllSessions[index],"Session", lA_AllSessions[index].SessionID)
			AddProperty(plA_AllSessions[index],"User"	, This.GetUserInfoFromSessionID(lA_AllSessions[index].SessionID,"User"))
			AddProperty(plA_AllSessions[index],"Station", lA_AllSessions[index].WinStationName)
			AddProperty(plA_AllSessions[index],"Domain"	, This.GetUserInfoFromSessionID(lA_AllSessions[index].SessionID,"Domain"))
			AddProperty(plA_AllSessions[index],"State"	, This.GetStateDescription(lA_AllSessions[index].State))
	    
	    EndFor
	    
	    Return count
	ENDFUNC
		
	*** <summary>
	*** Returns the Description of the User State
	*** </summary>
	*** <param name="pln_State"></param>
	*** <remarks></remarks>
	Hidden Procedure GetStateDescription
		Lparameters pln_State
		
		Local lc_State
		lc_State = 'Unknown'
		
		DO CASE
			CASE pln_State = WTS_Active
				lc_State = "Active"

			CASE pln_State = WTS_Connected
				lc_State = "Connected"
			
			CASE pln_State = WTS_ConnectQuery
				lc_State = "ConnectQuery"
			
			CASE pln_State = WTS_Disconnected
				lc_State = "Disconnected"	
			
			CASE pln_State = WTS_Idle
				lc_State = "Idle"
			
			CASE pln_State = WTS_Listen
				lc_State = "Listen"	
			
			CASE pln_State = WTS_Reset
				lc_State = "Reset"
			
			CASE pln_State = WTS_Down
				lc_State = "Down"
			
			CASE pln_State = WTS_Init
				lc_State = "Init"
							
		ENDCASE

		Return lc_State
	Endproc

	
	*** <summary>
	*** Returns the Info of the User associated with the session ID
	*** </summary>
	*** <param name="lnSessionID"></param>
	*** <remarks></remarks>
	Function GetUserInfo
		Lparameters lnSessionID
		
		Local loUserInfo
		loUserInfo = CreateObject("empty")
		AddProperty(loUserInfo,"Session", lnSessionID)
		AddProperty(loUserInfo,"User"	, This.GetUserInfoFromSessionID(lnSessionID,"User"))
		AddProperty(loUserInfo,"Station", This.GetUserInfoFromSessionID(lnSessionID,"Station"))
		AddProperty(loUserInfo,"Domain"	, This.GetUserInfoFromSessionID(lnSessionID,"Domain"))
		
		Return loUserInfo
	Endfunc
	
	*** <summary>
	*** Function to retrieve the data associated with a session ID
	*** </summary>
	*** <param name="lnSessionID">User Session ID</param>
	*** <remarks></remarks>
	FUNCTION GetUserInfoFromSessionID
		LPARAMETERS pln_SessionID, plc_Type
		
		LOCAL lcValue, lcBuffer, lnBytesReturned, ln_WTSType
	    ln_WTSType = -1
	    
	    Do case
	    	Case Lower(plc_Type) = "user"
	    		ln_WTSType = WTS_UserName
	    	
	    	Case Lower(plc_Type) = "domain"
	    		ln_WTSType = WTS_DomainName
	    		
	    	Case Lower(plc_Type) = "station"
	    		ln_WTSType = WTS_WinStationName	
	    	
	    *	Case Lower(plc_Type) = "state"
	    *		ln_WTSType = WTS_CONNECTSTATE_CLASS
	    	
	    	Otherwise 
	    		ln_WTSType = -1
	    EndCase
		
	*	Set Step On 
		* Get the Info associated with the session ID
		lcValue = This.WTSQuerySessionInformation(WTS_CURRENT_SERVER_HANDLE, pln_SessionID, ln_WTSType)
		
		RETURN lcValue
	ENDFUNC
	
	*** <summary>
	*** Parses the Buffer to a readable string
	*** </summary>
	*** <param name="plc_Buffer">Buffer to parse</param>
	*** <param name="pln_TotalBytes">Size, in bytes, of the Buffer</param>
	*** <remarks></remarks>
	Procedure ParseBuffer
		Lparameters plc_Buffer, pln_TotalBytes
	
		Local lcDataBuffer, ln_Buffer, ln_TotalBytes
		ln_Buffer = This.buf2Dword(plc_Buffer)
		ln_TotalBytes =  Iif(Type('pln_TotalBytes') == 'N', pln_TotalBytes, This.GlobalSize(ln_Buffer))
		
		lcDataBuffer = This.RtlMoveMemory(ln_Buffer, ln_TotalBytes) 
		
		This.WTSFreeMemory(ln_Buffer)
		Return lcDataBuffer
	Endproc
	
	*** <summary>
	*** Converts a four-byte buffer into a 32-bit unsigned integer
	*** </summary>
	*** <param name="lcBuffer">Four-byte buffer</param>
	*** <remarks></remarks>
	Procedure buf2dword (lcBuffer)
		RETURN Asc(SUBSTR(lcBuffer, 1,1)) + ;
		    Asc(SUBSTR(lcBuffer, 2,1)) * 256 +;
		    Asc(SUBSTR(lcBuffer, 3,1)) * 65536 +;
		    Asc(SUBSTR(lcBuffer, 4,1)) * 16777216

	Endproc
	
	*** <summary>
	*** Converts a four-byte buffer into a 32-bit unsigned integer
	*** </summary>
	*** <param name="lcBuffer">Four-byte buffer</param>
	*** <remarks></remarks>
	FUNCTION buf2dword_VFP9(lcBuffer)
		RETURN Asc(SUBSTR(lcBuffer, 1,1)) + ;
			BitLShift(Asc(SUBSTR(lcBuffer, 2,1)),  8) +;
			BitLShift(Asc(SUBSTR(lcBuffer, 3,1)), 16) +;
			BitLShift(Asc(SUBSTR(lcBuffer, 4,1)), 24)  
	endfunc
	
	
	*****************************************************************************************************************************
	* |  |                                     Instantiated from DLLs, OCXs and etc...                                     |  | *
	* V  V                                                                                                                 V  V *
	*****************************************************************************************************************************
	
	*** <summary>
	*** Returns Window Caption
	*** </summary>
	*** <param name="pln_Handle"Window Handle></param>
	*** <remarks></remarks>
	Hidden Procedure GetWindowText
		LPARAMETERS pln_Handle

		DECLARE INTEGER GetWindowText IN user32.dll as WinClass_GetWindowText INTEGER hWnd, STRING @lpString, INTEGER nMaxCount
		
		LOCAL lcTitle, nMaxLength, lnHandle
		nMaxLength = 255
		lnHandle = pln_Handle
		lcTitle = SPACE(nMaxLength)
		WinClass_GetWindowText(lnHandle, @lcTitle, nMaxLength)
		lcTitle = ALLTRIM(CHRtran(lcTitle,CHR(0),''))
		
		CLEAR dlls WinClass_GetWindowText
		
		RETURN lcTitle
	ENDFUNC

	*** <summary>
	*** Returns Window Process ID
	*** </summary>
	*** <param name="pln_Handle">Window Handle</param>
	*** <remarks></remarks>
	Hidden Procedure GetWindowThreadProcessId
		LPARAMETERS pln_Handle
		
		DECLARE INTEGER GetWindowThreadProcessId IN user32.dll AS WinClass_GetWindowThreadProcessId INTEGER, INTEGER @
		
		LOCAL lnHandle, nWindowProcessID
		STORE 0 TO lnHandle, nWindowProcessID
		
		lnHandle = pln_Handle
		
		WinClass_GetWindowThreadProcessId(lnHandle, @nWindowProcessID)
		
		CLEAR DLLS WinClass_GetWindowThreadProcessId
		
		RETURN nWindowProcessID
	ENDfunc

	*** <summary>
	*** Returns Styles of the Windows
	*** </summary>
	*** <param name="pln_Handle">Window Handle</param>
	*** <param name="pln_Style">Type of response</param>
	*** <remarks></remarks>
	Hidden Procedure GetWindowLong
		LPARAMETERS pln_Handle, pln_Style
		DECLARE INTEGER GetWindowLong IN user32.dll as WinClass_GetWindowLong INTEGER hWnd, INTEGER nIndex
		
		LOCAL lnHandle, lnStyle, lnLong
		lnHandle = pln_Handle
		lnStyle = pln_Style 
		
		lnLong = WinClass_GetWindowLong(lnHandle, lnStyle)
		
		CLEAR DLLS WinClass_GetWindowLong
		
		RETURN lnLong
	ENDFUNC

	*** <summary>
	*** Returns the Windows Handle from class name and window name
	*** </summary>
	*** <param name="plc_ClassName">Class Name</param>
	*** <param name="plc_WindowName">Window Name</param>
	*** <remarks></remarks>
	Hidden Procedure FindWindow
		LPARAMETERS plc_ClassName, plc_WindowName
		
		DECLARE INTEGER FindWindow	IN WIN32API AS WinClass_FindWindow	STRING lpClassName	, STRING lpWindowName
		
		LOCAL lnHandle, lcClassName, lcWindowName
		lcClassName 	= plc_ClassName
		lcWindowName	= plc_WindowName
	    lnHandle 		= WinClass_FindWindow(lcClassName, lcWindowName)
		
		CLEAR DLLS WinClass_FindWindow
		
		RETURN lnHandle
	ENDFUNC
		
	*** <summary>
	*** Returns Window Handle by another window handle and its relation with the current window
	*** </summary>
	*** <param name="pln_Handle">Window Handle</param>
	*** <param name="pln_Relation">Window Relation</param>
	*** <remarks></remarks>
	Hidden Procedure GetWindow
		LPARAMETERS pln_Handle, pln_Relation
		
		DECLARE INTEGER GetWindow	IN WIN32API as WinClass_GetWindow	INTEGER hWnd , INTEGER uCmd

		LOCAL lnHandle, lnRelation, lnNextWin_Handle
		lnHandle = pln_Handle
		lnRelation = pln_Relation
		
		lnNextWin_Handle = WinClass_GetWindow(lnHandle, lnRelation)

		CLEAR DLLS WinClass_GetWindow
		
		RETURN lnNextWin_Handle
	EndFunc
	
	*** <summary>
	*** Get User Current Session ID
	*** </summary>
	*** <remarks></remarks>
	Hidden Procedure WTSGetActiveConsoleSessionId
	    
	    DECLARE INTEGER WTSGetActiveConsoleSessionId IN win32api as WinClass_WTSGetActiveConsoleSessionId  LONG hToken
	    LOCAL nSessionID
	    nSessionID = WinClass_WTSGetActiveConsoleSessionId(2)
	    
	    CLEAR DLLS "WinClass_WTSGetActiveConsoleSessionId"
	    
	    RETURN nSessionID
	EndProc 
	
	*** <summary>
	*** Retrieve information about a specified Remote Desktop Services session
	*** </summary>
	*** <param name="pln_hServer"			> A handle to the Remote Desktop Session Host (RD Session Host) server.									</param>
	*** <param name="pln_SessionId"			> The identifier of the session for which you want to retrieve information								</param>
	*** <param name="pln_WTSInfoClass"		> The type of session information to retrieve (e.g., session ID, session username, session state, etc.)	</param>
	*** <remarks></remarks>
	Hidden Procedure WTSQuerySessionInformation
		Lparameters pln_hServer, pln_SessionId, pln_WTSInfoClass
		
		Local lc_Buffer, ln_pBytesReturned, lcValue
		lc_Buffer = This.Pointer 
		ln_pBytesReturned = 0
		
		DECLARE INTEGER WTSQuerySessionInformationA IN Wtsapi32 as WinClass_WTSQuerySessionInformation;
		    INTEGER hServer, ;
		    INTEGER SessionId, ;
		    INTEGER WTSInfoClass, ;
		    STRING @ppBuffer, ;
		    INTEGER @pBytesReturned
		
		WinClass_WTSQuerySessionInformation(pln_hServer, pln_SessionId, pln_WTSInfoClass, @lc_Buffer, @ln_pBytesReturned)
		
		lcValue = This.ParseBuffer(lc_Buffer, ln_pBytesReturned) 
		lcValue = Substr(lcValue, 1, At(This.Null, lcValue)-1)
		
		CLEAR DLLS "WinClass_WTSQuerySessionInformation"
		
		Return lcValue
	Endproc
	
	*** <summary>
	*** Returns the data from a pointer in Memory
	*** </summary>
	*** <param name="pln_Buffer"	> Pointer to the source memory block from which data will be copied	</param>
	*** <param name="pln_BufLength"	> Number of bytes to copy from the source to the destination		</param>
	*** <remarks></remarks>
	Hidden Procedure RtlMoveMemory
		Lparameters pln_Buffer, pln_BufLength
		
		Local lcDataBuf, ln_BufLength
		&& The least a character can have is 1 byte, só the length can be the numbers of bytes in the Buffer. 
		&& This way the String will never heave lenth inferior to the necessary
	*	ln_BufLength 	= This.GlobalSize(pln_Buffer) &&Iif(Type('pln_BufLength') == 'N', pln_BufLength, This.GlobalSize(pln_Buffer)) 
		ln_BufLength	= pln_BufLength
		lcDataBuf 		= Replicate(This.Null, ln_BufLength) 
	*	Set Step On 
		* Copies a block of memory from one location to another
		DECLARE RtlMoveMemory IN kernel32 As WinClass_RtlMoveMemory;
				STRING @, INTEGER, INTEGER 
		
		
		WinClass_RtlMoveMemory(@lcDataBuf, pln_Buffer, ln_BufLength) &&LEFT(lcBuffer, lnBytesReturned - 1)
		
		CLEAR DLLS "WinClass_RtlMoveMemory"
		
		Return lcDataBuf
	Endproc
	
	*** <summary>
	*** 
	*** </summary>
	*** <param name="pln_hServer"		> A handle to the RD Session Host server for which you want to enumerate sessions. You can typically pass WTS_CURRENT_SERVER_HANDLE to operate on the local server	</param>
	*** <param name="pln_Reserved"		> Reserved parameter. It should be set to 0																															</param>
	*** <param name="pln_Version"		> Reserved parameter. It should be set to 1																															</param>
	*** <param name="plA_AllSessions"	> A pointer to an array of WTS_SESSION_INFO																															</param>
	*** <remarks></remarks>
	Hidden Procedure WTSEnumerateSessions
		Lparameters pln_hServer, pln_Reserved, pln_Version, plA_AllSessions
		
		Local lc_ppSessionInfo, ln_pCount, lcInfo, index
		
		DECLARE INTEGER WTSEnumerateSessions IN Wtsapi32 as WinClass_WTSEnumerateSessions ;
		    INTEGER hServer, ;
		    INTEGER Reserved, ;
		    INTEGER Version, ;
		    STRING @ppSessionInfo, ;
		    INTEGER @pCount
		
		lc_ppSessionInfo = This.Pointer 
		ln_pCount = 0
		WinClass_WTSEnumerateSessions(pln_hServer, pln_Reserved, pln_Version, @lc_ppSessionInfo, @ln_pCount)
		
		lcInfo = This.ParseBuffer(lc_ppSessionInfo)
	    
	    Dimension plA_AllSessions[ln_pCount]
	    
	    index = 0
	    For index = 0 to ln_pCount-1
			
			SessionId 		= This.buf2dword(SUBSTR(lcInfo, index*PWTS_SESSION_INFO+1, 4))
			
			Local lnAddress
			lnAddress		= This.buf2dword(SUBSTR(lcInfo, index*PWTS_SESSION_INFO+5, 4))
			WinStationName	= SUBSTR(lcInfo, lnAddress - This.buf2dword(lc_ppSessionInfo)+1)
			WinStationName	= SUBSTR(WinStationName, 1, AT(CHR(0),WinStationName)-1)
			
			State 			= This.buf2dword(SUBSTR(lcInfo, index*PWTS_SESSION_INFO+9, 4))
			
			plA_AllSessions[index+1] = CreateObject("empty")
			AddProperty(plA_AllSessions[index+1],"SessionId"		, SessionId)
			AddProperty(plA_AllSessions[index+1],"WinStationName"	, WinStationName)
			AddProperty(plA_AllSessions[index+1],"State"			, State)
	    
	    EndFor
		
		CLEAR DLLS "WinClass_WTSEnumerateSessions"
		
		Return ln_pCount
	Endproc
	
	*** <summary>
	*** Function retrieves the current size, in bytes, of the specified global memory object
	*** </summary>
	*** <param name="plm_MemObj"></param>
	*** <remarks></remarks>
	Hidden Procedure GlobalSize
		Lparameters plm_MemObj
		
		Local lnObjSize
		*The GlobalSize function retrieves the current size, in bytes, of the specified global memory object
		DECLARE INTEGER GlobalSize IN kernel32 As WinClass_GlobalSize INTEGER hMem
		
		lnObjSize = WinClass_GlobalSize(plm_MemObj)
		
		CLEAR DLLS "WinClass_GlobalSize"
		
		Return lnObjSize
	Endproc

	*** <summary>
	*** Free Memory
	*** </summary>
	*** <param name="pln_Pointer">Pointer</param>
	*** <remarks></remarks>
	Hidden Procedure WTSFreeMemory
		Lparameters pln_Pointer
		
		DECLARE WTSFreeMemory IN Wtsapi32 As WinClass_WTSFreeMemory INTEGER pMemory
		
		WinClass_WTSFreeMemory(pln_Pointer)
	
		CLEAR DLLS "WinClass_WTSFreeMemory"
	Endproc
	
ENDDEFINE