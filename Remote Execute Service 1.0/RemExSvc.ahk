;@Ahk2Exe-SetName RemExSvc
;@Ahk2Exe-SetOrigFilename RemExSvc.exe
;@Ahk2Exe-SetDescription Remote Execute Service
;@Ahk2Exe-SetVersion 1.0.0.0
;@Ahk2Exe-SetCompanyName Provolve B.V.
;@Ahk2Exe-SetCopyright Ferry van Gelderen
;@Ahk2Exe-SetMainIcon main.ico
;@Ahk2Exe-AddResource Computer.ico, 160
;@Ahk2Exe-AddResource EXE.ico, 206
;@Ahk2Exe-AddResource StartIn.ico, 207
;@Ahk2Exe-AddResource Status.ico, 208
;@Ahk2Exe-AddResource PS.ico, 209

#Requires AutoHotkey v2.0
#SingleInstance off
#NoTrayIcon
Persistent

ProductName := "RemExSvc"
ProductVersion := "1.0"
ProductDescription := "Remote Execute Service"
if A_PtrSize = 4
	OS_Bit := "32-bit"
else
	OS_Bit := "64-bit"
SplitPath A_ScriptFullPath, &FileName, &Dir,, &NoExt_FileName

sessionId := 0
DllCall("ProcessIdToSessionId", "UInt", DllCall("GetCurrentProcessId"), "UInt*", &sessionId)
if sessionId = 0
	{
	ArgumentsNumber := A_Args.Length
	if ArgumentsNumber = 1 ; Start HOST service process. (first step)
		MyService.Start()
	if ArgumentsNumber = 2 ; Start HOST program process. (second step)
		{
		UsePsScriptManager := false
		ServerName := A_Args[1]
		LogFile := Dir "\" NoExt_FileName "_program.log"
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" FileName " started with Process Id: " ProcessExist() "`r`n", LogFile
		ReturnMsg := WinService.Stop("", ProductName)
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStop Windows Service " ProductName ". ReturnCode: " ReturnMsg "`r`n", LogFile
		ReturnMsg := WinService.Delete("", ProductName)
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDelete Windows Service " ProductName ". ReturnCode: " ReturnMsg "`r`n", LogFile
		ReturnCode := ""
		GetActiveUsersArray := GetActiveUsernames()
		for ActiveUser in GetActiveUsersArray
			{
			SessionId_UserName_array := StrSplit(ActiveUser, "`t")
			FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFound the following active user on this machine: " SessionId_UserName_array[2] "`r`n", LogFile
			ReturnCode .= SessionId_UserName_array[2] "`t"
			}
		if ReturnCode = ""
			FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNo active users found on this machine.`r`n", LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tWaiting for a task to be received from: " ServerName "`r`n", LogFile
		loop
			{
			Task := NamedPipeRpc.Send(ReturnCode, ProductName . OS_Bit)
			if InStr(Task, "Executable:`t",, 1)
				{
				ReturnCode := ""
				temp_array := StrSplit(Task, "`t")
				Target := TransForm(temp_array[2])
				WorkingDir := TransForm(temp_array[3])
				Interactive := Trim(temp_array[4])
				System := Trim(temp_array[5])
				User := Trim(temp_array[6])
				if !User
					System := 1
				if Interactive = 0
					User := "SYSTEM"
				temp_array := ""
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart executable: " Target "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing working direcory: " WorkingDir "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSystem: " System "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tInteractive: " Interactive "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart for account: " User "`r`n", LogFile
				if Interactive
					{
					SessionId := GetActiveUsernames(User)
					if SessionId
						{
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSession Id: " SessionId "`r`n", LogFile
						if System
							{
							Stop_PID := 0
							Stop_hProc := 0
							rc := SystemProcess.CreateProcessAsSystemInteractive(Target, "", WorkingDir, "", 0, SessionId , &Stop_PID, &Stop_hProc)
							if (rc = -1 OR rc = -2 OR rc = -3 OR rc = -4 OR rc = -5 OR rc = -6 OR rc = -7 OR rc = -8 OR rc = -9 OR rc = -10 OR rc = -11)
								{
								ErrorCode := A_LastError
								if rc = -1
									{
									ErrorMessage := "Command line is empty."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -2
									{
									ErrorMessage := "Executable not found."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -3
									{
									ErrorMessage := "Running process not found."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -4
									{
									ErrorMessage := "Session Id not found for username " User ". Please make sure the username is logged in."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -5
									{
									ErrorMessage := "GetWinlogonPid failed."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -6
									{
									ErrorMessage := "OpenProcess failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -7
									{
									ErrorMessage := "OpenProcessToken failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -8
									{
									ErrorMessage := "DuplicateTokenEx failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -9
									{
									ErrorMessage := "CreateEnvironmentBlock function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -10
									{
									ErrorMessage := "Failed attempt to launch program or document. Error: " ErrorCode 
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -11
									{
									ErrorMessage := "GetExitCodeProcess function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								Reply := NamedPipeRpc.Send("Failed", ProductName . OS_Bit . "_Stop_PID")
								}
							else
								{
								Stop_ExitCode := -1
								StopProcessRunning := true
								SetTimer MonitorProcess, 500
								while StopProcessRunning = true
									sleep 250
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tProgram stopped with return code: " Stop_ExitCode "`r`n", LogFile
								ReturnCode := "Info:`t" . Stop_ExitCode
								}
							}
						else
							{
							Stop_PID := 0
							Stop_hProc := 0
							rc := SystemProcess.CreateProcessAsUserInteractive(Target, "", WorkingDir, "", 0, SessionId, &Stop_PID, &Stop_hProc)
							if (rc = -1 OR rc = -2 OR rc = -3 OR rc = -4 OR rc = -6 OR rc = -7 OR rc = -8 OR rc = -9 OR rc = -10 OR rc = -11)
								{
								ErrorCode := A_LastError
								if rc = -1
									{
									ErrorMessage := "Command line is empty."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -2
									{
									ErrorMessage := "Executable not found."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -3
									{
									ErrorMessage := "Running process not found."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -4
									{
									ErrorMessage := "Session Id not found for username " User ". Please make sure the username is logged in."
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -6
									{
									ErrorMessage := "OpenProcess function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -7
									{
									ErrorMessage := "OpenProcessToken function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -8
									{
									ErrorMessage := "DuplicateTokenEx function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -9
									{
									ErrorMessage := "CreateEnvironmentBlock function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -10
									{
									ErrorMessage := "Failed attempt to launch program or document. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								if rc = -11
									{
									ErrorMessage := "GetExitCodeProcess function failed. Error: " ErrorCode
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
									ReturnCode := "Error:`t" . ErrorMessage
									}
								Reply := NamedPipeRpc.Send("Failed", ProductName . OS_Bit . "_Stop_PID")
								}
							else
								{
								Stop_ExitCode := -1
								StopProcessRunning := true
								SetTimer MonitorProcess, 500
								while StopProcessRunning = true
									sleep 250
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tProgram stopped with return code: " Stop_ExitCode "`r`n", LogFile
								ReturnCode := "Info:`t" . Stop_ExitCode
								}
							}
						}
					else
						{
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSession Id not found for username: " User "`r`n", LogFile
						ReturnCode := "Session Id not found for username " User ". Please make sure the username is logged in."
						ReturnCode := "Error:`t" . ReturnCode
						}
					}
				else
					{
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSession Id: 0`r`n", LogFile
					Stop_PID := 0
					Stop_hProc := 0
					rc := SystemProcess.CreateProcessAsSystem(Target, "", WorkingDir, 0, &Stop_PID, &Stop_hProc)
					if (rc = -1 OR rc = -2 OR rc = -10 OR rc = -11)
						{
						ErrorCode := A_LastError
						if rc = -1
							{
							ErrorMessage := "Command line is empty."
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
							ReturnCode := "Error:`t" . ErrorMessage
							}
						if rc = -2
							{
							ErrorMessage := "Executable not found."
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
							ReturnCode := "Error:`t" . ErrorMessage
							}
						if rc = -10
							{
							ErrorMessage := "Failed attempt to launch program or document. Error: " ErrorCode 
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
							ReturnCode := "Error:`t" . ErrorMessage
							}
						if rc = -11
							{
							ErrorMessage := "GetExitCodeProcess function failed. Error: " ErrorCode
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ErrorMessage "`r`n" , LogFile
							ReturnCode := "Error:`t" . ErrorMessage
							}
						Reply := NamedPipeRpc.Send("Failed", ProductName . OS_Bit . "_Stop_PID")
						}
					else
						{
						Stop_ExitCode := -1
						StopProcessRunning := true
						SetTimer MonitorProcess, 500
						while StopProcessRunning = true
							sleep 250
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tProgram stopped with return code: " Stop_ExitCode "`r`n", LogFile
						ReturnCode := "Info:`t" . Stop_ExitCode						
						}
					}
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart Executable task finished. Waiting for a task to be received from: " ServerName "`r`n", LogFile
				Sleep 250
				}
			if InStr(Task, "PowerShell:`t",, 1)
				{
				ReturnCode := ""
				temp_array := StrSplit(Task, "`t")
				PSCmd := Trim(temp_array[2])
				WorkingDir := TransForm(temp_array[3])
				temp_array := ""
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart PowerShell command: " PSCmd "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing working direcory: " WorkingDir "`r`n", LogFile
				if UsePsScriptManager = false
					{
					Result := PsScriptManager.Register()
					if !Result = 0
						{
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tPowerShell COM registration error. psScript.dll failed to register on this system. Error: " Result  "`r`n", LogFile
						ReturnCode := "PowerShell COM registration error."
						}
					else
						{
						UsePsScriptManager := true
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tPowerShell COM successfully registered.`r`n", LogFile
						}
					}
				if UsePsScriptManager = true
					{
					Command := "Set-Location -Path `"" WorkingDir "`"`r`n" PSCmd
					try
						psReturn := PsScriptManager.ComInstance.PS_Script(Command)
					catch as err
						{
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tPowerShell command stopped with error:`r`n" err.Message "`r`n" , LogFile
						ReturnCode := "Error:`t" . err.Message
						}
					else
						{
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tPowerShell command successfully stopped with return message:`r`n" psReturn "`r`n" , LogFile
						ReturnCode := "Info:`t" . psReturn
						}
					psReturn := ""
					}
				Result := ""
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart PowerShell command task finished. Waiting for a task to be received from: " ServerName "`r`n", LogFile
				}
			If Task = "Quit"
				{
				if UsePsScriptManager = true
					{
					Result := PsScriptManager.Unregister()
					if !Result = 0
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tPowerShell COM unregistration error. psScript.dll failed to unregister on this system. Error: " Result  "`r`n", LogFile
					else
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tPowerShell COM successfully unregistered.`r`n", LogFile
					}
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tReceived message: " Task "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tQuitting...`r`n", LogFile
				FileAppend "---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n", LogFile
				ExitApp 0
				}
			if (Task = "CreateNamedPipe failed" OR Task = "ConnectNamedPipe failed" OR Task = "Write failed" OR Task = "Timeout waiting for reply" OR Task = "ReadLine failed" OR Task = "Timeout waiting for client")
				{
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tNamed pipe communication failed. Error: " Task "`r`n", LogFile
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tQuitting...`r`n", LogFile
				FileAppend "---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n", LogFile
				ExitApp 1
				}
			}
		}
	ExitApp 0
	}
else ; Start CLIENT process. (GUI)
	{
	if CheckForOtherSessions(0) > 0
		{
		if WinExist(ProductDescription " " ProductVersion " (" OS_Bit ")")
			WinActivate
		ExitApp 0
		}
	; --------------------------------------------------------------- Apply dark theme to the built-in MsgBox and InputBox if enabled ---------------------------------------------------------------
	AppsUseLightTheme := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
	if AppsUseLightTheme = 0
		{
		#DllLoad gdi32.dll
		POINT(x := 0, y := 0)
			{
			NumPut("int", x, "int", y, buf := Buffer(8))
			buf.DefineProp("x", {Get: NumGet.Bind(, 0, "int"), Set: IntPut.Bind(0)})
			buf.DefineProp("y", {Get: NumGet.Bind(, 4, "int"), Set: IntPut.Bind(4)})
			return buf
			}
		RECT2(left := 0, top := 0, right := 0, bottom := 0)
			{
			static ofst := Map("left", 0, "top", 4, "right", 8, "bottom", 12)
			NumPut("int", left, "int", top, "int", right, "int", bottom, buf := Buffer(16))
			for k, v in ofst
				buf.DefineProp(k, {Get: NumGet.Bind(, v, "int"), Set: IntPut.Bind(v)})
			return buf
			}
		IntPut(ofst, _, v) => NumPut("int", v, _, ofst)
		class __MsgBox
			{
			static __New()
				{
				static nativeMsgbox   := MsgBox.Call.Bind(MsgBox)
				static nativeInputBox := InputBox.Call.Bind(InputBox)
				MsgBox.DefineProp("Call", {Call: BoxEx})
				InputBox.DefineProp("Call", {Call: BoxEx})
				BoxEx(_this, params*)
					{
					static WM_COMMNOTifY := 0x44
					static WM_INITDIALOG := 0x0110
					iconNumber := 1
					iconFile   := ""
					if (params.length = (_this.MaxParams + 2))
						iconNumber := params.Pop()
					if (params.length = (_this.MaxParams + 1))
						iconFile := params.Pop()
					SetThreadDpiAwarenessContext(-4)
					if (_this.Name = "MsgBox")
						OnMessage(WM_COMMNOTifY, ON_WM_COMMNOTifY, -1)
					else
						OnMessage(WM_INITDIALOG, ON_WM_INITDIALOG, -1)
					return native%_this.Name%(params*)
					ON_WM_INITDIALOG(wParam, lParam, msg, hwnd)
						{
						OnMessage(WM_INITDIALOG, ON_WM_INITDIALOG, 0)
						WNDENUMPROC(hwnd)
						}
					ON_WM_COMMNOTifY(wParam, lParam, msg, hwnd)
						{
						DetectHiddenWindows(true)
						if (msg = 68 && wParam = 1027)
							OnMessage(0x44, ON_WM_COMMNOTifY, 0),
							EnumThreadWindows(GetCurrentThreadId(), CallbackCreate(WNDENUMPROC), 0)
						}
					WNDENUMPROC(hwnd, *)
						{
						static SM_CICON         := "W" SysGet(11) " H" SysGet(12)
						static SM_CSMICON       := "W" SysGet(49) " H" SysGet(50)
						static ICON_BIG         := 1
						static ICON_SMALL       := 0
						static WM_SETICON       := 0x80
						static WS_CLIPCHILDREN  := 0x02000000
						static WS_CLIPSIBLINGS  := 0x04000000
						static WS_EX_COMPOSITED := 0x02000000
						static WS_VSCROLL       := 0x00200000
						static winAttrMap       := Map(2, 2, 4, 0, 10, true, 17, true, 20, true, 38, 2, 35, 0x2b2b2b)
						Critical()
						SetWinDelay(-1)
						SetControlDelay(-1)
						DetectHiddenWindows(true)
						if !WinExist("ahk_class #32770 ahk_id" hwnd)
							return 1
						WinSetStyle("+" (WS_CLIPCHILDREN | WS_CLIPSIBLINGS))
						WinSetExStyle("+" (WS_EX_COMPOSITED))
						SetWindowTheme(hwnd, "DarkMode_Explorer")
						if iconFile
							{
							hICON_SMALL := LoadPicture(iconFile, SM_CSMICON " Icon" iconNumber, &handleType)
							hICON_BIG   := LoadPicture(iconFile, SM_CICON " Icon" iconNumber, &handleType)
							PostMessage(WM_SETICON, ICON_SMALL, hICON_SMALL)
							PostMessage(WM_SETICON, ICON_BIG, hICON_BIG)
							}
						for dwAttribute, pvAttribute in winAttrMap
							DwmSetWindowAttribute(hwnd, dwAttribute, pvAttribute)
						GWL_WNDPROC(hwnd, hICON_SMALL?, hICON_BIG?)
						return 0
						}
					GWL_WNDPROC(winId := "", hIcons*)
						{
						static SetWindowLong     := DllCall.Bind(A_PtrSize = 8 ? "SetWindowLongPtr" : "SetWindowLong", "ptr",, "int",, "ptr",, "ptr")
						static BS_FLAT           := 0x8000
						static BS_BITMAP         := 0x0080
						static DPI               := (A_ScreenDPI / 96)
						static WM_CLOSE          := 0x0010
						static WM_CTLCOLORBTN    := 0x0135
						static WM_CTLCOLORDLG    := 0x0136
						static WM_CTLCOLOREDIT   := 0x0133
						static WM_CTLCOLORSTATIC := 0x0138
						static WM_DESTROY        := 0x0002
						static WM_SETREDRAW      := 0x000B
						DetectHiddenWindows(true)
						SetControlDelay(-1)
						btns    := []
						btnHwnd := hbrush1 := hbrush2 := ""
						for ctrl in WinGetControlsHwnd(winId)
							{
							classNN := ControlGetClassNN(ctrl)
							SetWindowTheme(ctrl, !InStr(classNN, "Edit") ? "DarkMode_Explorer" : "DarkMode_CFD")
							if !InStr(classNN, "B")
								continue
							ControlSetStyle("+" (BS_FLAT | BS_BITMAP), ctrl)
							btns.Push(btnHwnd := ctrl)
							}
						WindowProcOld := SetWindowLong(winId, -4, CallbackCreate(WNDPROC))
						WNDPROC(hwnd, uMsg, wParam, lParam)
							{
							Critical(-1)
							DetectHiddenWindows(true)
							SetWinDelay(-1)
							SetControlDelay(-1)
							if !hbrush1
								hbrush1 := CreateSolidBrush(0x202020)
							if !hbrush2
								hbrush2 := CreateSolidBrush(0x2b2b2b)
							switch uMsg
								{
								case WM_CTLCOLORSTATIC:
									{
									SelectObject(wParam, hbrush2)
									SetBkMode(wParam, 0)
									SetTextColor(wParam, 0xFFFFFF)
									SetBkColor(wParam, 0x2b2b2b)
									for _hwnd in btns
										PostMessage(WM_SETREDRAW,,,_hwnd)
									GetWindowRect(winId, rcW := RECT2())
									GetClientRect(winId, rcC := RECT2())
									GetWindowRect(btnHwnd, rcBtn := RECT2())
									pt   := POINT()
									pt.y := rcW.bottom - rcBtn.bottom
									ScreenToClient(winId, pt)
									hdc        := GetWindowDC(winId)
									rcC.top    := rcBtn.top + pt.y -2
									rcC.bottom *= 2
									rcC.right  *= 2
									SetBkMode(hdc, 0)
									FillRect(hdc, rcC, hbrush1)
									ReleaseDC(winId, hdc)
									for _hwnd in btns
										PostMessage(WM_SETREDRAW, 1,,_hwnd)
									return hbrush2
									}
								case WM_CTLCOLORDLG, WM_CTLCOLOREDIT:
									{
									SelectObject(wParam, hbrush2)
									SetBkMode(wParam, 0)
									SetTextColor(wParam, 0xFFFFFF)
									SetBkColor(wParam, 0x2b2b2b)
									return hbrush2
									}
								case WM_CTLCOLORBTN:
									{
									SelectObject(wParam, hbrush1)
									SetBkMode(wParam, 0)
									SetTextColor(wParam, 0xFFFFFF)
									SetBkColor(wParam, 0x202020)
									return hbrush2
									}
								case WM_DESTROY:
									{
									for v in [hbrush1, hbrush2]
										(v && DeleteObject(v))
									for v in hIcons
										(v??0) && DestroyIcon(v)
									}
								}
							return CallWindowProc(WindowProcOld, hwnd, uMsg, wParam, lParam)
							}
						}
					}
				CallWindowProc(lpPrevWndFunc, hWnd, uMsg, wParam, lParam) => DllCall("CallWindowProc", "Ptr", lpPrevWndFunc, "Ptr", hwnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam)
				ClientToScreen(hWnd, lpPoint) => DllCall("User32.dll\ClientToScreen", "ptr", hWnd, "ptr", lpPoint, "int")
				CreateSolidBrush(crColor) => DllCall('Gdi32.dll\CreateSolidBrush', 'uint', crColor, 'ptr')
				DestroyIcon(hIcon) => DllCall("DestroyIcon", "ptr", hIcon)
				DWMSetWindowAttribute(hwnd, dwAttribute, pvAttribute, cbAttribute := 4) => DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr" , hwnd, "UInt", dwAttribute, "Ptr*", &pvAttribute, "UInt", cbAttribute)
				DeleteObject(hObject) => DllCall('Gdi32.dll\DeleteObject', 'ptr', hObject, 'int')
				EnumThreadWindows(dwThreadId, lpfn, lParam) => DllCall("User32.dll\EnumThreadWindows", "uint", dwThreadId, "ptr", lpfn, "uptr", lParam, "int")
				FillRect(hDC, lprc, hbr) => DllCall("User32.dll\FillRect", "ptr", hDC, "ptr", lprc, "ptr", hbr, "int")
				GetClientRect(hWnd, lpRect) => DllCall("User32.dll\GetClientRect", "ptr", hWnd, "ptr", lpRect, "int")
				GetCurrentThreadId() => DllCall("kernel32.dll\GetCurrentThreadId", "uint")
				GetWindowDC(hwnd) => DllCall("User32.dll\GetWindowDC", "ptr", hwnd, "ptr")
				GetWindowRect(hWnd, lpRect) => DllCall("User32.dll\GetWindowRect", "ptr", hWnd, "ptr", lpRect, "uptr")
				GetWindowRgn(hWnd, hRgn) => DllCall("User32.dll\GetWindowRgn", "ptr", hWnd, "ptr", hRgn, "int")
				GetWindowRgnBox(hWnd, hRgn) => DllCall("User32.dll\GetWindowRgnBox", "ptr", hWnd, "ptr", hRgn, "int")
				ReleaseDC(hWnd, hDC) => DllCall("User32.dll\ReleaseDC", "ptr", hWnd, "ptr", hDC, "int")
				ScreenToClient(hWnd, lpPoint) => DllCall("User32.dll\ScreenToClient", "ptr", hWnd, "ptr", lpPoint, "int")
				SelectObject(hdc, hgdiobj) => DllCall('Gdi32.dll\SelectObject', 'ptr', hdc, 'ptr', hgdiobj, 'ptr')
				SetBkColor(hdc, crColor) => DllCall('Gdi32.dll\SetBkColor', 'ptr', hdc, 'uint', crColor, 'uint')
				SetBkMode(hdc, iBkMode) => DllCall('Gdi32.dll\SetBkMode', 'ptr', hdc, 'int', iBkMode, 'int')
				SetTextColor(hdc, crColor) => DllCall('Gdi32.dll\SetTextColor', 'ptr', hdc, 'uint', crColor, 'uint')
				SetThreadDpiAwarenessContext(dpiContext) => DllCall("SetThreadDpiAwarenessContext", "ptr", dpiContext, "ptr")
				SetWindowTheme(hwnd, pszSubAppName, pszSubIdList := "") => (!DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "ptr", StrPtr(pszSubAppName), "ptr", pszSubIdList ? StrPtr(pszSubIdList) : 0) ? true : false)
				}
			}
		}
	else
		AppsUseLightTheme := 1
	OSVersion := StrSplit(A_OSVersion, ".")
	if Not OSVersion[1] >= 10
		{
		MsgBox "The operating system version is not supported for this product. Microsoft Windows 10 or higher is required.", ProductName " " ProductVersion, "16 T30"
		ExitApp 1150
		}
	if A_IsAdmin = 0
		{
		MsgBox "Administrative privileges are required to run this program.", ProductDescription " " ProductVersion, "16 T30"
		ExitApp 5
		}
	if AppsUseLightTheme = 0
		{
		global DarkColors := Map("Background", 0x303030, "Controls", 0x202020, "Font", 0xA0A0A0)
		global TextBackgroundBrush := DllCall("Gdi32.dll\CreateSolidBrush", "UInt", DarkColors["Background"], "Ptr")
		BackColor := "303030"
		TextColor := "cADADAD"
		EditColor := "cF0F0F0"
		BackgroundColor := "Background202020"
		BackgroundColorDisabled := "Background282828"
		ProgressColor := "c0078D7"
		}
	else
		{
		BackColor := "E0E0E0"
		TextColor := "c404040"
		EditColor := "c000000"
		BackgroundColor := "BackgroundF0F0F0"
		ProgressColor := "c0078D7"
		}
	ErrorColor := "cE74856"
	PreventExit := false
	RunningProcess := false
	Connected := false
	MainGui := Gui("+DpiScale", ProductDescription " " ProductVersion " (" OS_Bit ")")
	MainGui.BackColor := BackColor
	if AppsUseLightTheme = 0
		{
		DllCall("dwmapi\DwmSetWindowAttribute", "ptr", MainGui.hwnd, "int", 20, "int*", 1, "int", 4)
		uxtheme := DllCall("GetModuleHandle", "str", "uxtheme", "ptr")
		SetPreferredAppMode := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 135, "ptr")
		FlushMenuThemes := DllCall("GetProcAddress", "ptr", uxtheme, "ptr", 136, "ptr")
		DllCall(SetPreferredAppMode, "int", 1)
		DllCall(FlushMenuThemes)
		}
	if A_IsCompiled
		{
		ComputerIcon := 'HBITMAP:*' LoadPicture(A_ScriptFullPath, 'Icon2')
		ExeTargetIcon := 'HBITMAP:*' LoadPicture(A_ScriptFullPath, 'Icon3')
		StartInIcon := 'HBITMAP:*' LoadPicture(A_ScriptFullPath, 'Icon4')
		StatusIcon := 'HBITMAP:*' LoadPicture(A_ScriptFullPath, 'Icon5')
		PSTargetIcon := 'HBITMAP:*' LoadPicture(A_ScriptFullPath, 'Icon6')
		}
	else
		{
		ComputerIcon := 'HBITMAP:*' LoadPicture(A_ScriptDir "\Computer.ico")
		ExeTargetIcon := 'HBITMAP:*' LoadPicture(A_ScriptDir "\EXE.ico")
		StartInIcon := 'HBITMAP:*' LoadPicture(A_ScriptDir "\StartIn.ico")
		StatusIcon := 'HBITMAP:*' LoadPicture(A_ScriptDir "\Status.ico")
		PSTargetIcon := 'HBITMAP:*' LoadPicture(A_ScriptDir "\PS.ico")
		}
	ComputerName := MainGui.Add("Text", "w100 Right", "Computer name:")
	ComputerName.Opt(TextColor)
	IconComputer := MainGui.Add("Picture", "x+10 w16 h16 BackgroundTrans")
	IconComputer.Value := ComputerIcon
	Target := MainGui.Add("Text", "xm y+12 w100 Right", "Target:")
	Target.Opt(TextColor)
	IconTarget := MainGui.Add("Picture", "x+10 w16 h16 BackgroundTrans")
	IconTarget.Value := ExeTargetIcon
	StartInText := MainGui.Add("Text", "xm y+12 w100 Right", "Start in:")
	StartInText.Opt(TextColor)
	IconStartIn := MainGui.Add("Picture", "x+10 w16 h16 BackgroundTrans")
	IconStartIn.Value := StartInIcon
	StatusText := MainGui.Add("Text", "xm y+12 w100 Right", "Status:")
	StatusText.Opt(TextColor)
	IconStatus := MainGui.Add("Picture", "x+10 w16 h16 BackgroundTrans")
	IconStatus.Value := StatusIcon
	ComputerName := MainGui.Add("Edit", "ym w500 -E0x200", A_ComputerName)
	ComputerName.Opt(EditColor)
	ComputerName.OnEvent("Change", Set_Default_Button_Connect_Disconnect)
	if AppsUseLightTheme = 0
		{
		ComputerName.Opt(BackgroundColor)
		DllCall("uxtheme\SetWindowTheme", "ptr", ComputerName.Hwnd, "str", "DarkMode_Explorer", "ptr", 0)
		}
	Command := MainGui.Add("Edit", "w500 Disabled -E0x200")
	Command.Opt(EditColor)
	Command.OnEvent("Change", Set_Default_Button_Start_Stop)
	if AppsUseLightTheme = 0
		{
		Command.Opt(BackgroundColorDisabled)
		DllCall("uxtheme\SetWindowTheme", "ptr", Command.Hwnd, "str", "DarkMode_Explorer", "ptr", 0)
		}
	WorkingDir := MainGui.Add("Edit", "w500 Disabled -E0x200")
	WorkingDir.Opt(EditColor)
	if AppsUseLightTheme = 0
		{
		WorkingDir.Opt(BackgroundColorDisabled)
		DllCall("uxtheme\SetWindowTheme", "ptr", WorkingDir.Hwnd, "str", "DarkMode_Explorer", "ptr", 0)
		}
	Status := MainGui.Add("Edit", "w500 r20 ReadOnly -Wrap HScroll -E0x200", "Ready for connection.")
	Status.Opt(TextColor)
	if AppsUseLightTheme = 0
		{
		Status.Opt(BackgroundColorDisabled)
		DllCall("uxtheme\SetWindowTheme", "ptr", Status.Hwnd, "str", "DarkMode_Explorer", "ptr", 0)
		}
	Progress := MainGui.Add("Progress", "w500 h5")
	Progress.Opt(ProgressColor)
	Progress.Opt(BackgroundColor)
	ButtonConnect := MainGui.Add("Button", "ym w120 Default", "Connect")
	ButtonConnect.OnEvent("Click", Connect)
	if AppsUseLightTheme = 0
		{
		ButtonConnect.Opt(BackgroundColor)
		DllCall("uxtheme\SetWindowTheme", "Ptr", ButtonConnect.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	ButtonStart := MainGui.Add("Button", "w120 Disabled", "Start")
	ButtonStart.OnEvent("Click", Start)
	if AppsUseLightTheme = 0
		{
		ButtonStart.Opt(BackgroundColor)
		DllCall("uxtheme\SetWindowTheme", "Ptr", ButtonStart.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	ButtonHelp := MainGui.Add("Button", "w120", "Help")
	ButtonHelp.OnEvent("Click", Help)
	if AppsUseLightTheme = 0
		{
		ButtonHelp.Opt(BackgroundColor)
		DllCall("uxtheme\SetWindowTheme", "Ptr", ButtonHelp.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	TimeoutText := MainGui.Add("Text", "w120 y+175 Right", "Timeout (seconds):")
	TimeoutText.Opt(TextColor)
	RadioStartPS := MainGui.Add("Radio", "w120 y+14 Disabled Right", "")
	if AppsUseLightTheme = 0
		{	
		DllCall("uxtheme\SetWindowTheme", "Ptr", RadioStartPS.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	RadioStartPS.OnEvent("Click", SetTarget)
	RadioStartExe := MainGui.Add("Radio", "w120 y+14 Checked Disabled Right", "")
	if AppsUseLightTheme = 0
		{
		DllCall("uxtheme\SetWindowTheme", "Ptr", RadioStartExe.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	RadioStartExe.OnEvent("Click", SetTarget)
	Interactive := MainGui.Add("Checkbox", "w120 y+14 Disabled", "Interactive")
	Interactive.OnEvent("Click", SetInteractive)
	if AppsUseLightTheme = 0
		{
		DllCall("uxtheme\SetWindowTheme", "Ptr", Interactive.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	System := MainGui.Add("Checkbox", "w120 y+14 Disabled Checked", "System")
	if AppsUseLightTheme = 0
		{
		DllCall("uxtheme\SetWindowTheme", "Ptr", System.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	ButtonDisconnect := MainGui.Add("Button", "ym w120 Disabled", "Disconnect")
	ButtonDisconnect.OnEvent("Click", Disconnect)
	if AppsUseLightTheme = 0
		{
		ButtonDisconnect.Opt(BackgroundColor)
		DllCall("uxtheme\SetWindowTheme", "Ptr", ButtonDisconnect.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	ButtonStop := MainGui.Add("Button", "w120 Disabled", "Stop")
	ButtonStop.OnEvent("Click", Stop)
	if AppsUseLightTheme = 0
		{
		ButtonStop.Opt(BackgroundColor)
		DllCall("uxtheme\SetWindowTheme", "Ptr", ButtonStop.hWnd, "Str", "DarkMode_Explorer", "Ptr", 0)
		}
	MainGui.Add("Button", "w120 Hidden", "")
	Timeout := MainGui.Add("Edit", "w120 y+175 Disabled Number -E0x200", "3600")
	Timeout.Opt(EditColor)
	if AppsUseLightTheme = 0
		{
		Timeout.Opt(BackgroundColorDisabled)
		DllCall("uxtheme\SetWindowTheme", "ptr", Timeout.Hwnd, "str", "DarkMode_Explorer", "ptr", 0)
		}	
	RadioStartPSText := MainGui.Add("Text", "w120  y+6", "PowerShell")
	RadioStartPSText.Opt(TextColor)	
	RadioStartExeText := MainGui.Add("Text", "w120", "Executable")
	RadioStartExeText.Opt(TextColor)
	StartForText := MainGui.Add("Text", "w120", "Start for account:")
	StartForText.Opt(TextColor)
	SelectUser := MainGui.Add("DropDownList", "w120 y+10 vUser -E0x200 Disabled")
	SelectUser.Opt(EditColor)
	if AppsUseLightTheme = 0
		{
		SelectUser.Opt(BackgroundColorDisabled)
		DllCall("uxtheme\SetWindowTheme", "ptr", SelectUser.Hwnd, "str", "DarkMode_CFD", "ptr", 0)
		SetWindowTheme(MainGui, AppsUseLightTheme = 0)
		}
	Status.SetFont("", "Consolas")
	MainGui.OnEvent("Close", CloseApp)
	MainGui.OnEvent("DropFiles", Gui_DropFiles)
	MainGui.Show()
	SetStopProcess := false
	return

	SetWindowTheme(GuiObj, DarkMode := True)
		{
		static GWL_WNDPROC := -4
		static Init := False
		global IsDarkMode := DarkMode
		Mode_CFD := (DarkMode ? "DarkMode_CFD" : "CFD")
		for hWnd, Ctrl in GuiObj
			{
			if (Ctrl.Type = "DDL" || Ctrl.Type = "ComboBox")
				DllCall("uxtheme\SetWindowTheme", "Ptr", Ctrl.hWnd, "Str", Mode_CFD, "Ptr", 0)
			}
		if (!Init)
			{
			global WindowProcNew := CallbackCreate(WindowProc)
			SetWindowLong := (A_PtrSize = 8) ? "SetWindowLongPtr" : "SetWindowLong"
			global WindowProcOld := DllCall(SetWindowLong, "Ptr", GuiObj.Hwnd, "Int", GWL_WNDPROC, "Ptr", WindowProcNew, "Ptr")
			Init := True
			}
		}
	WindowProc(hwnd, uMsg, wParam, lParam)
		{
		static WM_CTLCOLORLISTBOX := 0x0134
		static WM_CTLCOLOREDIT := 0x0133
		static DC_BRUSH := 18
		if (IsDarkMode)
			{
			switch uMsg
				{
				case WM_CTLCOLOREDIT, WM_CTLCOLORLISTBOX:
					{
					DllCall("Gdi32.dll\SetTextColor", "Ptr", wParam, "UInt", DarkColors["Font"])
					DllCall("Gdi32.dll\SetBkColor", "Ptr", wParam, "UInt", DarkColors["Controls"])
					DllCall("Gdi32.dll\SetDCBrushColor", "Ptr", wParam, "UInt", DarkColors["Controls"], "UInt")
					return DllCall("Gdi32.dll\GetStockObject", "Int", DC_BRUSH, "Ptr")
					}
				}
			}
		return DllCall("CallWindowProc", "Ptr", WindowProcOld, "Ptr", hwnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam)
		}

	Gui_DropFiles(GuiObj, GuiCtrlObj, FileArray, X, Y)
		{
		SplitPath FileArray[1], &GetCommand, &GetWorkingDir
		GetWorkingDir := StrReplace(GetWorkingDir, "\\" ComputerName.Value "\c$", "C:")
		GetWorkingDir := StrReplace(GetWorkingDir, "\\" ComputerName.Value "\admin$", "C:\Windows")
		Command.Value := GetCommand
		WorkingDir.Value := GetWorkingDir
		}

	Set_Default_Button_Connect_Disconnect(*)
		{
		if Connected = false
			{
			ButtonConnect.Opt("+Default")
			ButtonDisconnect.Opt("-Default")
			}
		else
			{
			ButtonConnect.Opt("-Default")
			ButtonDisconnect.Opt("+Default")
			}
		}
	
	Set_Default_Button_Start_Stop(*)
		{
		if RunningProcess = false
			{
			ButtonStart.Opt("+Default")
			ButtonStop.Opt("-Default")
			}
		else
			{
			ButtonStart.Opt("-Default")
			ButtonStop.Opt("+Default")
			}	
		}
	
	SetTarget(*)
		{
		Global DisableInteractive
		if RadioStartExe.Value = 1
			{
			Target.Value := "Target:"
			IconTarget.Value := ExeTargetIcon
			if DisableInteractive = false
				{
				System.Enabled := true
				Interactive.Enabled := true
				}
			Timeout.Enabled := false
			Timeout.Opt(TextColor)
			}
		else
			{
			Target.Value := "PowerShell:"
			IconTarget.Value := PSTargetIcon
			System.Enabled := false
			Interactive.Enabled := false
			SelectUser.Enabled := false
			Timeout.Enabled := true
			Timeout.Opt(EditColor)
			if AppsUseLightTheme = 0
				SelectUser.Opt(BackgroundColorDisabled)
			Status.Value := "Warning:`n`nStarting a new PowerShell or CMD process from the existing powershell`nsession will transfer the powershell controls to the new powershell or`nCMD session and will result in lossing connection with the host.`n`nUse the `"Executable`" option for starting a PowerShell.exe or CMD.exe`nprocess or use the `"PowerShell`" option to execute a PowerShell command`nonly."
			Status.Opt(EditColor)
			}
		}

	SetInteractive(*)
		{
		if Interactive.Value = 1
			{
			System.Enabled := true
			SelectUser.Enabled := true
			if AppsUseLightTheme = 0
				SelectUser.Opt(BackgroundColor)			
			}
		else
			{
			System.Enabled := false
			SelectUser.Enabled := false
			if AppsUseLightTheme = 0
				SelectUser.Opt(BackgroundColorDisabled)
			if System.Value != 1
				System.Value := 1
			}
		}

	Connect(*)
		{
		global PreventExit
		global Connected
		global EditColor
		global ErrorColor
		global DisableInteractive
		MainGui.Opt("+OwnDialogs")
		if ComputerName.Value
			{
			MainGui.Opt("Disabled")
			if (FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName) OR FileExist(A_WinDir . "\Temp\" . FileName))
				{
				if A_ComputerName = ComputerName.Value
					tempFilePath := A_WinDir . "\Temp\" . FileName
				else
					tempFilePath := "\\" . ComputerName.Value . "\admin$\Temp\" . FileName
				try
					FileDelete tempFilePath
				catch as err
					{
					Result := MsgBox("A connection is already established (from another computer). Would you like to continue connecting?",  ProductDescription " " ProductVersion, "48 YesNo")
					if Result = "Yes"
						{
						Status.Value := "Stop the current connection on " ComputerName.Value "..."
						Status.Opt(EditColor)
						sleep 50
						Progress.Opt("+0x8")
						SetTimer RunWaitProgress, 50
						ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Get Status", 5)
						if ReturnStatus = "Running"
							{
							Status.Value := "Stop the running process on " ComputerName.Value "..."
							Status.Opt(EditColor)
							sleep 50
							ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Stop", 5)
							ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Get Status", 5)
							}
						if ReturnStatus = "Stopped"
							ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Quit", 5)
						Status.Value := "Stop the remote execute service process on " ComputerName.Value "..."
						Status.Opt(EditColor)
						sleep 50
						ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "Quit", 15)
						Sleep 1000
						try 
							FileDelete tempFilePath
						catch as err
							{
							Status.Value := "Connection failed.`nManually stop the process " FileName " on the remote machine if needed."
							Status.Opt(ErrorColor)
							SetTimer RunWaitProgress, 0
							Progress.Opt("-0x8")
							Progress.Value := 0
							SoundPlay "*16"
							MainGui.Opt("-Disabled")
							return
							}
						Status.Value := "Connecting to " ComputerName.Value "..."
						Status.Opt(EditColor)
						SetTimer RunWaitProgress, 0
						Progress.Opt("-0x8")
						Progress.Value := 0
						sleep 50
						}
					else
						{
						MainGui.Opt("-Disabled")
						return
						}
					}
				tempFilePath := ""
				}
			else
				{
				MainGui.Opt("Disabled")
				Status.Value := "Connecting to " ComputerName.Value "..."
				Status.Opt(EditColor)
				sleep 50
				}
			if (DirExist("\\" . ComputerName.Value . "\admin$\Temp") OR DirExist(A_WinDir . "\Temp"))
				{
				PreventExit := true
				Progress.Value := 25
				ButtonConnect.Enabled := false
				ComputerName.Enabled := false
				if AppsUseLightTheme = 0
					ComputerName.Opt(BackgroundColorDisabled)
				SelectUser.Delete()
				Status.Value := "Copy file: " FileName "`nPlease Wait..."
				Status.Opt(EditColor)
				sleep 50
				BinaryPath := A_WinDir . "\Temp\" . FileName
				if A_ComputerName = ComputerName.Value
					{
					try 
						FileCopy A_ScriptFullPath, A_WinDir . "\Temp", 1
					}
				else
					{
					try 
						FileCopy A_ScriptFullPath, "\\" . ComputerName.Value . "\admin$\Temp", 1
					}
				Progress.Value := 50
				Status.Value := "Create and start Windows Service:`n" ProductName "`nPlease Wait..."
				Status.Opt(EditColor)
				sleep 50
				ReturnMsg := WinService.Add(ComputerName.Value, ProductName, BinaryPath . " " . A_ComputerName, "Demand", ProductName, "", "")
				if (ReturnMsg = 1 OR ReturnMsg = 1056)
					{
					Progress.Value := 75
					ActiveUsers := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "", 15)
					if ActiveUsers = ""
						{
						DisableInteractive := true
						Interactive.Enabled := false
						Interactive.Value := 0
						System.Enabled := false
						SelectUser.Enabled := false
						if AppsUseLightTheme = 0
							SelectUser.Opt(BackgroundColorDisabled)
						}
					else
						{
						DisableInteractive := false
						Interactive.Enabled := true
						System.Enabled := true
						ChooseUser := 1
						Loop parse, ActiveUsers, "`t"
							{
							if A_LoopField
								{
								SelectUser.Add([ A_LoopField ])
								if A_LoopField = A_UserName
									ChooseUser := A_Index
								}
							}
						SelectUser.Choose(ChooseUser)
						}
					ButtonDisconnect.Enabled := true
					ButtonStart.Enabled := true
					Command.Enabled := true
					if AppsUseLightTheme = 0
						Command.Opt(BackgroundColor)
					WorkingDir.Enabled := true
					if AppsUseLightTheme = 0
						WorkingDir.Opt(BackgroundColor)
					RadioStartExe.Enabled := true
					RadioStartPS.Enabled := true
					Connected := true
					if DisableInteractive = true
						Status.Value := "Connected to " ComputerName.Value ".`nNo active users found on this machine."
					else
						Status.Value := "Connected to " ComputerName.Value "."
					Status.Opt(EditColor)
					Message := ""
					Progress.Value := 100
					}
				else
					{
					Message := OSError(ReturnMsg).Message
					Progress.Value := 25
					Status.Value := "Delete Windows Service:`n" ProductName "..."
					Status.Opt(EditColor)
					ReturnMsg := WinService.Delete(ComputerName.Value, ProductName)
					if FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName)
						{
						try 
							FileDelete "\\" . ComputerName.Value . "\admin$\Temp\" . FileName
						}
					if FileExist(A_WinDir . "\Temp\" . FileName)
						{
						try 
							FileDelete A_WinDir . "\Temp\" . FileName
						}
					Progress.Value := 0
					ButtonConnect.Enabled := true
					ButtonStart.Enabled := false
					ComputerName.Enabled := true
					if AppsUseLightTheme = 0
						ComputerName.Opt(BackgroundColor)
					Command.Enabled := false
					if AppsUseLightTheme = 0
						Command.Opt(BackgroundColorDisabled)
					WorkingDir.Enabled := false
					if AppsUseLightTheme = 0
						WorkingDir.Opt(BackgroundColorDisabled)
					RadioStartExe.Enabled := false
					RadioStartPS.Enabled := false
					Interactive.Enabled := false
					System.Enabled := false
					SelectUser.Enabled := false
					if AppsUseLightTheme = 0
						SelectUser.Opt(BackgroundColorDisabled)
					Connected := false
					Status.Value := "Create and start Windows Service failed with Error:`n" Message
					Status.Opt(ErrorColor)
					SoundPlay "*16"
					Message := ""
					}
				PreventExit := false
				}
			else
				{
				Status.Value := "Computer " ComputerName.Value . " not found.`nMake sure this computer is active and the `"File and Printer Sharing`" firewall exceptions are enabled on this and the remote computer (Port 139 and 445 TCP)."
				Status.Opt(ErrorColor)
				SoundPlay "*48"
				}
			MainGui.Opt("-Disabled")
			}	
		return
		}

	Disconnect(*)
		{
		global PreventExit
		global Connected
		global EditColor
		if ComputerName.Value
			{
			MainGui.Opt("Disabled")
			Status.Value := "Disconnecting from " ComputerName.Value
			Status.Opt(EditColor)
			if (FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName) OR FileExist(A_WinDir . "\Temp\" . FileName))
				{
				PreventExit := true
				ButtonDisconnect.Enabled := false
				Progress.Value := 75
				ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "Quit", 15)
				Sleep 1000
				Progress.Value := 50
				Status.Value := "Delete file \\" . ComputerName.Value . "\admin$\Temp\" . FileName
				if FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName)
					{
					try 
						FileDelete "\\" . ComputerName.Value . "\admin$\Temp\" . FileName
					}
				if FileExist(A_WinDir . "\Temp\" . FileName)
					{
					try 
						FileDelete A_WinDir . "\Temp\" . FileName
					}
				Progress.Value := 0
				ButtonConnect.Enabled := true
				ButtonStart.Enabled := false
				ComputerName.Enabled := true
				if AppsUseLightTheme = 0
					ComputerName.Opt(BackgroundColor)
				Command.Enabled := false
				if AppsUseLightTheme = 0
					Command.Opt(BackgroundColorDisabled)
				WorkingDir.Enabled := false
				if AppsUseLightTheme = 0
					WorkingDir.Opt(BackgroundColorDisabled)
				RadioStartExe.Enabled := false
				RadioStartPS.Enabled := false
				Interactive.Enabled := false
				System.Enabled := false
				SelectUser.Enabled := false
				if AppsUseLightTheme = 0
					SelectUser.Opt(BackgroundColorDisabled)
				Timeout.Enabled := false
				Connected := false
				PreventExit := false
				if RadioStartExe.Value != 1
					{
					RadioStartExe.Value := 1
					Target.Value := "Target:"
					IconTarget.Value := ExeTargetIcon
					if System.Value != 1
						System.Value := 1
					if Interactive.Value = 1
						Interactive.Value := 0
					}
				Status.Value := "Ready for connection."
				Status.Opt(TextColor)
				}
			MainGui.Opt("-Disabled")
			}
		return
		}

	Start(*)
		{
		global RunningProcess
		global EditColor
		global ErrorColor
		global SetStopProcess
		MainGui.Opt("+OwnDialogs")
		if Command.Value
			{
			ButtonStart.Enabled := false
			ButtonDisconnect.Enabled := false
			Command.Enabled := false
			if AppsUseLightTheme = 0
				Command.Opt(BackgroundColorDisabled)
			WorkingDir.Enabled := false
			if AppsUseLightTheme = 0
				WorkingDir.Opt(BackgroundColorDisabled)
			RadioStartExe.Enabled := false
			RadioStartPS.Enabled := false
			CommunicationLost := false
			if RadioStartExe.Value = 1
				{
				Interactive.Enabled := false
				System.Enabled := false
				SelectUser.Enabled := false
				if AppsUseLightTheme = 0
					SelectUser.Opt(BackgroundColorDisabled)
				ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "Executable:`t" Command.Value "`t" WorkingDir.Value "`t" Interactive.Value "`t" System.Value "`t" SelectUser.Text, 15)
				Status.Value := "Starting: " Command.Value "`nPlease Wait..."
				Status.Opt(EditColor)
				RunningProcess := false
				SetStopProcess := false
				loop 
					{
					ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Get Status", 60)
					if ReturnStatus = "Failed"
						break					
					if ReturnStatus = "Running"
						{
						if RunningProcess = false
							{
							RunningProcess := true
							Status.Value := "Running: " Command.Value
							Status.Opt(EditColor)
							Progress.Opt("+0x8")
							SetTimer RunWaitProgress, 50
							ButtonStop.Enabled := true
							}
						}
					if ReturnStatus = "Stopped"
						{
						ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Quit", 5)
						break
						}
					if ReturnStatus = "NoPID"
						{
						ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Timeout waiting for client", 15)
						CommunicationLost := true
						Status.Value := "Failed to properly start the process: " Command.Value
						Status.Opt(ErrorColor)
						break
						}
					if (ReturnStatus = "Timeout waiting for host" OR ReturnStatus = "Open pipe failed" OR ReturnStatus = "Timeout reading message" OR ReturnStatus = "ReadLine failed" OR ReturnStatus = "Write reply failed")
						{
						CommunicationLost := true
						SetTimer RunWaitProgress, 0
						Progress.Opt("-0x8")
						break
						}
					if SetStopProcess = true
						{
						ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit . "_Stop_PID", "Stop", 15)
						Status.Value := "Stopping: " Command.Value "..."
						Status.Opt(EditColor)
						SetStopProcess := false
						}
					sleep 500
					}
				}
			else
				{
				Timeout.Enabled := false
				ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "PowerShell:`t" Command.Value "`t" WorkingDir.Value, 15)
				if (ReturnStatus = "Timeout waiting for host" OR ReturnStatus = "Open pipe failed" OR ReturnStatus = "Timeout reading message" OR ReturnStatus = "ReadLine failed" OR ReturnStatus = "Write reply failed")
					CommunicationLost := true
				else
					{
					Status.Value := "PS " WorkingDir.Value ">`n" Command.Value
					Status.Opt(EditColor)
					Progress.Opt("+0x8")
					SetTimer RunWaitProgress, 50
					}
				}
			if CommunicationLost = false
				{
				ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "", Timeout.Value)
				SetTimer RunWaitProgress, 0
				Progress.Opt("-0x8")
				Progress.Value := 100
				ButtonStop.Enabled := false
				ButtonStart.Enabled := true
				ButtonDisconnect.Enabled := true
				Command.Enabled := true
				if AppsUseLightTheme = 0
					Command.Opt(BackgroundColor)
				WorkingDir.Enabled := true
				if AppsUseLightTheme = 0
					WorkingDir.Opt(BackgroundColor)
				RadioStartExe.Enabled := true
				RadioStartPS.Enabled := true
				if (ReturnStatus = "Timeout waiting for host" OR ReturnStatus = "Open pipe failed" OR ReturnStatus = "Timeout reading message" OR ReturnStatus = "ReadLine failed" OR ReturnStatus = "Write reply failed")
					{
					if RadioStartExe.Value = 1
						{
						if DisableInteractive = false
							{
							Interactive.Enabled := true
							System.Enabled := true
							SelectUser.Enabled := true
							if AppsUseLightTheme = 0
								SelectUser.Opt(BackgroundColor)
							}
						RunningProcess := false
						}
					else
						Timeout.Enabled := true
					Status.Value := "Timeout waiting for host process. Error: " ReturnStatus "`nIncrease the `"timeout`" value if needed."
					Status.Opt(ErrorColor)
					SoundPlay "*16"
					}
				else
					{
					if InStr(ReturnStatus, "Error:`t",, 1)
						{
						ReturnStatus := StrReplace(ReturnStatus, "Error:`t")
						Status.Opt(ErrorColor)
						}
					else
						{
						ReturnStatus := StrReplace(ReturnStatus, "Info:`t")
						Status.Opt(EditColor)
						}				
					if RadioStartExe.Value = 1
						{
						if DisableInteractive = false
							{
							Interactive.Enabled := true
							System.Enabled := true
							SelectUser.Enabled := true
							if AppsUseLightTheme = 0
								SelectUser.Opt(BackgroundColor)
							}
						RunningProcess := false
						Status.Value := "Program stopped with return message:`n" OSError(ReturnStatus).Message
						SoundPlay "*64"
						}
					else
						{
						Timeout.Enabled := true
						Status.Value := "PowerShell stopped with return message:`n" ReturnStatus
						}
					}
				}
			else
				{
				Progress.Value := 0
				MainGui.Opt("Disabled")
				Status.Opt(EditColor)
				ButtonConnect.Enabled := false
				ButtonDisconnect.Enabled := false
				ButtonStart.Enabled := false
				ButtonStop.Enabled := false
				ComputerName.Enabled := false
				if AppsUseLightTheme = 0
					ComputerName.Opt(BackgroundColorDisabled)
				Command.Enabled := false
				if AppsUseLightTheme = 0
					Command.Opt(BackgroundColorDisabled)
				WorkingDir.Enabled := false
				if AppsUseLightTheme = 0
					WorkingDir.Opt(BackgroundColorDisabled)
				RadioStartExe.Enabled := false
				RadioStartPS.Enabled := false
				Interactive.Enabled := false
				System.Enabled := false
				SelectUser.Enabled := false
				if AppsUseLightTheme = 0
					SelectUser.Opt(BackgroundColorDisabled)
				if FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName)
					{
					try 
						FileDelete "\\" . ComputerName.Value . "\admin$\Temp\" . FileName
					}
				if FileExist("\\" . ComputerName.Value . "\admin$\Temp\psScript.dll")
					{
					try 
						FileDelete "\\" . ComputerName.Value . "\admin$\Temp\psScript.dll"
					}
				if FileExist(A_WinDir . "\Temp\" . FileName)
					{
					try 
						FileDelete A_WinDir . "\Temp\" . FileName
					}
				if FileExist(A_WinDir . "\Temp\psScript.dll")
					{
					try 
						FileDelete A_WinDir . "\Temp\psScript.dll"
					}
				MsgBox "Communication is lost with error: " ReturnStatus "`nThis program will exit.", ProductDescription " " ProductVersion, "16 T30"
				ExitApp 1
				}
			}
		return
		}

	Stop(*)
		{
		global SetStopProcess
		if SetStopProcess = false
			SetStopProcess := true
		return
		}

	Help(*)
		{
		MainGui.Opt("+OwnDialogs")
		MsgBox ProductDescription " (" ProductName ") is a light-weight application that lets you execute programs or PowerShell commands on remote Microsoft Windows systems without having to manually install client software. These programs can be started for the local system account (with or without user interaction) or the logged on user with interaction. The PowerShell commands will run under the local system account on the remote system without user interaction. The exit code (return code) or PowerShell command result will be send back when the remote process is finished or stopped, or the PowerShell command is finished.`n`nThe `"File and Printer Sharing`" firewall exceptions needs to be enabled on this and the remote computer (Port 139 and 445 TCP). You also need to have the appropriate administrative rights to establish a connection to the admin share (admin$) and to the Service Control Manager (SCM) on the remote system. When creating a connection, this program can temporarily be tagged as non-responsive by the operating system if the connection takes up some time to complete. The speed for setting up a connection can be improved when adding the appropriate firewall rules.`n`nDrag-and-drop executable (network) files on the GUI and the `"Target`" and `"Start in`" values are automatically set. This software is provided `"as is`". Use of the software is free and at your own risk.`n`n2025-2026 Ferry van Gelderen", ProductDescription " " ProductVersion, "64"
		return
		}

	CloseApp(*)
		{
		global PreventExit
		global RunningProcess
		global Connected
		global EditColor
		global ErrorColor
		MainGui.Opt("+OwnDialogs")
		if RunningProcess = true
			{
			MsgBox "A process is still running. Please stop this process first.", ProductDescription " " ProductVersion, "48 T30"
			return true
			}
		else
			{
			if PreventExit = true
				return true
			else
				{
				MainGui.Opt("Disabled")
				Status.Opt(EditColor)
				ButtonConnect.Enabled := false
				ButtonDisconnect.Enabled := false
				ButtonStart.Enabled := false
				ButtonStop.Enabled := false
				ComputerName.Enabled := false
				if AppsUseLightTheme = 0
					ComputerName.Opt(BackgroundColorDisabled)
				Command.Enabled := false
				if AppsUseLightTheme = 0
					Command.Opt(BackgroundColorDisabled)
				WorkingDir.Enabled := false
				if AppsUseLightTheme = 0
					WorkingDir.Opt(BackgroundColorDisabled)
				RadioStartExe.Enabled := false
				RadioStartPS.Enabled := false
				Interactive.Enabled := false
				System.Enabled := false
				SelectUser.Enabled := false
				if AppsUseLightTheme = 0
					SelectUser.Opt(BackgroundColorDisabled)
				if Connected = true
					{
					Status.Value := "Disconnecting from " ComputerName.Value
					Status.Opt(EditColor)
					if (FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName) OR FileExist(A_WinDir . "\Temp\" . FileName))
						{
						ButtonDisconnect.Enabled := false
						Progress.Value := 75
						ReturnStatus := NamedPipeRpc.Receive(ComputerName.Value, ProductName . OS_Bit, "Quit", 15)
						Sleep 1000
						Progress.Value := 50
						Status.Value := "Delete file \\" . ComputerName.Value . "\admin$\Temp\" . FileName
						Status.Opt(EditColor)
						if FileExist("\\" . ComputerName.Value . "\admin$\Temp\" . FileName)
							{
							try 
								FileDelete "\\" . ComputerName.Value . "\admin$\Temp\" . FileName
							}
						if FileExist("\\" . ComputerName.Value . "\admin$\Temp\psScript.dll")
							{
							try 
								FileDelete "\\" . ComputerName.Value . "\admin$\Temp\psScript.dll"
							}
						if FileExist(A_WinDir . "\Temp\" . FileName)
							{
							try 
								FileDelete A_WinDir . "\Temp\" . FileName
							}
						if FileExist(A_WinDir . "\Temp\psScript.dll")
							{
							try 
								FileDelete A_WinDir . "\Temp\psScript.dll"
							}
						Progress.Value := 0
						}
					}
				ExitApp 0
				}
			}
		}
	}

; --------------------------------------------------------------- Function to determine if other sessions are running with the option to stop these sessions [ CLIENT ] ---------------------------------------------------------------
CheckForOtherSessions(Stop := 0)
	{
	FoundSessions := 0
	s := 4096
	h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", ProcessExist(), "Ptr")
	DllCall("Advapi32.dll\OpenProcessToken", "Ptr", h, "UInt", 32, "PtrP", &t := 0)
	DllCall("Advapi32.dll\LookupPrivilegeValue", "Ptr", 0, "Str", "SeDebugPrivilege", "Int64P", &luid := 0)
	ti := Buffer(16, 0)
	NumPut( "UInt", 1
		, "Int64", luid
		, "UInt", 2
		, ti)
	r := DllCall("Advapi32.dll\AdjustTokenPrivileges", "Ptr", t, "Int", false, "Ptr", ti, "UInt", 0, "Ptr", 0, "Ptr", 0)
	DllCall("CloseHandle", "Ptr", t)
	DllCall("CloseHandle", "Ptr", h)
	hModule := DllCall("LoadLibrary", "Str", "Psapi.dll")
	a := Buffer(s)
	DllCall("Psapi.dll\EnumProcesses", "Ptr", a, "UInt", s, "UIntP", &r)
	Loop r // 4
		{
		id := NumGet(a, A_Index * 4, "UInt")
		h := DllCall("OpenProcess", "UInt", 0x0010 | 0x0400, "Int", false, "UInt", id, "Ptr")
		if !h
			continue
		n := Buffer(s, 0)
		e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Ptr", n, "UInt", s//2)
		if !e
			e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Ptr", n, "UInt", s//2)
		SplitPath StrGet(n), &n
		DllCall("CloseHandle", "Ptr", h)
		if (n && e)
			{
			if n = A_ScriptName
				{
				if id != ProcessExist()
					{
					FoundSessions++
					if Stop = 1
						ProcessClose id
					}
				}
			}
		}
	DllCall("FreeLibrary", "Ptr", hModule)
	return FoundSessions
	}

; --------------------------------------------------------------- GUI timer for Run and wait notification progress bar [ CLIENT ] ---------------------------------------------------------------
RunWaitProgress(ProgressBar := 0)
	{
	Progress.Value := ProgressBar++
	if ProgressBar = 100
		ProgressBar := 0
	Sleep 50
	}

; --------------------------------------------------------------- Monitor running process timer function [ HOST ] ---------------------------------------------------------------
MonitorProcess()
	{
	global Stop_ExitCode
	global Stop_PID
	global Stop_hProc
	global StopProcessRunning
	global ProductName
	global OS_Bit
	global LogFile
	static MonitorProcessMessage := ""
	static TimeoutCounter := 0
	static StartedLogged := false
	StopProcessRunning := true
	if (Stop_PID = 0)
		{
		TimeoutCounter++
		if (TimeoutCounter > 100)
			{
			FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t[MonitorProcess] Timeout for Process ID (" Stop_PID "). Program failed to return the Process ID.`r`n", LogFile
			MonitorProcessMessage := "NoPID"
			}
		}
	exitCode := 0
	if (Stop_hProc)
		{
		if DllCall("kernel32.dll\GetExitCodeProcess", "Ptr", Stop_hProc, "UInt*", &exitCode := 0)
			{
			if (exitCode = 259)
				{
				if (!StartedLogged)
					{
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[MonitorProcess] Program is started successfully with Process ID: " Stop_PID "`r`n", LogFile
					StartedLogged := true
					}
				MonitorProcessMessage := "Running"
				}
			else
				{
				Stop_ExitCode := exitCode
				if (StartedLogged)
					{
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[MonitorProcess] Program is stopped with exit code: " Stop_ExitCode "`r`n", LogFile
					StartedLogged := false
					}
				MonitorProcessMessage := "Stopped"
				}
			}
		}
	Reply := NamedPipeRpc.Send(MonitorProcessMessage, ProductName . OS_Bit . "_Stop_PID")
	if (Reply = "Stop")
		{
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[MonitorProcess] Stop request received. Stop program with Process ID: " Stop_PID "`r`n", LogFile
		if (Stop_hProc)
			DllCall("kernel32.dll\TerminateProcess", "Ptr", Stop_hProc, "UInt", 1)
		}
	else if (Reply = "Quit")
		{
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[MonitorProcess] Quit request received. Exiting monitor process function.`r`n", LogFile
		if (Stop_hProc)
			{
			DllCall("kernel32.dll\CloseHandle", "Ptr", Stop_hProc)
			Stop_hProc := 0
			}
		SetTimer MonitorProcess, 0
		StopProcessRunning := false
		return
		}
	if (Reply = "CreateNamedPipe failed" OR Reply = "ConnectNamedPipe failed" OR Reply = "Write failed" OR Reply = "Timeout waiting for reply" OR Reply = "ReadLine failed" OR Reply = "Timeout waiting for client")
		{
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t[MonitorProcess] Named pipe communication failed. Error: " Reply "`r`n", LogFile
		if (Stop_hProc)
			{
			DllCall("kernel32.dll\CloseHandle", "Ptr", Stop_hProc)
			Stop_hProc := 0
			}
		SetTimer MonitorProcess, 0
		StopProcessRunning := false
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[MonitorProcess] Quitting...`r`n", LogFile
		FileAppend "---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n", LogFile
		ExitApp 1		
		}
	}

; --------------------------------------------------------------- Function for translating existing environment variables to their actual values [ HOST ] ---------------------------------------------------------------
TransForm(Str) 
	{
	spo := 1
	out := ""
	Str := Trim(Str)
	while (fpo := RegexMatch(Str, "(%(.*?)%)|``(.)", &m, spo))
		{
		out .= SubStr(Str, spo, fpo-spo)
		spo := fpo + StrLen(m[0])
		if (m[1])
			out .= (env := EnvGet(m[2])) ? env : m[1]
		else
 			out .= m[3]
		}
	return out SubStr(Str, spo)
	}

; --------------------------------------------------------------- Function to get all active usernames with their session id on this computer [ HOST ] ---------------------------------------------------------------
GetActiveUsernames(ActiveUserNameorSessionId := "")
	{
	WTS_CURRENT_SERVER_HANDLE := 0
	pInfo := 0
	count := 0
	if ActiveUserNameorSessionId = ""
		UserName_Sessions := []
	else
		UserName_Sessions := ""
	if !DllCall("Wtsapi32\WTSEnumerateSessions", "Ptr", WTS_CURRENT_SERVER_HANDLE, "UInt", 0, "UInt", 1, "Ptr*", &pInfo, "UInt*", &count)
		return UserName_Sessions
	structSize := (A_PtrSize = 8) ? 24 : 12
	stateOffset := (A_PtrSize = 8) ? 16 : 8
	Loop count
		{
		pItem := pInfo + (A_Index - 1) * structSize
		sessionId := NumGet(pItem, 0, "UInt")
		state := NumGet(pItem, stateOffset, "Int")
		if (state = 0 && sessionId > 0)
			{
			pBuffer := 0
			bytesReturned := 0
			WTSUserName := 5
			if DllCall("Wtsapi32\WTSQuerySessionInformationW", "Ptr", WTS_CURRENT_SERVER_HANDLE, "UInt", sessionId, "Int", WTSUserName, "Ptr*", &pBuffer, "UInt*", &bytesReturned)
				{
				username := StrGet(pBuffer, "UTF-16")
				DllCall("Wtsapi32\WTSFreeMemory", "Ptr", pBuffer)
				if username != ""
					{
					if ActiveUserNameorSessionId = ""
						UserName_Sessions.Push(sessionId "`t" username)
					else
						{
						if IsNumber(ActiveUserNameorSessionId)
							{
							if ActiveUserNameorSessionId = sessionId
								UserName_Sessions := username
							}
						else
							{
							if ActiveUserNameorSessionId = username
								UserName_Sessions := sessionId
							}
						}
					}
				}
			}
		}
	DllCall("Wtsapi32\WTSFreeMemory", "Ptr", pInfo)
	return UserName_Sessions
	}

; --------------------------------------------------------------- NamedPipe RPC class [ HOST & CLIENT ] ---------------------------------------------------------------
class NamedPipeRpc
	{
	static END_MARKER := "__END__"
	static DefaultSddl := "D:(A;;GA;;;SY)(A;;GA;;;BA)"
	; ===========================================================================
	; Public: Host/server side
	; ===========================================================================
	; Message: string
	; PipeName: "PipeName"
	; TimeoutMs: overall timeout in ms
	; Sddl: optional override
	; ===========================================================================
	static Send(Message := "", PipeName := "PipeName", TimeoutMs := 5000, Sddl := "")
		{
		endMarker := this.END_MARKER
		PipeFullName := "\\.\pipe\" . PipeName
		if (Sddl = "")
			Sddl := this.DefaultSddl
		saObj := ""
		hPipe := 0
		f := ""
		try
			{
			saObj := this.MakePipeSA(Sddl)
			hPipe := DllCall("kernel32.dll\CreateNamedPipeW", "WStr", PipeFullName, "UInt", 0x00000003, "UInt", 0x00000000, "UInt", 255, "UInt", 4096, "UInt", 4096, "UInt", 0, "Ptr",  saObj.sa.Ptr, "Ptr")
			if (hPipe = -1)
				return "CreateNamedPipe failed"
			if !DllCall("kernel32.dll\ConnectNamedPipe", "Ptr", hPipe, "Ptr", 0)
				{
				err := DllCall("kernel32.dll\GetLastError")
				if (err != 535)
					return "ConnectNamedPipe failed"
				}
			f := FileOpen(hPipe, "h")
			try
				{
				for line in StrSplit(Message, "`n", "`r")
					f.Write(line . "`n")
				f.Write(endMarker . "`n")
				DllCall("kernel32.dll\FlushFileBuffers", "Ptr", hPipe)
				}
			catch
				return "Write failed"
			start := A_TickCount
			reply := ""
			loop
				{
				if (A_TickCount - start > TimeoutMs)
					{
					reply := "Timeout waiting for reply"
					break
					}
				line := this.ReadLineWithTimeout(f, 1000)
				if (line = "__TIMEOUT__")
					continue
				if InStr(line, "__ERROR__:")
					{
					reply := "ReadLine failed"
					break
					}
				if (line = endMarker)
					break
				reply .= line . "`r`n"
				}
			return RTrim(reply, "`r`n")
			}
		finally
			{
			if IsObject(f)
				f.Close()
			if (hPipe && hPipe != -1)
				{
				DllCall("kernel32.dll\DisconnectNamedPipe", "Ptr", hPipe)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hPipe)
				}
			this.FreePipeSA(saObj)
			}
		}
	; ===========================================================================
	; Public: Client side
	; ===========================================================================
	; ComputerName: "." of remote host
	; PipeName: "PipeName"
	; Reply: string
	; TimeoutSec: "" of seconds
	; ===========================================================================
	static Receive(ComputerName := ".", PipeName := "PipeName", Reply := "", TimeoutSec := "")
		{
		endMarker := this.END_MARKER
		if (ComputerName = A_ComputerName)
			ComputerName := "."
		PipeFullName := "\\" . ComputerName . "\pipe\" . PipeName
		waitMs := TimeoutSec ? (TimeoutSec * 1000) : 0
		start := A_TickCount
		while !DllCall("kernel32.dll\WaitNamedPipeW", "WStr", PipeFullName, "UInt", 200)
			{
			if (TimeoutSec && (A_TickCount - start > waitMs))
				return "Timeout waiting for host"
			Sleep 50
			}
		hPipe := -1
		while true
			{
			hPipe := DllCall("kernel32.dll\CreateFileW", "WStr", PipeFullName, "UInt", 0xC0000000, "UInt", 0, "Ptr",  0, "UInt", 3, "UInt", 0, "Ptr",  0, "Ptr")
			if (hPipe != -1)
				break
			if (TimeoutSec && (A_TickCount - start > waitMs))
				return "Open pipe failed"
			DllCall("kernel32.dll\WaitNamedPipeW", "WStr", PipeFullName, "UInt", 200)
			Sleep 50
			}
		f := ""
		try
			{
			f := FileOpen(hPipe, "h")
			msg := ""
			readStart := A_TickCount
			perReadTimeout := TimeoutSec ? 1000 : 60000
			loop
				{
				if (TimeoutSec && (A_TickCount - readStart > waitMs))
					{
					msg := "Timeout reading message"
					break
					}
				line := this.ReadLineWithTimeout(f, perReadTimeout)
				if (line = "__TIMEOUT__")
					continue
				if InStr(line, "__ERROR__:")
					{
					msg := "ReadLine failed"
					break
					}
				if (line = endMarker)
					break
				msg .= line . "`r`n"
				}
			try
				{
				for line in StrSplit(Reply, "`n", "`r")
					f.Write(line . "`n")
				f.Write(endMarker . "`n")
				DllCall("kernel32.dll\FlushFileBuffers", "Ptr", hPipe)
				}
			catch
				return "Write reply failed"
			return RTrim(msg, "`r`n")
			}
		finally
			{
			if IsObject(f)
				f.Close()
			if (hPipe != -1)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hPipe)
			}
		}
	; ===========================================================================
	; Private: ReadLine with PeekNamedPipe (static buffer)
	; ===========================================================================
	; clear := true -> buffer entry for handle removed
	; ===========================================================================
	static ReadLineWithTimeout(f, timeoutMs := 1000, pollMs := 10, encoding := "UTF-8", clear := false)
		{
		static PipeLineBuf := Map()
		h := f.Handle
		if clear
			{
			if PipeLineBuf.Has(h)
				PipeLineBuf.Delete(h)
			return ""
			}
		if !PipeLineBuf.Has(h)
			PipeLineBuf[h] := ""
		start := A_TickCount
		while (A_TickCount - start < timeoutMs)
			{
			buf := PipeLineBuf[h]
			if (pos := InStr(buf, "`n"))
				{
				line := SubStr(buf, 1, pos - 1)
				PipeLineBuf[h] := SubStr(buf, pos + 1)
				return RTrim(line, "`r")
				}
			bytesAvail := 0
			ok := DllCall("kernel32.dll\PeekNamedPipe", "Ptr", h, "Ptr", 0, "UInt", 0, "Ptr", 0, "UInt*", &bytesAvail, "Ptr", 0)
			if !ok
				{
				err := DllCall("kernel32.dll\GetLastError")
				if (err = 109)
					{
					PipeLineBuf.Delete(h)
					return this.END_MARKER
					}
				this.WriteLog("e", "kernel32.dll\PeekNamedPipe: " err)
				return "__ERROR__:" err
				}
			if (bytesAvail = 0)
				{
				Sleep pollMs
				continue
				}
			toRead := (bytesAvail > 4096) ? 4096 : bytesAvail
			raw := Buffer(toRead, 0)
			bytesRead := 0
			ok := DllCall("kernel32.dll\ReadFile", "Ptr", h, "Ptr", raw.Ptr, "UInt", toRead, "UInt*", &bytesRead, "Ptr", 0)
			if !ok
				{
				err := DllCall("kernel32.dll\GetLastError")
				if (err = 109)
					{
					leftover := PipeLineBuf[h]
					PipeLineBuf.Delete(h)
					return (leftover != "") ? RTrim(leftover, "`r`n") : ""
					}
				this.WriteLog("e", "kernel32.dll\ReadFile: " err)
				return "__ERROR__:" err
				}
			if (bytesRead > 0)
				PipeLineBuf[h] .= StrGet(raw, bytesRead, encoding)
			}
		return "__TIMEOUT__"
		}
    ; ===========================================================================
	; Private: SECURITY_ATTRIBUTES from SDDL
	; ===========================================================================
	static MakePipeSA(sddl)
		{
		sdPtr := 0
		if !DllCall("advapi32.dll\ConvertStringSecurityDescriptorToSecurityDescriptorW", "WStr", sddl, "UInt", 1, "Ptr*", &sdPtr, "Ptr", 0)
			throw OSError()
		saSize := (A_PtrSize = 8) ? 24 : 12
		sa := Buffer(saSize, 0)
		NumPut("UInt", saSize, sa, 0)
		if (A_PtrSize = 8)
			{
			NumPut("Ptr", sdPtr, sa, 8)
			NumPut("Int", 0, sa, 16)
			}
		else
			{
			NumPut("Ptr", sdPtr, sa, 4)
			NumPut("Int", 0, sa, 8)
			}
		return { sa: sa, sdPtr: sdPtr }
		}
	static FreePipeSA(saObj)
		{
		if (IsObject(saObj) && saObj.HasProp("sdPtr") && saObj.sdPtr)
			DllCall("kernel32.dll\LocalFree", "Ptr", saObj.sdPtr, "Ptr")
		}
	static WriteLog(Status := "i", Message := "")
		{
		SplitPath A_ScriptFullPath,, &Dir,, &NoExt_FileName
		LogFile := Dir "\" NoExt_FileName "_NamedPipeRpc.log"
		timestamp := FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss")
		switch Status
			{
			case "i":
				FileAppend timestamp "`t[ INFO  ]`t" Message "`r`n", LogFile
			case "w":
				FileAppend timestamp "`t[ WARN  ]`t" Message "`r`n", LogFile
			case "e":
				FileAppend timestamp "`t[ ERROR ]`t" Message "`r`n", LogFile
			case "x":
				FileAppend "---------- --------`t---------`t--------------------------------------------------`r`n" , LogFile
			}
		}
	}

; --------------------------------------------------------------- MyService Class [ HOST ] ---------------------------------------------------------------
class MyService
	{
	static Start()
		{
		svc := MyService()
		svc.SvcStart()
		}
	__New()
		{
		this.SERVICE_CONTROL_STOP        := 0x00000001
		this.SERVICE_STOPPED             := 0x00000001
		this.SERVICE_START_PENDING       := 0x00000002
		this.SERVICE_STOP_PENDING        := 0x00000003
		this.SERVICE_RUNNING             := 0x00000004
		this.SERVICE_ACCEPT_STOP         := 0x00000001
		this.SERVICE_WIN32_OWN_PROCESS   := 0x00000010
		this.NO_ERROR                    := 0x0
		this.ServerName      := A_Args[1]
		this.NameValue       := SubStr(A_ScriptName, 1, -4)
		this.SvcStatus       := Buffer(32, 0)
		this.DispatchTable   := Buffer(A_PtrSize * 2 + A_PtrSize, 0)
		this.SvcMainCB       := 0
		this.SvcCtrlCB       := 0
		this.TimerCB         := 0
		this.UINT_PTR        := 0
		this.hSvcStopEvent   := 0
		this.SvcStatusHandle := 0
		this.ProcessID       := 0
		}
	__Delete()
		{
		this.Cleanup()
		}
	SvcStart()
		{
		this.WriteLog("i", "ServerName: " this.ServerName)
		this.SvcMainCB := CallbackCreate(ObjBindMethod(this, "SvcMain"), "Fast", 2)
		NumPut("Ptr", StrPtr(this.NameValue), this.DispatchTable, 0)
		NumPut("Ptr", this.SvcMainCB, this.DispatchTable, A_PtrSize)
		this.WriteLog("i", "RegisterCallback MainAddress: " this.SvcMainCB)
		result := DllCall("advapi32.dll\StartServiceCtrlDispatcherW", "Ptr", this.DispatchTable)
		this.WriteLog("i", "StartServiceCtrlDispatcher: " result)
		this.Cleanup()
		this.WriteLog("x", "")
		}
	SvcMain(dwNumServicesArgs := 0, lpServiceArgVectors := 0)
		{
		Critical
		this.SvcCtrlCB := CallbackCreate(ObjBindMethod(this, "SvcCtrlHandler"), "Fast", 1)
		this.SvcStatusHandle := DllCall("advapi32.dll\RegisterServiceCtrlHandlerW", "Str", this.NameValue, "Ptr", this.SvcCtrlCB, "Ptr")
		this.WriteLog("i", "RegisterServiceCtrlHandler: " this.SvcStatusHandle)
		NumPut("UInt", this.SERVICE_WIN32_OWN_PROCESS, this.SvcStatus, 0)
		this.ReportSvcStatus(this.SERVICE_START_PENDING, this.NO_ERROR, 3000)
		this.hSvcStopEvent := DllCall("CreateEventW", "Ptr", 0, "Int", true, "Int", false, "Ptr", 0, "Ptr")
		this.WriteLog("i", "CreateEvent: " this.hSvcStopEvent)
		if !this.hSvcStopEvent
			{
			this.ReportSvcStatus(this.SERVICE_STOPPED, this.NO_ERROR, 0)
			return
			}
		this.ReportSvcStatus(this.SERVICE_RUNNING, this.NO_ERROR, 0)
		this.TimerCB := CallbackCreate(ObjBindMethod(this, "UserProg"), "Fast", 4)
		this.UINT_PTR := DllCall("User32.dll\SetTimer", "Ptr", 0, "Ptr", 1, "UInt", 1000, "Ptr", this.TimerCB, "Ptr")
		this.WriteLog("i", "SetTimer: " this.UINT_PTR)
		Sleep 1000
		DllCall("WaitForSingleObject", "Ptr", this.hSvcStopEvent, "UInt", 0xFFFFFFFF)
		this.ReportSvcStatus(this.SERVICE_STOPPED, this.NO_ERROR, 0)
		}
	SvcCtrlHandler(dwCtrl)
		{
		Critical
		if (dwCtrl = this.SERVICE_CONTROL_STOP)
			{
			this.ReportSvcStatus(this.SERVICE_STOP_PENDING, this.NO_ERROR, 0)
			DllCall("SetEvent", "Ptr", this.hSvcStopEvent)
			}
		this.WriteLog("i", "SvcCtrlHandler: " dwCtrl)
		}
	ReportSvcStatus(CurrentState, Win32ExitCode, WaitHint)
		{
		NumPut("UInt", CurrentState, this.SvcStatus, 4)
		NumPut("UInt", Win32ExitCode, this.SvcStatus, 12)
		NumPut("UInt", WaitHint, this.SvcStatus, 24)
		if (CurrentState = this.SERVICE_START_PENDING)
			NumPut("UInt", 0, this.SvcStatus, 8)
		else
			NumPut("UInt", this.SERVICE_ACCEPT_STOP, this.SvcStatus, 8)
		if (CurrentState = this.SERVICE_RUNNING || CurrentState = this.SERVICE_STOPPED)
			NumPut("UInt", 0, this.SvcStatus, 20)
		else
			NumPut("UInt", NumGet(this.SvcStatus, 20, "UInt") + 1, this.SvcStatus, 20)
		result := DllCall("advapi32.dll\SetServiceStatus", "Ptr", this.SvcStatusHandle, "Ptr", this.SvcStatus)
		this.WriteLog("i", "SetServiceStatus: " result)
		}
	UserProg(hwnd := 0, msg := 0, idEvent := 0, time := 0)
		{
		Critical
		if (this.ProcessID = 0 OR !ProcessExist(this.ProcessID))
			{
			Run(A_ScriptFullPath . " " . this.ServerName . " com",,, &PID)
			this.ProcessID := PID
			this.WriteLog("i", "Starting UserProg with Process Id: " this.ProcessID)
			}
		}
	Cleanup()
		{
		if this.UINT_PTR
			{
			DllCall("User32.dll\KillTimer", "Ptr", 0, "Ptr", this.UINT_PTR)
			this.UINT_PTR := 0
			}
		if this.hSvcStopEvent
			{
			DllCall("CloseHandle", "Ptr", this.hSvcStopEvent)
			this.hSvcStopEvent := 0
			}
		if this.TimerCB
			{
			CallbackFree(this.TimerCB)
			this.TimerCB := 0
			}
		if this.SvcCtrlCB
			{
			CallbackFree(this.SvcCtrlCB)
			this.SvcCtrlCB := 0
			}
		if this.SvcMainCB
			{
			CallbackFree(this.SvcMainCB)
			this.SvcMainCB := 0
			}
		this.WriteLog("i", "Cleanup Complete.")
		}
	WriteLog(Status := "i", Message := "")
		{
		SplitPath A_ScriptFullPath,, &Dir,, &NoExt_FileName
		LogFile := Dir "\" NoExt_FileName "_service.log"
		timestamp := FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss")
		switch Status
			{
			case "i":
				FileAppend timestamp "`t[ INFO  ]`t" Message "`r`n", LogFile
			case "w":
				FileAppend timestamp "`t[ WARN  ]`t" Message "`r`n", LogFile
			case "e":
				FileAppend timestamp "`t[ ERROR ]`t" Message "`r`n", LogFile
			case "x":
				FileAppend "---------- --------`t---------`t--------------------------------------------------`r`n" , LogFile
			}
		}
	}

; --------------------------------------------------------------- WinService Class for controling Windows Services [ CLIENT ] ---------------------------------------------------------------
class WinService
	{
	static SERVICE_QUERY_STATUS			:= 0x4
	static SERVICE_START				:= 0x10
	static SERVICE_STOP					:= 0x20
	static SERVICE_ALL_ACCESS			:= 0xF01FF
	static SC_MANAGER_CONNECT			:= 0x1
	static SC_MANAGER_CREATE_SERVICE	:= 0x2
	static SERVICE_WIN32_OWN_PROCESS	:= 0x10
	static SERVICE_ERROR_NORMAL			:= 0x1
	static State(ComputerName := "", ServiceName := "", textResult := false)
		{
		if ComputerName
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Ptr", StrPtr(ComputerName), "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		else
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32.dll\OpenServiceW", "Ptr", scm, "Str", ServiceName, "UInt", this.SERVICE_QUERY_STATUS)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
			return err
			}
		buf := Buffer(28, 0)
		ok := DllCall("advapi32.dll\QueryServiceStatus", "Ptr", hSvc, "Ptr", buf.ptr)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
		if !ok
			return A_LastError
		state := NumGet(buf, 4, "UInt")
		if !textResult
			return state
		return Map(
			1, "Stopped",
			2, "Start Pending",
			3, "Stop Pending",
			4, "Running",
			5, "Continue Pending",
			6, "Pause Pending",
			7, "Paused"
			).Get(state, "Unknown")
		}
	static Start(ComputerName := "", ServiceName := "")
		{
		if ComputerName
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Ptr", StrPtr(ComputerName), "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		else
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32.dll\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_START)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
			return err
			}
		ok := DllCall("advapi32.dll\StartServiceW", "Ptr", hSvc, "UInt", 0, "Ptr", 0)
		err := ok ? ok : A_LastError
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
		return err
		}
	static Stop(ComputerName := "", ServiceName := "")
		{
		if ComputerName
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Ptr", StrPtr(ComputerName), "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		else
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32.dll\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_STOP)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
			return err
			}
		status := Buffer((A_PtrSize = 4) ? 28 : 32, 0)
		ok := DllCall("advapi32.dll\ControlService", "Ptr", hSvc, "UInt", 1, "Ptr", status.ptr)
		err := ok ? ok : A_LastError
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
		return err
		}
	static Add(ComputerName := "", ServiceName := "", BinaryPath := "", StartType := "", DisplayName := "", ServiceStartName := "", Password := "")
		{
		StartType := (StartType = "Auto" || StartType = "Automatic") ? 0x2
				: (StartType = "Demand" || StartType = "OnDemand") ? 0x3
				: 0x4
		if ComputerName
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Ptr", StrPtr(ComputerName), "Int", 0, "UInt", this.SC_MANAGER_CREATE_SERVICE)
		else
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CREATE_SERVICE)
		if ServiceStartName
			svc := DllCall("advapi32.dll\CreateServiceW","Ptr", scm, "Ptr", StrPtr(ServiceName), "Ptr", StrPtr(DisplayName ? DisplayName : ServiceName), "UInt", this.SERVICE_ALL_ACCESS, "UInt", this.SERVICE_WIN32_OWN_PROCESS, "UInt", StartType, "UInt", this.SERVICE_ERROR_NORMAL, "Ptr", StrPtr(BinaryPath), "Ptr", 0, "UInt", 0, "Ptr", 0, "Ptr", StrPtr(ServiceStartName), "Ptr", StrPtr(Password))
		else
			svc := DllCall("advapi32.dll\CreateServiceW","Ptr", scm, "Ptr", StrPtr(ServiceName), "Ptr", StrPtr(DisplayName ? DisplayName : ServiceName), "UInt", this.SERVICE_ALL_ACCESS, "UInt", this.SERVICE_WIN32_OWN_PROCESS, "UInt", StartType, "UInt", this.SERVICE_ERROR_NORMAL, "Ptr", StrPtr(BinaryPath), "Ptr", 0, "UInt", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0)
		result := A_LastError ? svc "," A_LastError : 1
		hSvc := DllCall("advapi32.dll\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_START)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
			return err
			}
		ok := DllCall("advapi32.dll\StartServiceW", "Ptr", hSvc, "UInt", 0, "Ptr", 0)
		err := ok ? ok : A_LastError
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", svc)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
		return err
		}
	static Delete(ComputerName := "", ServiceName := "")
		{
		if ComputerName
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Ptr", StrPtr(ComputerName), "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		else
			scm := DllCall("advapi32.dll\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32.dll\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_ALL_ACCESS)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
			return err
			}
		ok := DllCall("advapi32.dll\DeleteService", "Ptr", hSvc)
		err := ok ? ok : A_LastError
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32.dll\CloseServiceHandle", "Ptr", scm)
		return err
		}
	}

; --------------------------------------------------------------- Class for creating a new interactive SYSTEM or USER process ---------------------------------------------------------------
class SystemProcess
	{
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Create an interactive SYSTEM ( CreateProcessAsSystem / CreateProcessAsSystemInteractive / CreateProcessAsSystemWinlogon ) or USER ( CreateProcessAsUserInteractive ) process.
	; cmd                       := Executable (path).
	; args                      := Command line arguments.
	; workingDir                := Working directory.
	; processName               := Set default to "explorer.exe" when empty or use another existing user process to get the session Id from.
	; wait                      := 1 to wait for exit, 0 to return PID.
	; ActiveUserNameorSessionId := Empty when using processName or you can either use the session Id number or username.
	; outPid                    := ByRef to return the PID immediately.
	; outHandle					;= ByRef to return the Handle for retrieving the process exitcode if wait is 0.
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	; Return error Codes:
	; -1						:= Command line is empty.
	; -2						:= Executable not found.
	; -3						:= Running process not found.
	; -4						:= SessionId function failed
	; -5						:= GetWinlogonPid function failed.
	; -6						:= OpenProcess function failed.
	; -7						:= OpenProcessToken function failed.
	; -8						:= DuplicateTokenEx function failed.
	; -9						:= CreateEnvironmentBlock function failed.
	; -10						:= CreateProcessAsUserW function failed.
	; -11						:= GetExitCodeProcess failed
	; -12						:= WTSGetActiveConsoleSessionId failed / no console
	; ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
	static CreateProcessAsSystem(cmd := "", args := "", workingDir := "", wait := 1, &outPid := 0, &outHandle := 0)
		{
		if !SystemProcess.SplitExeAndArgs(cmd, workingDir, &exe, &inlineArgs)
			return -1
		if !FileExist(exe)
			return -2
		fullArgs := inlineArgs
		if (args != "")
			fullArgs := (fullArgs != "" ? fullArgs " " : "") args
		appName := exe
		fullCmdLine := SystemProcess.QuoteIfNeeded(exe)
		if (fullArgs != "")
			fullCmdLine .= " " fullArgs
		hProcess := 0
		hThread := 0
		try
			{
			cmdBuf := Buffer((StrLen(fullCmdLine) + 1) * 2, 0)
			StrPut(fullCmdLine, cmdBuf, "UTF-16")
			startupInfo := Buffer(A_PtrSize = 8 ? 104 : 68, 0)
			NumPut("UInt", startupInfo.Size, startupInfo, 0)
			processInfo := Buffer(2 * A_PtrSize + 2 * 4, 0)
			flags := 0x00000000
			wdType := (workingDir = "") ? "Ptr" : "WStr"
			wdVal  := (workingDir = "") ? 0 : workingDir
			if !DllCall("kernel32.dll\CreateProcessW", "Ptr", 0, "Ptr", cmdBuf, "Ptr", 0, "Ptr", 0, "Int", 0, "UInt", flags, "Ptr", 0, wdType, wdVal, "Ptr", startupInfo, "Ptr", processInfo)
				return -10
			hProcess := NumGet(processInfo, 0, "Ptr")
			hThread := NumGet(processInfo, A_PtrSize, "Ptr")
			newPid := NumGet(processInfo, 2 * A_PtrSize, "UInt")
			outPid := newPid
			if (wait != 1)
				outHandle := hProcess
			else
				outHandle := 0
			if (hThread)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hThread)
			if (wait != 1)
				return newPid
			loop
				{
				if !DllCall("kernel32.dll\GetExitCodeProcess", "Ptr", hProcess, "UInt*", &exitCode := 0)
					return -11
				if (exitCode = 259)
					Sleep 50
				else
					break
				}
			return exitCode
			}
		finally
			{
			if (hProcess)
				{
				if (wait = 1)
					DllCall("kernel32.dll\CloseHandle", "Ptr", hProcess)
				}
			}
		}
	static CreateProcessAsSystemWinlogon(cmd := "", args := "", workingDir := "", wait := 1, &outPid := 0, &outHandle := 0)
		{
		if !SystemProcess.SplitExeAndArgs(cmd, workingDir, &exe, &inlineArgs)
			return -1
		if !FileExist(exe)
			return -2
		SystemProcess.EnablePrivilege("SeDebugPrivilege")
		SystemProcess.EnablePrivilege("SeAssignPrimaryTokenPrivilege")
		SystemProcess.EnablePrivilege("SeIncreaseQuotaPrivilege")
		sessionId := SystemProcess.GetWinlogonScreenSessionId()
		if (!sessionId)
			return -12
		fullArgs := inlineArgs
		if (args != "")
			fullArgs := (fullArgs != "" ? fullArgs " " : "") args
		appName := exe
		fullCmdLine := SystemProcess.QuoteIfNeeded(exe)
		if (fullArgs != "")
			fullCmdLine .= " " fullArgs
		hWinlogonProc := 0
		hWinlogonTok := 0
		hDupToken := 0
		lpEnvironment := 0
		hProcess := 0
		hThread := 0
		try
			{
			winlogonPid := SystemProcess.GetWinlogonPid(sessionId)
			if (winlogonPid = -1 || winlogonPid = 0)
				return -5
			PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
			hWinlogonProc := DllCall("kernel32.dll\OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "Int", 0, "UInt", winlogonPid, "Ptr")
			if !hWinlogonProc
				return -6
			TOKEN_DUPLICATE := 0x2
			TOKEN_ASSIGN_PRIMARY := 0x1
			TOKEN_QUERY := 0x8
			if !DllCall("advapi32.dll\OpenProcessToken", "Ptr", hWinlogonProc, "UInt", TOKEN_DUPLICATE|TOKEN_ASSIGN_PRIMARY|TOKEN_QUERY, "Ptr*", &hWinlogonTok := 0)
				return -7
			if !DllCall("advapi32.dll\DuplicateTokenEx", "Ptr", hWinlogonTok, "UInt", 0xF01FF, "Ptr", 0, "UInt", 2, "UInt", 1, "Ptr*", &hDupToken := 0)
				return -8
			if !DllCall("Userenv.dll\CreateEnvironmentBlock", "Ptr*", &lpEnvironment := 0, "Ptr", hDupToken, "Int", 0) || !lpEnvironment
				return -9
			cmdBuf := Buffer((StrLen(fullCmdLine) + 1) * 2, 0)
			StrPut(fullCmdLine, cmdBuf, "UTF-16")
			startupInfo := Buffer(A_PtrSize = 8 ? 104 : 68, 0)
			NumPut("UInt", startupInfo.Size, startupInfo, 0)
			desktop := "winsta0\winlogon"
			NumPut("Ptr", StrPtr(desktop), startupInfo, (A_PtrSize = 8) ? 16 : 8)
			NumPut("UInt", 1, startupInfo, (A_PtrSize = 8) ? 60 : 44)
			NumPut("UShort", 5, startupInfo, (A_PtrSize = 8) ? 64 : 48)
			processInfo := Buffer(2 * A_PtrSize + 2 * 4, 0)
			CREATE_UNICODE_ENVIRONMENT := 0x400
			flags := CREATE_UNICODE_ENVIRONMENT
			wdType := (workingDir = "") ? "Ptr" : "WStr"
			wdVal  := (workingDir = "") ? 0 : workingDir
			if !DllCall("Advapi32.dll\CreateProcessAsUserW", "Ptr", hDupToken, "WStr", appName, "Ptr", cmdBuf, "Ptr", 0, "Ptr", 0, "Int", 0, "UInt", flags, "Ptr", lpEnvironment, wdType, wdVal, "Ptr", startupInfo, "Ptr", processInfo)
				return -10
			hProcess := NumGet(processInfo, 0, "Ptr")
			hThread := NumGet(processInfo, A_PtrSize, "Ptr")
			newPid := NumGet(processInfo, 2 * A_PtrSize, "UInt")
			outPid := newPid
			if (wait != 1)
				outHandle := hProcess
			else
				outHandle := 0
			if (hThread)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hThread)
			if (wait != 1)
				return newPid
			loop
				{
				if !DllCall("kernel32.dll\GetExitCodeProcess", "Ptr", hProcess, "UInt*", &exitCode := 0)
					return -11
				if (exitCode = 259)
					Sleep 50
				else
					break
				}
			return exitCode
			}
		finally
			{
			if (hProcess)
				{
				if (wait = 1)
					DllCall("kernel32.dll\CloseHandle", "Ptr", hProcess)
				}
			if (lpEnvironment)
				DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
			if (hDupToken)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hDupToken)
			if (hWinlogonTok)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hWinlogonTok)
			if (hWinlogonProc)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hWinlogonProc)
			}
		}
	static CreateProcessAsSystemInteractive(cmd := "", args := "", workingDir := "", processName := "explorer.exe", wait := 1, ActiveUserNameorSessionId := "", &outPid := 0, &outHandle := 0)
		{
		if !SystemProcess.SplitExeAndArgs(cmd, WorkingDir, &exe, &inlineArgs)
			return -1
		if !FileExist(exe)
			return -2
		SystemProcess.EnablePrivilege("SeDebugPrivilege")
		SystemProcess.EnablePrivilege("SeAssignPrimaryTokenPrivilege")
		SystemProcess.EnablePrivilege("SeIncreaseQuotaPrivilege")
		if ActiveUserNameorSessionId = ""
			{
			if !(ProcessId := ProcessExist(processName))
				return -3
			DllCall("kernel32.dll\ProcessIdToSessionId", "UInt", ProcessId, "Ptr*", &SessionId := 0)
			if !SessionId
				return -4
			}
		else
			{
			if IsNumber(ActiveUserNameorSessionId)
				SessionId := ActiveUserNameorSessionId
			else
				{
				SessionId := SystemProcess.GetActiveSessionId(ActiveUserNameorSessionId)
				if !SessionId
					return -4
				}
			}
		fullArgs := inlineArgs
		if (args != "")
			fullArgs := (fullArgs != "" ? fullArgs " " : "") args
		appName := exe
		fullCmdLine := SystemProcess.QuoteIfNeeded(exe)
		if (fullArgs != "")
			fullCmdLine .= " " fullArgs
		hWinlogonProc := 0
		hWinlogonTok := 0
		hDupToken := 0
		lpEnvironment := 0
		hProcess := 0
		hThread := 0
		try
			{
			winlogonPid := SystemProcess.GetWinlogonPid(SessionId)
			if (winlogonPid = -1 OR winlogonPid = 0)
				return -5
			PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
			hWinlogonProc := DllCall("kernel32.dll\OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "Int", 0, "UInt", winlogonPid, "Ptr")
			if !hWinlogonProc
				return -6
			TOKEN_DUPLICATE := 0x2
			TOKEN_ASSIGN_PRIMARY := 0x1
			TOKEN_QUERY := 0x8
			if !DllCall("advapi32.dll\OpenProcessToken", "Ptr", hWinlogonProc, "UInt", TOKEN_DUPLICATE|TOKEN_ASSIGN_PRIMARY|TOKEN_QUERY, "Ptr*", &hWinlogonTok := 0)
				return -7
			if !DllCall("advapi32.dll\DuplicateTokenEx", "Ptr", hWinlogonTok, "UInt", 0xF01FF, "Ptr", 0, "UInt", 2, "UInt", 1, "Ptr*", &hDupToken := 0)
				return -8
			if !DllCall("Userenv.dll\CreateEnvironmentBlock", "Ptr*", &lpEnvironment := 0, "Ptr", hDupToken, "Int", 0) OR !lpEnvironment
				return -9
			cmdBuf := Buffer((StrLen(fullCmdLine) + 1) * 2, 0)
			StrPut(fullCmdLine, cmdBuf, "UTF-16")
			startupInfo := Buffer(A_PtrSize = 8 ? 104 : 68, 0)
			NumPut("UInt", startupInfo.Size, startupInfo)
			desktop := "winsta0\default"
			NumPut("Ptr", StrPtr(desktop), startupInfo, (A_PtrSize = 8) ? 16 : 8)
			NumPut("UInt", 1, startupInfo, (A_PtrSize = 8) ? 60 : 44)
			NumPut("UShort", 5, startupInfo, (A_PtrSize = 8) ? 64 : 48)
			processInfo := Buffer(2 * A_PtrSize + 2 * 4, 0)
			flags := 0x00000400
			wdType := (workingDir = "") ? "Ptr" : "WStr"
			wdVal  := (workingDir = "") ? 0 : workingDir
			CREATE_UNICODE_ENVIRONMENT := 0x400
			if !DllCall("Advapi32.dll\CreateProcessAsUserW", "Ptr", hDupToken, "WStr", appName, "Ptr", cmdBuf, "Ptr", 0, "Ptr", 0, "Int", 0, "UInt", flags, "Ptr", lpEnvironment, wdType, wdVal, "Ptr", startupInfo, "Ptr", processInfo)
				return -10
			hProcess := NumGet(processInfo, 0, "Ptr")
			hThread  := NumGet(processInfo, A_PtrSize, "Ptr")
			newPid := NumGet(processInfo, 2 * A_PtrSize, "UInt")
			outPid := newPid
			if (wait != 1)
				outHandle := hProcess
			else
				outHandle := 0
			if (hThread)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hThread)
			if (wait != 1)
				return newPid
			loop
				{
				if !DllCall("kernel32.dll\GetExitCodeProcess", "Ptr", hProcess, "UInt*", &exitCode := 0)
					return -11
				if exitCode = 259
					Sleep 50
				else
					break
				}
			return exitCode
			}
		finally
			{
			if (hProcess)
				{
				if (wait = 1)
					DllCall("kernel32.dll\CloseHandle", "Ptr", hProcess)
				}
			if (lpEnvironment)
				DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
			if (hDupToken)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hDupToken)
			if (hWinlogonTok)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hWinlogonTok)
			if (hWinlogonProc)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hWinlogonProc)			
			}
		}
	static CreateProcessAsUserInteractive(cmd := "", args := "", workingDir := "", processName := "explorer.exe", wait := 1, ActiveUserNameorSessionId := "", &outPid := 0, &outHandle := 0)
		{
		if !SystemProcess.SplitExeAndArgs(cmd, WorkingDir, &exe, &inlineArgs)
			return -1
		if !FileExist(exe)
			return -2
		ProcessId := ""
		if ActiveUserNameorSessionId = ""
			{
			if !(ProcessId := ProcessExist(processName))
				return -3
			DllCall("ProcessIdToSessionId", "UInt", ProcessId, "Ptr*", &SessionId := 0)
			if !SessionId
				return -4
			}
		else
			{
			if IsNumber(ActiveUserNameorSessionId)
				SessionId := ActiveUserNameorSessionId
			else
				{
				SessionId := SystemProcess.GetActiveSessionId(ActiveUserNameorSessionId)
				if !SessionId
					return -4
				}
			}
		fullArgs := inlineArgs
		if (args != "")
			fullArgs := (fullArgs != "" ? fullArgs " " : "") args
		appName := exe
		fullCmdLine := SystemProcess.QuoteIfNeeded(exe)
		if (fullArgs != "")
			fullCmdLine .= " " fullArgs
		hProcess := 0
		lpEnvironment := 0
		hToken := 0
		hThread := 0
		try
			{
			if !DllCall("wtsapi32.dll\WTSQueryUserToken", "Int", SessionId, "Ptr*", &hToken := 0)
				{
				if !ProcessId
					return -3
				PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
				hProc := DllCall("kernel32.dll\OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "Int", 0, "UInt", ProcessId, "Ptr")
				if !hProc
					return -6
				TOKEN_DUPLICATE := 0x2
				TOKEN_ASSIGN_PRIMARY := 0x1
				TOKEN_QUERY := 0x8
				if !DllCall("advapi32.dll\OpenProcessToken", "Ptr", hProc, "UInt", TOKEN_DUPLICATE|TOKEN_ASSIGN_PRIMARY|TOKEN_QUERY, "Ptr*", &hToken := 0)
					return -7
				hDup := 0
				if !DllCall("advapi32.dll\DuplicateTokenEx", "Ptr", hToken, "UInt", 0xF01FF, "Ptr", 0, "UInt", 2, "UInt", 1, "Ptr*", &hDup := 0)
					return -8
				DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
				hToken := hDup
				DllCall("kernel32.dll\CloseHandle", "Ptr", hProc)
				}
			if !DllCall("Userenv.dll\CreateEnvironmentBlock", "Ptr*", &lpEnvironment := 0, "Ptr", hToken, "Int", 0) OR !lpEnvironment
				return -9
			cmdBuf := Buffer((StrLen(fullCmdLine) + 1) * 2, 0)
			StrPut(fullCmdLine, cmdBuf, "UTF-16")
			startupInfo := Buffer(A_PtrSize = 8 ? 104 : 68, 0)
			NumPut("UInt", startupInfo.Size, startupInfo, 0)
			desktop := "winsta0\default"
			NumPut("Ptr", StrPtr(desktop), startupInfo, (A_PtrSize = 8) ? 16 : 8)
			NumPut("UInt", 1, startupInfo, (A_PtrSize = 8) ? 60 : 44)
			NumPut("UShort", 5, startupInfo, (A_PtrSize = 8) ? 64 : 48)
			processInfo := Buffer(2 * A_PtrSize + 2 * 4, 0)
			flags := 0x00000400
			wdType := (workingDir = "") ? "Ptr" : "WStr"
			wdVal  := (workingDir = "") ? 0 : workingDir
			if !DllCall("Advapi32.dll\CreateProcessAsUserW", "Ptr", hToken, "WStr", appName, "Ptr", cmdBuf, "Ptr", 0, "Ptr", 0, "Int", 0, "UInt", flags, "Ptr", lpEnvironment, wdType, wdVal, "Ptr", startupInfo, "Ptr", processInfo)
				return -10
			hProcess := NumGet(processInfo, 0, "Ptr")
			hThread  := NumGet(processInfo, A_PtrSize, "Ptr")
			newPid   := NumGet(processInfo, 2 * A_PtrSize, "UInt")
			outPid := newPid
			if (wait != 1)
				outHandle := hProcess
			else
				outHandle := 0
			if (hThread)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hThread)
			if (wait != 1)
				return newPid			
			loop
				{
				if !DllCall("kernel32.dll\GetExitCodeProcess", "Ptr", hProcess, "UInt*", &exitCode := 0)
					return -11
				if exitCode = 259
					Sleep 50
				else
					break
				}
			return exitCode
			}
		finally
			{
			if (hProcess)
				{
				if (wait = 1)
					DllCall("kernel32.dll\CloseHandle", "Ptr", hProcess)
				}
			if (lpEnvironment)
				DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
			if (hToken)
				DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
			}
		}
	static SplitExeAndArgs(cmdLine, WorkingDir, &exe, &rest)
		{
		cmdLine := Trim(cmdLine)
		exe := "", rest := ""
		if (cmdLine = "")
			return false
		if (SubStr(cmdLine, 1, 1) = '"')
			{
			p := InStr(cmdLine, '"', , 2)
			if (!p)
				{
				cmdLine := StrReplace(cmdLine, '"',,,, 1)
				if FileExist(cmdLine)
					{
					exe := cmdLine
					rest := ""
					return true
					}
				}
			else
				{
				cmdLine := StrReplace(cmdLine, '"',,,, 2)
				if FileExist(cmdLine)
					{
					exe := cmdLine
					rest := LTrim(SubStr(cmdLine, p+1))
					return true
					}
				}
			}
		sp := InStr(cmdLine, " ")
		if (!sp)
			{
			if FileExist(cmdLine)
				{
				exe := cmdLine
				rest := ""
				return true
				}
			}
		if RegExMatch(cmdLine, "i)\.(exe|com|bat|cmd)(?=\s|$)", &m, 1)
			{
			endPos := m.Pos + m.Len - 1
			exeCand := Trim(SubStr(cmdLine, 1, endPos))
			if FileExist(exeCand)
				{
				exe  := exeCand
				rest := LTrim(SubStr(cmdLine, endPos + 1))
				return true
				}
			}
		parts := StrSplit(cmdLine, A_Space)
		cand := parts[1]
		i := 1
		while (i < parts.Length)
			{
			if FileExist(cand)
				{
				exe := cand
				rest := LTrim(SubStr(cmdLine, StrLen(exe) + 1))
				return true
				}
			i += 1
			cand .= " " parts[i]
			}
		if FileExist(cmdLine)
			{
			exe := cmdLine
			rest := ""
			return true
			}
		firstTok := parts[1]
		wd := Trim(workingDir)
		if (wd != "")
			{
			wd := RegExReplace(wd, "[\\/]+$")
			for suffix in ["", ".exe", ".com", ".bat", ".cmd"]
				{
				cand2 := wd "\" firstTok suffix
				if FileExist(cand2)
					{
					exe := cand2
					rest := LTrim(SubStr(cmdLine, StrLen(firstTok) + 1))
					return true
					}
				}
			}
		resolved := SystemProcess.ResolveExePath(firstTok)
		if (resolved != "")
			{
			exe := resolved
			rest := LTrim(SubStr(cmdLine, StrLen(firstTok) + 1))
			return true
			}
		exe := firstTok
		rest := LTrim(SubStr(cmdLine, StrLen(firstTok) + 1))
		return true
		}
	static ResolveExePath(exe)
		{
		if InStr(exe, "\") || InStr(exe, "/") || RegExMatch(exe, "i)^[A-Z]:")
			return exe
		bufSize := 32768
		buf := Buffer(bufSize * 2, 0)
		len := DllCall("kernel32.dll\SearchPathW", "Ptr", 0, "WStr", exe, "WStr", ".exe", "UInt", bufSize, "Ptr", buf, "Ptr*", 0, "UInt")
		if (len = 0)
			return ""
		return StrGet(buf, "UTF-16")
		}
	static QuoteIfNeeded(s) => InStr(s, " ") ? '"' s '"' : s
	static GetWinlogonPid(sessionId)
		{
		snapshot := DllCall("kernel32.dll\CreateToolhelp32Snapshot", "UInt", 0x2, "UInt", 0, "Ptr")
		if snapshot = -1
			return -1
		peSize := (A_PtrSize = 8) ? 568 : 556
		pe := Buffer(peSize, 0)
		NumPut("UInt", peSize, pe)
		if !DllCall("kernel32.dll\Process32FirstW", "Ptr", snapshot, "Ptr", pe)
			{
			DllCall("CloseHandle", "Ptr", snapshot)
			return -1
			}
		loop
			{
			pid := NumGet(pe, 8, "UInt")
			nameOffset := (A_PtrSize = 8) ? 44 : 36
			name := StrGet(pe.Ptr + nameOffset,, "UTF-16")
			if (name = "winlogon.exe")
				{
				if DllCall("kernel32.dll\ProcessIdToSessionId", "UInt", pid, "UInt*", &sid := 0)
					{
					if sid = sessionId
						{
						DllCall("CloseHandle", "Ptr", snapshot)
						return pid
						}
					}
				}
			if !DllCall("kernel32.dll\Process32NextW", "Ptr", snapshot, "Ptr", pe)
				break
			}
		DllCall("CloseHandle", "Ptr", snapshot)
		return 0
		}
	static EnablePrivilege(privilege)
		{
		if !DllCall("advapi32.dll\OpenProcessToken", "Ptr", DllCall("kernel32.dll\GetCurrentProcess", "Ptr"), "UInt", 0x20 | 0x8, "Ptr*", &hToken := 0)
			return false
		luid := Buffer(8)
		if !DllCall("advapi32.dll\LookupPrivilegeValueW", "Ptr", 0, "WStr", privilege, "Ptr", luid)
			return false
		tp := Buffer(16, 0)
		NumPut("UInt", 1, tp, 0)
		NumPut("Int64", NumGet(luid, 0, "Int64"), tp, 4)
		NumPut("UInt", 2, tp, 12)
		DllCall("advapi32.dll\AdjustTokenPrivileges", "Ptr", hToken, "Int", 0, "Ptr", tp, "UInt", 0, "Ptr", 0, "Ptr", 0)
		DllCall("CloseHandle", "Ptr", hToken)
		return true
		}
	static GetActiveSessionId(ActiveUserName := "")
		{
		WTS_CURRENT_SERVER_HANDLE := 0
		pInfo := 0
		count := 0
		ActiveSessionId := 0
		if !DllCall("Wtsapi32.dll\WTSEnumerateSessions", "Ptr", WTS_CURRENT_SERVER_HANDLE, "UInt", 0, "UInt", 1, "Ptr*", &pInfo, "UInt*", &count)
			return ActiveSessionId
		structSize := (A_PtrSize = 8) ? 24 : 12
		stateOffset := (A_PtrSize = 8) ? 16 : 8
		loop count
			{
			pItem := pInfo + (A_Index - 1) * structSize
			sessionId := NumGet(pItem, 0, "UInt")
			state := NumGet(pItem, stateOffset, "Int")
			if (state = 0 && sessionId > 0)
				{
				pBuffer := 0
				bytesReturned := 0
				WTSUserName := 5
				if DllCall("Wtsapi32.dll\WTSQuerySessionInformationW", "Ptr", WTS_CURRENT_SERVER_HANDLE, "UInt", sessionId, "Int", WTSUserName, "Ptr*", &pBuffer, "UInt*", &bytesReturned)
					{
					username := StrGet(pBuffer, "UTF-16")
					DllCall("Wtsapi32.dll\WTSFreeMemory", "Ptr", pBuffer)
					if username != ""
						{
						if ActiveUserName = username
							ActiveSessionId := sessionId
						}
					}
				}
			}
		DllCall("Wtsapi32.dll\WTSFreeMemory", "Ptr", pInfo)
			return ActiveSessionId
		}
	static GetWinlogonScreenSessionId()
		{
		sid := DllCall("kernel32.dll\WTSGetActiveConsoleSessionId", "UInt")
		if (sid != 0xFFFFFFFF)
			{
			if (SystemProcess.GetWinlogonPid(sid) > 0)
				return sid
			}
		WTS_CURRENT_SERVER_HANDLE := 0
		pInfo := 0, count := 0
		if !DllCall("wtsapi32.dll\WTSEnumerateSessionsW", "Ptr", WTS_CURRENT_SERVER_HANDLE, "UInt", 0, "UInt", 1, "Ptr*", &pInfo, "UInt*", &count)
			return 0
		structSize := (A_PtrSize = 8) ? 24 : 12
		nameOffset := (A_PtrSize = 8) ? 8  : 4 
		stateOffset := (A_PtrSize = 8) ? 16 : 8
		bestSid := 0
		bestScore := -1
		loop count
			{
			pItem := pInfo + (A_Index - 1) * structSize
			sid2  := NumGet(pItem, 0, "UInt")
			if (sid2 <= 0)
				continue
			pName := NumGet(pItem, nameOffset, "Ptr")
			wsName := pName ? StrGet(pName, "UTF-16") : ""
			state  := NumGet(pItem, stateOffset, "Int")
			if (SystemProcess.GetWinlogonPid(sid2) <= 0)
				continue
			score := 0
			if (wsName = "Console")
				score += 100
			if (state = 0)
				score += 50
			else if (state = 1)
				score += 40
			score += (1000 - sid2)
			if (score > bestScore)
				{
				bestScore := score
				bestSid := sid2
				}
			}
		DllCall("wtsapi32.dll\WTSFreeMemory", "Ptr", pInfo)
		return bestSid
		}
	}

; --------------------------------------------------------------- psScript Manager Class for directly executing PowerShell code [ HOST ] ---------------------------------------------------------------
class PsScriptManager
	{
	static RootKey        := "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\"
	static ClassName      := "psScript"
	static Assembly       := "psScript, Version=0.0.0.0, Culture=neutral, PublicKeyToken=9c6ab8a0165a1068"
	static RuntimeVersion := "v4.0.30319"
	static CLSID          := "{58628207-D74E-392A-8A85-FDF6107305F0}"
	static Component      := "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"
	static ComInstance := unset
	static Register(SetInstallDir := A_Temp)
		{
		SetCodeBase := "file:///" . StrReplace(SetInstallDir, "\", "/") . "/psScript.DLL"
		if !DirExist(SetInstallDir)
			DirCreate SetInstallDir
		if !FileExist(SetInstallDir . "\psScript.dll")
			FileInstall "psScript.dll", SetInstallDir . "\psScript.dll", 1
		RegWrite this.ClassName, "REG_SZ", this.RootKey . this.ClassName
		RegWrite this.CLSID, "REG_SZ", this.RootKey . this.ClassName . "\CLSID"
		RegWrite this.ClassName, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID
        RegCreateKey this.RootKey . "CLSID\" . this.CLSID . "\Implemented Categories\" . this.Component
		RegWrite "mscoree.dll", "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32"
		RegWrite "Both", "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32", "ThreadingModel"
		RegWrite this.ClassName, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32", "Class"
		RegWrite this.Assembly, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32", "Assembly"
		RegWrite this.RuntimeVersion, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32", "RuntimeVersion"
		RegWrite SetCodeBase, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32", "CodeBase"
		RegWrite this.ClassName, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32\0.0.0.0"
		RegWrite this.Assembly, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32\0.0.0.0", "Assembly"
		RegWrite this.RuntimeVersion, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32\0.0.0.0", "RuntimeVersion"
		RegWrite SetCodeBase, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\InprocServer32\0.0.0.0", "CodeBase"
		RegWrite this.ClassName, "REG_SZ", this.RootKey . "CLSID\" . this.CLSID . "\ProgId"
		RegWrite ".NET Category", "REG_SZ", this.RootKey . "Component Categories\" . this.Component, "0"
		try
			this.ComInstance := ComObject(this.ClassName)
		catch as Err
			{
			this.Unregister(SetInstallDir)
			return err.message
			}
		return 0
		}
	static Unregister(SetInstallDir := A_Temp)
		{
		if (this.ComInstance is ComObject)
			this.ComInstance := unset
		RegDeleteKey this.RootKey . this.ClassName
		RegDeleteKey this.RootKey . "CLSID\" . this.CLSID
		RegDelete this.RootKey . "Component Categories\" . this.Component, "0"
		if FileExist(SetInstallDir . "\psScript.dll")
			{
			try 
				FileDelete SetInstallDir . "\psScript.dll"
			}
		return 0
		}
	}