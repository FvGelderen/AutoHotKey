;@Ahk2Exe-SetName Intune Win32 Launcher 64-bit
;@Ahk2Exe-SetOrigFilename IntuneLauncher.exe
;@Ahk2Exe-SetDescription Intune Win32 Launcher 1.9
;@Ahk2Exe-SetVersion 1.9.0.0
;@Ahk2Exe-SetCompanyName Provolve B.V.
;@Ahk2Exe-SetCopyright Ferry van Gelderen
;@Ahk2Exe-SetMainIcon Grey.ico
;@Ahk2Exe-AddResource Red.ico, 160
;@Ahk2Exe-AddResource Green.ico, 206
;@Ahk2Exe-AddResource Yellow.ico, 207
;@Ahk2Exe-AddResource Blue.ico, 208

#Requires AutoHotkey v2.0
#SingleInstance off
#NoTrayIcon

ProductName := "Intune Win32 Launcher"
ProductVersion := "1.9.0.0"
SystemPipeName := "Intune Win32 Launcher\SYSTEM"
UserPipeName := "Intune Win32 Launcher\USER"

OSVersion := StrSplit(A_OSVersion, ".")
If Not OSVersion[1] >= 10
	{
	MsgBox "The operating system version is not supported for this product. Microsoft Windows 10 or higher is required.", ProductName " " ProductVersion, "16 T30"
	ExitApp 1150
	}
SplitPath A_ScriptFullPath,,,, &FileName
IniFile := A_ScriptDir "\" FileName ".ini"
TitleIcon := A_ScriptDir "\" FileName ".ico"
SetAltSubmit := False
If FileExist(A_ScriptDir "\" FileName ".png")
	{
	TitleIcon := A_ScriptDir "\" FileName ".png"
	SetAltSubmit := True
	}
A_PID := ProcessExist()
EnvSet "WorkingDir", A_ScriptDir
EnvSet "StartMenuDir", A_StartMenuCommon
EnvSet "ProgramsDir", A_ProgramsCommon
EnvSet "DesktopDir", A_DesktopCommon
EnvSet "StartupDir", A_StartupCommon
EnvSet "ReturnCode", 0
EnvSet "RebootRequired", PendingReboot()
If (PID := ProcessExist("explorer.exe"))
	EnvSet "UserLogin", 1
Else
	EnvSet "UserLogin", 0
ArgumentsNumber := A_Args.Length
If ArgumentsNumber > 0
	{
	If FileExist(IniFile)
		{
		SetEnvSection := "Variables"
		SetEnvCounter := SetEnv(SetEnvSection)
		If SetEnv(A_Language) = 0
			SetEnv("[0409")
		}
	Argument := A_Args[1]
	If InStr(Argument, ":")
		{
		Arg_array := StrSplit(Argument, ":")
		Section := Arg_array[1]
		NotifyCode := Arg_array[2]
		}
	Else
		{
		Section := Argument
		NotifyCode := ""
		}
	If NotifyCode = ""
		{
		If A_IsAdmin = 0
			{
			MsgBox "Administrative privileges are required to execute the system operations.", ProductName " " ProductVersion, "16 T30"
			ExitApp 5
			}
		LogFile := IniRead(IniFile, Section, "LogFile", A_Temp "\" FileName ".log")
		If LogFile != A_Temp "\" FileName ".log"
			{
			LogFile := TransForm(LogFile)
			SplitPath LogFile,, &LogPath
			If !DirExist(LogPath)
				{
				Try
					DirCreate LogPath
				Catch as err
					LogFile := A_Temp "\" FileName ".log"
				}
			}
		FileAppend A_Now "`t[ INFO  ]`t" ProductName " " ProductVersion " started with Process Id " A_PID " for section: " Section "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tComputer name: " A_ComputerName "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tOperating System version: " A_OSVersion " (64-bit=" A_Is64bitOS ")`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tOperating System language code: " A_Language "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tUser name: " A_UserName "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tRebootRequired: " (EnvGet("RebootRequired")) "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tUserLogin: " (EnvGet("UserLogin")) "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tWorking directory: " A_ScriptDir "`r`n" , LogFile
		}
	Else
		{
		If NotifyCode = "*"
			{
			If FileExist(IniFile)
				{
				; --------------------------------------------------------------- USER Notification process ---------------------------------------------------------------
				LogFile := ""
				SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
				DllCall("kernel32.dll\SetProcessShutdownParameters", "UInt", 0x4FF, "UInt", 0)
				OnMessage(0x0011, On_WM_QUERYENDSESSION)
				ActiveProgress := 0
				Loop
					{
					GetMessage := Receive_NamedPipeMessage(SystemPipeName)
					If ActiveProgress = 1
						{
						SetTimer RunWaitProgress, 0
						NotifyGui.Destroy()
						ActiveProgress := 0
						}
					Arg_array := StrSplit(GetMessage, ":")
					Section := Arg_array[1]
					NotifyCode := Arg_array[2]
					SectionType := IniRead(IniFile, Section, "SectionType", "")
					If (SectionType = "Notification" or SectionType = "Progress" or SectionType = "Message" or Section = "*" or Section = "?")
						{
						If SectionType = "Notification"
							{
							RegWrite ProductVersion, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, ProductVersion
							RegWrite A_Now, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, "TimeStamp"
							BackColor := "F0F0F0"
							TextColor := "c"
							TimeColor := "c721C24"
							LinkOptions := "s7 underline c194499"
							AppsUseLightTheme := RegKeyExists("HKCU", "SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
							If AppsUseLightTheme
								{
								AppsUseLightTheme := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
								If AppsUseLightTheme = 0
									{
									BackColor := "2C2C2C"
									TextColor := "cF0F0F0"
									TimeColor := "cE78F79"
									LinkOptions := "s7 underline c6D99EE"
									}
								}
							Else
								AppsUseLightTheme := 1
							ProductNameText := IniRead(IniFile, Section, "ProductNameText", "")
							ProductNameText := TransForm(ProductNameText)
							ProductVersionText := IniRead(IniFile, Section, "ProductVersionText", "")
							ProductVersionText := TransForm(ProductVersionText)
							InformationText := IniRead(IniFile, Section, "InformationText", "")
							InformationText := TransForm(InformationText)
							TimerText := IniRead(IniFile, Section, "TimerText", "")
							TimerText := TransForm(TimerText)
							DeferLinkText := IniRead(IniFile, Section, "DeferLinkText", "Defer")
							DeferLinkText := TransForm(DeferLinkText)
							ContinueLinkText := IniRead(IniFile, Section, "ContinueLinkText", "Continue")
							ContinueLinkText := TransForm(ContinueLinkText)
							Wait := IniRead(IniFile, Section, "Wait", 60)
							HideDeferLink := 0
							DeferTimes := IniRead(IniFile, Section, "DeferTimes", "")
							DeferTimes := Trim(DeferTimes)
							If (DeferTimes = "" or DeferTimes = 0)
								HideDeferLink := 1
							Else
								{
								DeferredCount := RegRead("HKEY_CURRENT_USER\SOFTWARE\" . ProductName . "\DeferTimes", Section . " " . ProductNameText . " " . ProductVersionText, 0)
								If DeferTimes <= DeferredCount
									HideDeferLink := 1					
								}
							NotifyGui := Gui("+DpiScale -Caption +AlwaysOnTop +Owner +E0x08000000", Section)
							Hwnd := NotifyGui.Hwnd
							NotifyGui.MarginX := 10
							NotifyGui.MarginY := 0
							NotifyGui.BackColor := BackColor
							ShowBorder := NotifyGui.Add("Text", "x0 y0 w64 h128 vLineL Background646464")
							If FileExist(TitleIcon)
								{
								If SetAltSubmit = True
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans AltSubmit vIcon", TitleIcon)
								Else
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon", TitleIcon)
								}
							Else
								ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon Icon1", A_ScriptFullPath)
							NotifyGui.SetFont("S12 bold", "Consolas")
							ShowProductName := NotifyGui.Add("Text", "x70 y5 w256 R1 BackgroundTrans Right", ProductNameText)
							ShowProductName.Opt(TextColor)
							NotifyGui.SetFont("S9 norm", "Consolas")
							ShowProductVersion := NotifyGui.Add("Text", "xp yp+20 wp R1 BackgroundTrans Right", ProductVersionText)
							ShowProductVersion.Opt(TextColor)
							NotifyGui.SetFont("S7 norm", "Consolas")
							ShowInfo := NotifyGui.Add("Edit", "xp yp+20 wp h50 Left ReadOnly -E0x200", InformationText)
							ShowInfo.Opt(TextColor)
							If AppsUseLightTheme = 0
								{
								EditHwnd := ShowInfo.Hwnd
								ShowInfo.Opt("Background2C2C2C")
								DllCall("uxtheme\SetWindowTheme", "ptr", EditHwnd, "str", "DarkMode_Explorer", "ptr", 0)	
								}
							ShowTextTime := NotifyGui.Add("Text", "xp y100 wp-35 R1 BackgroundTrans Left", TimerText)
							ShowTextTime.Opt(TextColor)
							ShowTime := NotifyGui.Add("Text", "x+5 yp R1 BackgroundTrans Right", "00:00")
							ShowTime.Opt(TimeColor)
							If HideDeferLink = 0
								{
								DeferLink := NotifyGui.Add("Text", "x70 yp+15 w128", DeferLinkText)
								DeferLink.SetFont(LinkOptions)
								DeferLink.OnEvent("Click", CloseNotificationRetry)
								ContinueLink := NotifyGui.Add("Text", "xp+128 yp w128 Right", ContinueLinkText)
								}
							Else
								ContinueLink := NotifyGui.Add("Text", "x198 yp+15 w128 Right", ContinueLinkText)
							ContinueLink.SetFont(LinkOptions)
							ContinueLink.OnEvent("Click", CloseNotificationContinue)
							Dummy := NotifyGui.Add("Edit", "xp yp w0 h0 ReadOnly", "")
							VirtualScreenHeight := SysGet(79)
							NotifyGui.Show("x0 y" . VirtualScreenHeight)
							WinExist("A")
							Dummy.Focus()
							MonitorPrimary := MonitorGetPrimary()
							MonitorGetWorkArea MonitorPrimary, &mLeft, &mTop, &mRight, &mBottom
							RECT := Buffer(16, 0)
							DllCall("User32.dll\GetWindowRect", "ptr",Hwnd, "ptr",RECT)
							W := NumGet(RECT,  8, "int") - NumGet(RECT, 0, "int")
							H := NumGet(RECT, 12, "int") - NumGet(RECT, 4, "int")
							X := (mRight - W) - 10
							Y := (mBottom - H) - 10
							DllCall("User32.dll\MoveWindow", "ptr",Hwnd, "int",X, "int",Y, "int",W, "int",H, "int",1)
							SoundPlay "*64"
							UserInteraction := 0
							Counter := Wait
							Loop Wait
								{
								Counter--
								TimeLeft := FormatSeconds(Counter)
								ShowTime.Value := TimeLeft
								Sleep 1000
								If (Counter = 0 or UserInteraction = 1)
									Break
								}
							If UserInteraction = 0
								{
								NotifyGui.Destroy()
								Send_NamedPipeMessage("0", UserPipeName)
								}
							}
						If SectionType = "Progress"
							{
							RegWrite ProductVersion, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, ProductVersion
							RegWrite A_Now, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, "TimeStamp"
							BackColor := "F0F0F0"
							TextColor := "c"
							TimeColor := "c721C24"
							LinkOptions := "s7 underline c194499"
							AppsUseLightTheme := RegKeyExists("HKCU", "SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
							If AppsUseLightTheme
								{
								AppsUseLightTheme := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
								If AppsUseLightTheme = 0
									{
									BackColor := "2C2C2C"
									TextColor := "cF0F0F0"
									TimeColor := "cE78F79"
									LinkOptions := "s7 underline c6D99EE"
									}
								}
							Else
								AppsUseLightTheme := 1
							ProductNameText := IniRead(IniFile, Section, "ProductNameText", "")
							ProductNameText := TransForm(ProductNameText)
							ProductVersionText := IniRead(IniFile, Section, "ProductVersionText", "")
							ProductVersionText := TransForm(ProductVersionText)
							InformationText := IniRead(IniFile, Section, "InformationText", "")
							InformationText := TransForm(InformationText)
							NotifyGui := Gui("+DpiScale -Caption +AlwaysOnTop +Owner +E0x08000000", Section)
							Hwnd := NotifyGui.Hwnd
							NotifyGui.MarginX := 10
							NotifyGui.MarginY := 0
							NotifyGui.BackColor := BackColor
							ShowBorder := NotifyGui.Add("Text", "x0 y0 w64 h128 vLineL Background3399FF")
							If FileExist(TitleIcon)
								{
								If SetAltSubmit = True
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans AltSubmit vIcon", TitleIcon)
								Else
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon", TitleIcon)
								}
							Else
								ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon Icon5", A_ScriptFullPath)
							NotifyGui.SetFont("S12 bold", "Consolas")
							ShowProductName := NotifyGui.Add("Text", "x70 y5 w256 R1 BackgroundTrans Right", ProductNameText)
							ShowProductName.Opt(TextColor)
							NotifyGui.SetFont("S9 norm", "Consolas")
							ShowProductVersion := NotifyGui.Add("Text", "xp yp+20 wp R1 BackgroundTrans Right", ProductVersionText)
							ShowProductVersion.Opt(TextColor)
							NotifyGui.SetFont("S7 norm", "Consolas")
							ShowInfo := NotifyGui.Add("Edit", "xp yp+20 wp h50 Left ReadOnly -E0x200", InformationText)
							ShowInfo.Opt(TextColor)
							If AppsUseLightTheme = 0
								{
								EditHwnd := ShowInfo.Hwnd
								ShowInfo.Opt("Background2C2C2C")
								DllCall("uxtheme\SetWindowTheme", "ptr", EditHwnd, "str", "DarkMode_Explorer", "ptr", 0)	
								}
							NotifyProgress := NotifyGui.Add("Progress", "x70 y100 h8 w256 0x8")
							NotifyProgress.Opt("c3399FF")
							If AppsUseLightTheme = 0
								NotifyProgress.Opt("Background2C2C2C")
							Dummy := NotifyGui.Add("Edit", "xp yp w0 h0 ReadOnly", "")
							VirtualScreenHeight := SysGet(79)
							NotifyGui.Show("x0 y" . VirtualScreenHeight)
							WinExist("A")
							Dummy.Focus()
							MonitorPrimary := MonitorGetPrimary()
							MonitorGetWorkArea MonitorPrimary, &mLeft, &mTop, &mRight, &mBottom
							RECT := Buffer(16, 0)
							DllCall("User32.dll\GetWindowRect", "ptr",Hwnd, "ptr",RECT)
							W := NumGet(RECT,  8, "int") - NumGet(RECT, 0, "int")
							H := NumGet(RECT, 12, "int") - NumGet(RECT, 4, "int")
							X := (mRight - W) - 10
							Y := (mBottom - H) - 10
							DllCall("User32.dll\MoveWindow", "ptr",Hwnd, "int",X, "int",Y, "int",W, "int",H, "int",1)
							SetTimer RunWaitProgress, 50
							ActiveProgress := 1
							Send_NamedPipeMessage("0", UserPipeName)
							}
						If SectionType = "Message"
							{
							If NotifyCode = ""
								NotifyCode := 0
							RegWrite ProductVersion, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, ProductVersion
							RegWrite A_Now, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, "TimeStamp"
							BackColor := "F0F0F0"
							TextColor := "c"
							TimeColor := "c721C24"
							LinkOptions := "s7 underline c194499"
							AppsUseLightTheme := RegKeyExists("HKCU", "SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
							If AppsUseLightTheme
								{
								AppsUseLightTheme := RegRead("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize", "AppsUseLightTheme", 1)
								If AppsUseLightTheme = 0
									{
									BackColor := "2C2C2C"
									TextColor := "cF0F0F0"
									TimeColor := "cE78F79"
									LinkOptions := "s7 underline c6D99EE"
									}
								}
							Else
								AppsUseLightTheme := 1
							ProductNameText := IniRead(IniFile, Section, "ProductNameText", "")
							ProductNameText := TransForm(ProductNameText)
							ProductVersionText := IniRead(IniFile, Section, "ProductVersionText", "")
							ProductVersionText := TransForm(ProductVersionText)
							ContinueLinkText := IniRead(IniFile, Section, "ContinueLinkText", "Continue")
							ContinueLinkText := TransForm(ContinueLinkText)
							Status := IniRead(IniFile, Section, "Status", "Info")
							Status := TransForm(Status)
							Wait := IniRead(IniFile, Section, "Wait", 30)
							ShowWarning := 1
							ReturnCodes := IniRead(IniFile, Section, "ReturnCodes", "0, 1707")
							For index, ReturnCode in StrSplit(ReturnCodes, ",")
								{
								ReturnCode := Trim(ReturnCode)
								If NotifyCode = ReturnCode
									ShowWarning := 0
								}
							If ShowWarning = 0
								{
								UserScript := IniRead(IniFile, Section, "UserScript", "")
								UserScript := Trim(UserScript)
								InformationText := IniRead(IniFile, Section, "InformationText", "")
								InformationText := TransForm(InformationText)
								If (Status = "Warning" or Status = "Error")
									{
									If Status = "Warning"
										{
										ShowIconNumber := "Icon4"
										BarColor := "BackgroundEBB800"
										ProgressColor := "cEBB800"
										}
									If Status = "Error"
										{
										ShowIconNumber := "Icon2"
										BarColor := "BackgroundE40000"
										ProgressColor := "cE40000"
										}
									}
								Else
									{
									ShowIconNumber := "Icon3"
									BarColor := "Background429300"
									ProgressColor := "c429300"
									}
								}
							Else
								{
								UserScript := ""
								InformationText := GetSysErrorText(NotifyCode)
								ShowIconNumber := "Icon4"
								BarColor := "BackgroundEBB800"
								ProgressColor := "cEBB800"
								}
							NotifyGui := Gui("+DpiScale -Caption +AlwaysOnTop +Owner +E0x08000000", Section)
							Hwnd := NotifyGui.Hwnd
							NotifyGui.MarginX := 10
							NotifyGui.MarginY := 0
							NotifyGui.BackColor := BackColor
							ShowBorder := NotifyGui.Add("Text", "x0 y0 w64 h128 vLineL " . BarColor)
							If FileExist(TitleIcon)
								{
								If SetAltSubmit = True
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans AltSubmit", TitleIcon)
								Else
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans", TitleIcon)
								}
							Else
								ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon " . ShowIconNumber, A_ScriptFullPath)
							NotifyGui.SetFont("S12 bold", "Consolas")
							ShowProductName := NotifyGui.Add("Text", "x70 y5 w256 R1 BackgroundTrans Right", ProductNameText)
							ShowProductName.Opt(TextColor)
							NotifyGui.SetFont("S9 norm", "Consolas")
							ShowProductVersion := NotifyGui.Add("Text", "xp yp+20 wp R1 BackgroundTrans Right", ProductVersionText)
							ShowProductVersion.Opt(TextColor)
							NotifyGui.SetFont("S7 norm", "Consolas")
							ShowInfo := NotifyGui.Add("Edit", "xp yp+20 wp h50 Left ReadOnly -E0x200", "")
							ShowInfo.Opt(TextColor)
							If AppsUseLightTheme = 0
								{
								EditHwnd := ShowInfo.Hwnd
								ShowInfo.Opt("Background2C2C2C")
								DllCall("uxtheme\SetWindowTheme", "ptr", EditHwnd, "str", "DarkMode_Explorer", "ptr", 0)	
								}
							NotifyProgress := NotifyGui.Add("Progress", "x70 y100 h8 w256 0x8")
							NotifyProgress.Opt(ProgressColor)
							If AppsUseLightTheme = 0
								NotifyProgress.Opt("Background2C2C2C")
							Dummy := NotifyGui.Add("Edit", "xp yp w0 h0 ReadOnly", "")
							VirtualScreenHeight := SysGet(79)
							NotifyGui.Show("x0 y" . VirtualScreenHeight)
							WinExist("A")
							Dummy.Focus()
							MonitorPrimary := MonitorGetPrimary()
							MonitorGetWorkArea MonitorPrimary, &mLeft, &mTop, &mRight, &mBottom
							RECT := Buffer(16, 0)
							DllCall("User32.dll\GetWindowRect", "ptr",Hwnd, "ptr",RECT)
							W := NumGet(RECT,  8, "int") - NumGet(RECT, 0, "int")
							H := NumGet(RECT, 12, "int") - NumGet(RECT, 4, "int")
							X := (mRight - W) - 10
							Y := (mBottom - H) - 10
							DllCall("User32.dll\MoveWindow", "ptr",Hwnd, "int",X, "int",Y, "int",W, "int",H, "int",1)
							If UserScript
								{
								EnvSet "WorkingDir", A_ScriptDir
								EnvSet "StartMenuDir", A_StartMenu
								EnvSet "ProgramsDir", A_Programs
								EnvSet "DesktopDir", A_Desktop
								EnvSet "StartupDir", A_Startup
								EnvSet "ReturnCode", 0
								EnvSet "RebootRequired", PendingReboot()
								ScriptSections := ""
								UserScriptCounter := Script(UserScript)
								}
							NotifyProgress.Opt("-0x8")
							NotifyProgress.Value := 100
							ShowInfo.Value := InformationText
							ContinueLink := NotifyGui.Add("Text", "xp+128 yp+15 w128 Right", ContinueLinkText)
							ContinueLink.SetFont(LinkOptions)
							ContinueLink.OnEvent("Click", CloseNotificationContinue)
							If (Status = "Warning" or Status = "Error")
								{
								If Status = "Warning"
									SoundPlay "*48"
								If Status = "Error"
									SoundPlay "*16"
								}
							Else
								SoundPlay "*64"
							UserInteraction := 0
							Loop Wait
								{
								Sleep 1000
								If UserInteraction = 1
									Break
								}
							If UserInteraction = 0
								{
								NotifyGui.Destroy()
								Send_NamedPipeMessage("0", UserPipeName)
								}
							}
						If Section = "*"
							{
							Send_NamedPipeMessage("0", UserPipeName)
							ExitApp 0
							}
						If Section = "?"
							{
							Send_NamedPipeMessage(A_UserName, UserPipeName)
							System_PID := NotifyCode
							SetTimer SystemProcessCheck, 1000
							}
						}
					Else
						Send_NamedPipeMessage("Error", UserPipeName)
					}
				}
			Else
				ExitApp 1
			}
		}
	If FileExist(IniFile)
		{
		If FileExist(LogFile)
			FileAppend A_Now "`t[ INFO  ]`tINI file: " IniFile "`r`n" , LogFile
		SectionType := IniRead(IniFile, Section, "SectionType", "")
		SectionType := Trim(SectionType)
		If (SectionType = "Execution")
			{
			; --------------------------------------------------------------- SYSTEM Execution process ---------------------------------------------------------------
			FoundSessions := CheckForOtherSessions(0)
			If FoundSessions > 0
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ WARN  ]`t" ProductName " is already started for another session. This process will stop with return code: 1618 (Another Installation is already in progress)`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
				ExitApp 1618
				}
			Result := PsScriptManager.Register()
			If !Result = 0
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ ERROR ]`tPowershell COM registration error. psScript.dll failed to register on this system. Error: " Result  "`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
				ExitApp 1603
				}
			UserScript := ""
			SetTimer UserLoginPendingReboot, 250
			ActivePendingReboot := PendingReboot()
			If ActivePendingReboot
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ WARN  ]`tPending reboot has been detected.`r`n" , LogFile
				}
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Performing Executing section: " Section " ]`r`n" , LogFile
			FoundProcess := 0
			CheckProcess := IniRead(IniFile, Section, "CheckProcess", "")
			If CheckProcess
				{
				CheckProcess := TransForm(CheckProcess)
				For index, Process in StrSplit(CheckProcess, ",")
					{
					Process := Trim(Process)
					If FileExist(LogFile)
						FileAppend A_Now "`t[ INFO  ]`tCheckProcess: " Process "`r`n" , LogFile
					If (PID := ProcessExist(Process))
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`t" Process " exists with Process Id: " PID "`r`n" , LogFile
						FoundProcess++
						}
					}
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tContinuing without notification...`r`n" , LogFile
				}
			EnvSet "FoundProcess", FoundProcess
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`tFoundProcess: " (EnvGet("FoundProcess")) "`r`n" , LogFile
			cpauPID := 0
			If FoundProcess > 0
				{
				If (PID := ProcessExist("explorer.exe"))
					{
					If FileExist(LogFile)
						FileAppend A_Now "`t[ INFO  ]`t[ Activating Notification Process ]`r`n" , LogFile
					If FileExist(A_ScriptFullPath)
						SetSecurityError := SetSecurityFile(A_ScriptFullPath, "S-1-1-0", "1179817", "1")
					If FileExist(IniFile)
						SetSecurityError := SetSecurityFile(IniFile, "S-1-1-0", "1179817", "1")
					If FileExist(TitleIcon)
						SetSecurityError := SetSecurityFile(TitleIcon, "S-1-1-0", "1179817", "1")
					If FileExist(A_Temp . "\psScript.dll")
						SetSecurityError := SetSecurityFile(A_Temp . "\psScript.dll", "S-1-1-0", "1179817", "1")
					cpauPID := CreateProcessAsUser(A_ScriptName, "*:*", A_ScriptDir, "explorer.exe", 0)
					If (cpauPID = -1 or cpauPID = -2)
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ ERROR ]`tCreateProcessAsUser function stopped with error: " cpauPID "`r`n" , LogFile
						}
					Else
						{
						Send_NamedPipeMessage("?:" . A_PID, SystemPipeName)
						GetUserName := Receive_NamedPipeMessage(UserPipeName)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tCreateProcessAsUser function is started for " GetUserName " with Process Id: " cpauPID "`r`n" , LogFile
						EnvSet "UserLoginName", GetUserName
						}
					}
				If ActivePendingReboot
					{
					NotifyPendingReboot := IniRead(IniFile, Section, "NotifyPendingReboot", "")
					If NotifyPendingReboot
						{
						If (cpauPID := ProcessExist(cpauPID))
							{
							NotifyPendingReboot := Trim(NotifyPendingReboot)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`t[ Performing Notification section: " NotifyPendingReboot " ]`r`n" , LogFile
							Send_NamedPipeMessage(NotifyPendingReboot . ":0", SystemPipeName)
							ReturnCode := Receive_NamedPipeMessage(UserPipeName)
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
								FileAppend A_Now "`t[ WARN  ]`t" ProductName " is cancelled successfully due to a pending reboot detection. (Return Code = 350)`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
								}
							Send_NamedPipeMessage("*:0", SystemPipeName)
							ReturnCode := Receive_NamedPipeMessage(UserPipeName)
							If (cpauPID := ProcessExist(cpauPID))
								ProcessClose cpauPID
							PsScriptManager.Unregister()
							ExitApp 350
							}
						}
					}
				NotifyStart := IniRead(IniFile, Section, "NotifyStart", "")
				If NotifyStart
					{
					If (cpauPID := ProcessExist(cpauPID))
						{
						NotifyStart := Trim(NotifyStart)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`t[ Performing Notification section: " NotifyStart " ]`r`n" , LogFile
						Send_NamedPipeMessage(NotifyStart . ":0", SystemPipeName)
						ReturnCode := Receive_NamedPipeMessage(UserPipeName)
						If (ReturnCode = 1602 or ReturnCode = 1618)
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`t" ProductName " is cancelled successfully. (Notification process returned: " ReturnCode ")`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
							Send_NamedPipeMessage("*:0", SystemPipeName)
							ReturnCode := Receive_NamedPipeMessage(UserPipeName)
							If (cpauPID := ProcessExist(cpauPID))
								ProcessClose cpauPID
							PsScriptManager.Unregister()
							ExitApp 1602
							}
						Else
							{
							If ReturnCode = 0
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									{
									FileAppend A_Now "`t[ WARN  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
									FileAppend A_Now "`t[ INFO  ]`t" ProductName " is cancelled successfully. (Return Code = 1602)`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
									}
								If !ReturnCode = "ExitApp"
									{
									Send_NamedPipeMessage("*:0", SystemPipeName)
									ReturnCode := Receive_NamedPipeMessage(UserPipeName)
									}
								If (cpauPID := ProcessExist(cpauPID))
									ProcessClose cpauPID
								PsScriptManager.Unregister()
								ExitApp 1602
								}
							}
						}
					}
				NotifyProgress := IniRead(IniFile, Section, "NotifyProgress", "")
				If NotifyProgress
					{
					If (cpauPID := ProcessExist(cpauPID))
						{
						NotifyProgress := Trim(NotifyProgress)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`t[ Performing Notification section: " NotifyProgress " ]`r`n" , LogFile
						Send_NamedPipeMessage(NotifyProgress . ":0", SystemPipeName)
						ReturnCode := Receive_NamedPipeMessage(UserPipeName)
						If ReturnCode = 0
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
							}
						}
					}
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`t" CheckProcess " not running. Continuing without notification...`r`n" , LogFile
				}
			PreLaunchSection := IniRead(IniFile, Section, "PreLaunchScript", "")
			If PreLaunchSection
				{
				PreLaunchSection := Trim(PreLaunchSection)
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`t[ Performing PreLaunch script section: " PreLaunchSection " ]`r`n" , LogFile
				ScriptSections := ""
				PreLaunchCounter := Script(PreLaunchSection)
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tContinuing without PreLaunch script section...`r`n" , LogFile
				}
			ExitCode := 0
			TaskKill := IniRead(IniFile, Section, "TaskKill", "")
			If TaskKill
				{
				TaskKill := TransForm(TaskKill)
				For index, StopProcess in StrSplit(TaskKill, ",")
					{
					StopProcess := Trim(StopProcess)
					Loop
						{
						If (PID := ProcessExist(StopProcess))
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tStop process: " StopProcess " with Process Id: " PID "`r`n" , LogFile
							ProcessClose PID
							Sleep 1000
							}
						Else
							Break
						}
					}
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tContinuing without TaskKill...`r`n" , LogFile
				}
			WorkingDir := IniRead(IniFile, Section, "WorkingDir", "")
			If WorkingDir
				{
				WorkingDir := TransForm(WorkingDir)
				WorkingDir := A_ScriptDir "\" WorkingDir
				}
			Else
				WorkingDir := A_ScriptDir
			SetSourceDir := IniRead(IniFile, Section, "SetSourceDir", "")
			If SetSourceDir
				{
				SetSourceDir := TransForm(SetSourceDir)
				DeleteSourceDir := IniRead(IniFile, Section, "DeleteSourceDir", "")
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tCreate SourceDir folder: " SetSourceDir " and copy content from: " WorkingDir "`r`n" , LogFile
				If !DirExist(SetSourceDir)
					DirCreate SetSourceDir
				Try
					DirCopy WorkingDir, SetSourceDir, 1
				Catch as err
					{
					If FileExist(LogFile)
						FileAppend A_Now "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`r`n" , LogFile
					}
				Else
					{
					If FileExist(LogFile)
						FileAppend A_Now "`t[ INFO  ]`tCreate SourceDir and copy content successfully executed.`r`n" , LogFile
					}
				}
			Else
				{
				DeleteSourceDir := ""
				SetSourceDir := WorkingDir
				}
			Options := IniRead(IniFile, Section, "Options", "")
			Options := TransForm(Options)
			Loop
				{
				Command := IniRead(IniFile, Section, "Command" A_Index, "")
				If Command = ""
					Break
				Else
					{
					Command := TransForm(Command)
					If (SubStr(Command, 1, 1) = "|" OR SubStr(Command, 1, 1) = ">")
						{
						If SubStr(Command, 1, 1) = "|"
							{
							MSIParam := StrSplit(Command, "|")
							If SubStr(SetSourceDir, StrLen(SetSourceDir), StrLen(SetSourceDir)) = "\"
								SetSourceDir := RTrim(SetSourceDir, "\")
							MSIPackagePath := SetSourceDir . "\" . MSIParam[2]
							MSICommandLine := MSIParam[3]
							MSILogFile := MSIParam[4]
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tUsing MSI PackagePath: " MSIPackagePath "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tUsing MSI CommandLine: " MSICommandLine "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tUsing MSI LogFile: " MSILogFile "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tStarting MsiInstallProduct...`r`n" , LogFile
								}
							Try
								ExitCode := MsiInstallProduct(MSIPackagePath, MSICommandLine, MSILogFile)
							Catch as err
								{
								ExitCode :=	1603
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tError: " err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tMsiInstallProduct successfully stopped with return code: " ExitCode "`r`n" , LogFile
								}
							}
						If SubStr(Command, 1, 1) = ">"
							{
							Command := LTrim(Command, ">")
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tUsing working directory: " SetSourceDir "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tUsing Powershell CommandLine: " Command "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tStarting PS_Script...`r`n" , LogFile
								}
							SplitPath Command,, &PSDir, &Extension
							If Extension = "ps1"
								{
								Try
									Command := FileRead(Command)
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , LogFile
									}
								Else
									If PSDir
										Command := "Set-Location -Path `"" PSDir "`"`r`n" Command
								}
							Else
								{
								Command := StrReplace(Command, A_Space . "\n" . A_Space, "`r`n")
								Command := "Set-Location -Path `"" SetSourceDir "`"`r`n" Command
								}
							Command := Format(Command)
							try 
								psReturn := PsScriptManager.ComInstance.PS_Script(Command)
							Catch as err
								{
								ExitCode :=	1603
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tPowershell command stopped with error:`r`n" err.Message "`r`n" , LogFile
								}
							Else
								{
								ExitCode := 0
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tPowershell command successfully stopped (return code: 0) with return message:`r`n" psReturn "`r`n" , LogFile
								}
							}
						}
					Else
						{
						If FileExist(LogFile)
							{
							FileAppend A_Now "`t[ INFO  ]`tUsing working directory: " SetSourceDir "`r`n" , LogFile
							FileAppend A_Now "`t[ INFO  ]`tUsing options: " Options "`r`n" , LogFile
							FileAppend A_Now "`t[ INFO  ]`tStarting command: " Command "`r`n" , LogFile
							}
						Try
							ExitCode := RunWait(Command, SetSourceDir, Options, &PID)
						Catch as err
							{
							ExitCode :=	1603
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`tError: " err.Message "`r`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCommand with Process Id '" PID "' successfully stopped with return code: " ExitCode "`r`n" , LogFile
							}
						}
					EnvSet "ReturnCode", ExitCode
					}
				}
			PostExitSection := IniRead(IniFile, Section, "PostExitScript", "")
			If PostExitSection
				{
				PostExitSection := Trim(PostExitSection)
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`t[ Performing PostExit script section: " PostExitSection " ]`r`n" , LogFile
				ScriptSections := ""
				PostExitCounter := Script(PostExitSection)
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tContinuing without PostExit script section...`r`n" , LogFile
				}
			If DeleteSourceDir = 1
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tDelete SourceDir folder: " SetSourceDir "`r`n" , LogFile
				Try
					DirDelete SetSourceDir, 1
				Catch as err
					{
					If FileExist(LogFile)
						FileAppend A_Now "`t[ WARN  ]`tDelete SourceDir folder failed with Error: " err.Message "`r`n" , LogFile
					}
				Else
					{
					If FileExist(LogFile)
						FileAppend A_Now "`t[ INFO  ]`tDelete SourceDir folder successfully executed.`r`n" , LogFile
					}
				}
			If (cpauPID := ProcessExist(cpauPID))
				{
				NotifyFinish := IniRead(IniFile, Section, "NotifyFinish", "")
				If NotifyFinish
					{
					NotifyFinish := Trim(NotifyFinish)
					If FileExist(LogFile)
						FileAppend A_Now "`t[ INFO  ]`t[ Performing Notification section: " NotifyFinish " ]`r`n" , LogFile
					Send_NamedPipeMessage(NotifyFinish . ":" . ExitCode, SystemPipeName)
					ReturnCode := Receive_NamedPipeMessage(UserPipeName)
					If ReturnCode = 0
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
						Send_NamedPipeMessage("*:0", SystemPipeName)
						ReturnCode := Receive_NamedPipeMessage(UserPipeName)
						}
					Else
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ ERROR ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
						}
					}
				}
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t" ProductName " stopped successfully. (Return Code = " ExitCode ")`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
			If (cpauPID := ProcessExist(cpauPID))
				ProcessClose cpauPID
			PsScriptManager.Unregister()
			ExitApp ExitCode
			}
		Else
			{
			If FileExist(LogFile)
				FileAppend A_Now "`t[ ERROR ]`tExecution SectionType for section '" Section "' not found. This process will stop with return code: 1 (Incorrect function.)`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
			ExitApp 1
			}
		}
	Else
		{
		MsgBox "No INI file found. Please make sure the following INI file exist:`n`n" IniFile, ProductName " " ProductVersion, "48 T30"
		ExitApp 1
		}
	}
Else
	{
	If FileExist(IniFile)
		{
		MsgBox "No starting section name. Please make sure a starting section name is added as a parameter for this executable that will be used for the execution section in the following INI file:`n`n" IniFile, ProductName " " ProductVersion, "48 T30"
		ExitApp 1
		}
	Else
		{
		; --------------------------------------------------------------- Create INI File ---------------------------------------------------------------
		SetupFilePath := FileSelect(3, A_ScriptDir, "Select a Windows Installer or Package file (exe, msi, mst, appv, msix, appx)", "(*.exe; *.msi; *.mst; *.appv; *.msix; *.appx")
		If SetupFilePath
			{
			ReservedCharacters := "<>:;`"/\|?*,|"
			SplitPath SetupFilePath, &SetupFileName, &Path, &Extension, &SetupFileNameNoExt
			WorkingDir := StrReplace(Path, A_ScriptDir, "")
			WorkingDir := LTrim(WorkingDir, "\")
			If Extension = "mst"
				TransformFileName := SetupFileName
			Else
				TransformFileName := ""
			If TransformFileName
				{
				SetupFilePath := FileSelect(3, A_ScriptDir . "\" . WorkingDir, "Select the Windows Installer file (*.msi) for " TransformFileName, "(*.msi)")
				If SetupFilePath
					SplitPath SetupFilePath, &SetupFileName, &Path, &Extension
				}
			FileAppend "───────────────────────────────┤ " . ProductName . " " . ProductVersion . " Configuration file ├───────────────────────────────`r`n", IniFile, "UTF-16"
			FileAppend "`r`n", IniFile
			If Extension = "exe"
				{
				GetManufacturer := GetFileVersionInfo(SetupFilePath, "CompanyName")
				GetProductName := GetFileVersionInfo(SetupFilePath, "ProductName")
				GetProductVersion := GetFileVersionInfo(SetupFilePath, "ProductVersion")
				GetFileDescription := GetFileVersionInfo(SetupFilePath, "FileDescription")
				LogFileName := GetManufacturer . "_" . GetProductName . "_" . GetProductVersion
				loop parse ReservedCharacters
					If InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")
				IniWrite A_UserName, IniFile, "Variables", "ScriptAuthor"
				IniWrite FormatTime(, "ShortDate"), IniFile, "Variables", "ScriptDate"
				IniWrite GetManufacturer, IniFile, "Variables", "Manufacturer"
				IniWrite GetProductName, IniFile, "Variables", "ProductName"
				IniWrite GetProductVersion, IniFile, "Variables", "ProductVersion"
				IniWrite GetFileDescription, IniFile, "Variables", "FileDescription"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", IniFile, "Variables", "LogFolder"
				IniWrite LogFileName, IniFile, "Variables", "LogFileName"
				IniWrite SetupFileName, IniFile, "Variables", "SetupFileName"
				FileAppend "`r`n", IniFile
				IniWrite "Pending reboot detected. Please restart your system first to continue this installation.", IniFile, "0409", "InformationTextNotifyPendingReboot"
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", IniFile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed. Save your work and close this application when it is active.", IniFile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", IniFile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", IniFile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being installed. Do not turn off your computer during the installation. Please wait...", IniFile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", IniFile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", IniFile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", IniFile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", IniFile, "0409", "DeferLinkText"
				IniWrite "Continue", IniFile, "0409", "ContinueLinkText"
				IniWrite "Close", IniFile, "0409", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "In afwachting van opnieuw opstarten. Graag zelf uw computer opnieuw opstarten zodat deze installatie doorgang krijgt.", IniFile, "0413", "InformationTextNotifyPendingReboot"
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", IniFile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", IniFile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt geïnstalleerd. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", IniFile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", IniFile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", IniFile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", IniFile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", IniFile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", IniFile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", IniFile, "0413", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", IniFile, "Install", "LogFile"
				IniWrite WorkingDir, IniFile, "Install", "WorkingDir"
				IniWrite "", IniFile, "Install", "SetSourceDir"
				IniWrite "", IniFile, "Install", "DeleteSourceDir"
				IniWrite "", IniFile, "Install", "TaskKill"
				IniWrite "%SetupFileName% /install /quiet /log `"%LogFolder%\%LogFileName%_EXE_Install.log`"", IniFile, "Install", "Command1"
				IniWrite "explorer.exe", IniFile, "Install", "CheckProcess"
				IniWrite "NotifyPendingReboot", IniFile, "Install", "NotifyPendingReboot"
				IniWrite "NotifyStartInstall", IniFile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", IniFile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", IniFile, "Install", "NotifyFinish"
				IniWrite "", IniFile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", IniFile, "Install", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_Install]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install,, Register installation when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_Install]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %Manufacturer% %ProductName% %ProductVersion%, 1`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyPendingReboot", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyPendingReboot", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyPendingReboot", "ProductVersionText"
				IniWrite "%InformationTextNotifyPendingReboot%", IniFile, "NotifyPendingReboot", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyPendingReboot", "ContinueLinkText"
				IniWrite "Error", IniFile, "NotifyPendingReboot", "Status"
				IniWrite "", IniFile, "NotifyPendingReboot", "UserScript"
				IniWrite "300", IniFile, "NotifyPendingReboot", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyStart%", IniFile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", IniFile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyProgress%", IniFile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%ProductName%", IniFile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", IniFile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishInstall", "UserScript"
				IniWrite "60", IniFile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", IniFile, "UnInstall", "LogFile"
				IniWrite WorkingDir, IniFile, "UnInstall", "WorkingDir"
				IniWrite "", IniFile, "UnInstall", "SetSourceDir"
				IniWrite "", IniFile, "UnInstall", "DeleteSourceDir"
				IniWrite "", IniFile, "UnInstall", "TaskKill"
				IniWrite "%SetupFileName% /uninstall /quiet /log `"%LogFolder%\%LogFileName%_EXE_UnInstall.log`"", IniFile, "UnInstall", "Command1"
				IniWrite "explorer.exe", IniFile, "UnInstall", "CheckProcess"
				IniWrite "NotifyStartUnInstall", IniFile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", IniFile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", IniFile, "UnInstall", "NotifyFinish"
				IniWrite "", IniFile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", IniFile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_UnInstall]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall,, Register removal when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_UnInstall]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %Manufacturer% %ProductName% %ProductVersion%, 0`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyStart%", IniFile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", IniFile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyProgress%", IniFile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%ProductName%", IniFile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", IniFile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "60", IniFile, "NotifyFinishUnInstall", "Wait"
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected setup executable File:`n`n" IniFile "`n`nPlease check the install and uninstall commandline parameters.", ProductName " " ProductVersion, "64 T30"
				}
			If Extension = "msi"
				{
				GetManufacturer := WindowsInstaller(SetupFilePath, "Manufacturer")
				GetProductName := WindowsInstaller(SetupFilePath, "ProductName")
				GetProductVersion := WindowsInstaller(SetupFilePath, "ProductVersion")
				GetProductCode := WindowsInstaller(SetupFilePath, "ProductCode")
				LogFileName := GetManufacturer . "_" . GetProductName . "_" . GetProductVersion
				loop parse ReservedCharacters
					If InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")
				IniWrite A_UserName, IniFile, "Variables", "ScriptAuthor"
				IniWrite FormatTime(, "ShortDate"), IniFile, "Variables", "ScriptDate"
				IniWrite GetManufacturer, IniFile, "Variables", "Manufacturer"
				IniWrite GetProductName, IniFile, "Variables", "ProductName"
				IniWrite GetProductVersion, IniFile, "Variables", "ProductVersion"
				IniWrite GetProductCode, IniFile, "Variables", "ProductCode"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", IniFile, "Variables", "LogFolder"
				IniWrite LogFileName, IniFile, "Variables", "LogFileName"
				IniWrite SetupFileName, IniFile, "Variables", "SetupFileName"
				If TransformFileName
					IniWrite TransformFileName, IniFile, "Variables", "TransformFileName"
				FileAppend "`r`n", IniFile
				IniWrite "Pending reboot detected. Please restart your system first to continue this installation.", IniFile, "0409", "InformationTextNotifyPendingReboot"
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", IniFile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed. Save your work and close this application when it is active.", IniFile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", IniFile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", IniFile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being installed. Do not turn off your computer during the installation. Please wait...", IniFile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", IniFile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", IniFile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", IniFile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", IniFile, "0409", "DeferLinkText"
				IniWrite "Continue", IniFile, "0409", "ContinueLinkText"
				IniWrite "Close", IniFile, "0409", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "In afwachting van opnieuw opstarten. Graag zelf uw computer opnieuw opstarten zodat deze installatie doorgang krijgt.", IniFile, "0413", "InformationTextNotifyPendingReboot"
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", IniFile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", IniFile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt geïnstalleerd. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", IniFile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", IniFile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", IniFile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", IniFile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", IniFile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", IniFile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", IniFile, "0413", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", IniFile, "Install", "LogFile"
				IniWrite WorkingDir, IniFile, "Install", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%ProductCode%", IniFile, "Install", "SetSourceDir"
				IniWrite "0", IniFile, "Install", "DeleteSourceDir"
				IniWrite "", IniFile, "Install", "TaskKill"
				If TransformFileName
					IniWrite "|%SetupFileName%|TRANSFORMS=`"%TransformFileName%`" TRANSFORMSSECURE=`"1`" REBOOT=`"ReallySuppress`"|%LogFolder%\%LogFileName%_MSI_Install.log", IniFile, "Install", "Command1"
				Else
					IniWrite "|%SetupFileName%|REBOOT=`"ReallySuppress`"|%LogFolder%\%LogFileName%_MSI_Install.log", IniFile, "Install", "Command1"
				IniWrite "explorer.exe", IniFile, "Install", "CheckProcess"
				IniWrite "NotifyPendingReboot", IniFile, "Install", "NotifyPendingReboot"
				IniWrite "NotifyStartInstall", IniFile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", IniFile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", IniFile, "Install", "NotifyFinish"
				IniWrite "", IniFile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", IniFile, "Install", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_Install]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install,, Register installation when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_Install]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %Manufacturer% %ProductName% %ProductVersion%, 1`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyPendingReboot", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyPendingReboot", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyPendingReboot", "ProductVersionText"
				IniWrite "%InformationTextNotifyPendingReboot%", IniFile, "NotifyPendingReboot", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyPendingReboot", "ContinueLinkText"
				IniWrite "Error", IniFile, "NotifyPendingReboot", "Status"
				IniWrite "", IniFile, "NotifyPendingReboot", "UserScript"
				IniWrite "300", IniFile, "NotifyPendingReboot", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyStart%", IniFile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", IniFile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyProgress%", IniFile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%ProductName%", IniFile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", IniFile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishInstall", "UserScript"
				IniWrite "60", IniFile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", IniFile, "UnInstall", "LogFile"
				IniWrite WorkingDir, IniFile, "UnInstall", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%ProductCode%", IniFile, "UnInstall", "SetSourceDir"
				IniWrite "1", IniFile, "UnInstall", "DeleteSourceDir"				
				IniWrite "", IniFile, "UnInstall", "TaskKill"
				IniWrite "|%SetupFileName%|REMOVE=`"ALL`" REBOOT=`"ReallySuppress`"|%LogFolder%\%LogFileName%_MSI_UnInstall.log", IniFile, "UnInstall", "Command1"
				IniWrite "explorer.exe", IniFile, "UnInstall", "CheckProcess"
				IniWrite "NotifyStartUnInstall", IniFile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", IniFile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", IniFile, "UnInstall", "NotifyFinish"
				IniWrite "", IniFile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", IniFile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_UnInstall]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall,, Register removal when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_UnInstall]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %Manufacturer% %ProductName% %ProductVersion%, 0`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyStart%", IniFile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", IniFile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%ProductName%", IniFile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyProgress%", IniFile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%ProductName%", IniFile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", IniFile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "60", IniFile, "NotifyFinishUnInstall", "Wait"
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected Windows Installer File(s):`n`n" IniFile, ProductName " " ProductVersion, "64 T30"
				}
			If Extension = "appv"
				{
				GetPackageId := ""
				GetVersionId := ""
				GetProductVersion := ""
				GetDisplayName := ""
				GetPackageDescription := ""
				AppxManifestXml := Read_AppxManifest(SetupFilePath)
				Loop parse, AppxManifestXml, "`n", "`r"
					{
					If InStr(A_LoopField, "<Identity")
						{
						Identity := Trim(A_LoopField)
						Identity := StrReplace(Identity, "<Identity")
						Identity := StrReplace(Identity, "/>")
						Identity := Trim(Identity)
						Identity := StrReplace(Identity, "Name=", "`nName=")
						Identity := StrReplace(Identity, "Publisher=", "`nPublisher=")
						Identity := StrReplace(Identity, "Version=", "`nVersion=")
						Identity := StrReplace(Identity, "appv:PackageId=", "`nappv:PackageId=")
						Identity := StrReplace(Identity, "appv:VersionId=", "`nappv:VersionId=")
						Loop parse, Identity, "`n"
							{
							If InStr(A_LoopField, "Version=")
								{
								GetProductVersion := StrReplace(A_LoopField , "Version=")
								GetProductVersion := StrReplace(GetProductVersion, "`"")
								}
							If InStr(A_LoopField, "appv:PackageId=")
								{
								GetPackageId := StrReplace(A_LoopField , "appv:PackageId=")
								GetPackageId := StrReplace(GetPackageId, "`"")
								}
							If InStr(A_LoopField, "appv:VersionId=")
								{
								GetVersionId := StrReplace(A_LoopField , "appv:VersionId=")
								GetVersionId := StrReplace(GetVersionId, "`"")
								}
							}
						}
					If InStr(A_LoopField, "<DisplayName>")
						{
						GetDisplayName := Trim(A_LoopField)
						GetDisplayName := StrReplace(GetDisplayName , "<DisplayName>")
						GetDisplayName := StrReplace(GetDisplayName , "</DisplayName>")
						}
					If InStr(A_LoopField, "<appv:AppVPackageDescription>")
						{
						GetPackageDescription := Trim(A_LoopField)
						GetPackageDescription := StrReplace(GetPackageDescription , "<appv:AppVPackageDescription>")
						GetPackageDescription := StrReplace(GetPackageDescription , "</appv:AppVPackageDescription>")
						}
					}
				AppxManifestXml := ""
				LogFileName := GetDisplayName
				loop parse ReservedCharacters
					If InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")				
				IniWrite A_UserName, IniFile, "Variables", "ScriptAuthor"
				IniWrite FormatTime(, "ShortDate"), IniFile, "Variables", "ScriptDate"
				IniWrite GetDisplayName, IniFile, "Variables", "DisplayName"
				IniWrite GetProductVersion, IniFile, "Variables", "ProductVersion"
				IniWrite GetPackageDescription, IniFile, "Variables", "PackageDescription"
				IniWrite GetPackageId, IniFile, "Variables", "PackageId"
				IniWrite GetVersionId, IniFile, "Variables", "VersionId"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", IniFile, "Variables", "LogFolder"
				IniWrite LogFileName, IniFile, "Variables", "LogFileName"
				IniWrite SetupFileName, IniFile, "Variables", "SetupFileName"
				FileAppend "`r`n", IniFile
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", IniFile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed from your system. Save your work and close this application when it is active.", IniFile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", IniFile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", IniFile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being added to your system. Do not turn off your computer during the installation. Please wait...", IniFile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", IniFile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", IniFile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", IniFile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", IniFile, "0409", "DeferLinkText"
				IniWrite "Continue", IniFile, "0409", "ContinueLinkText"
				IniWrite "Close", IniFile, "0409", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", IniFile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", IniFile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt toegevoegd aan het systeem. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", IniFile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", IniFile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", IniFile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", IniFile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", IniFile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", IniFile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", IniFile, "0413", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", IniFile, "Install", "LogFile"
				IniWrite WorkingDir, IniFile, "Install", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageId%", IniFile, "Install", "SetSourceDir"
				IniWrite "0", IniFile, "Install", "DeleteSourceDir"
				IniWrite "", IniFile, "Install", "TaskKill"
				IniWrite ">Add-AppVClientPackage -path '%SetupFileName%' \n Publish-AppVClientPackage -Global -PackageId %PackageId% -VersionId %VersionId%", IniFile, "Install", "Command1"
				IniWrite "explorer.exe", IniFile, "Install", "CheckProcess"
				IniWrite "", IniFile, "Install", "NotifyPendingReboot"
				IniWrite "", IniFile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", IniFile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", IniFile, "Install", "NotifyFinish"
				IniWrite "Check_AppVStatus", IniFile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", IniFile, "Install", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_AppVStatus]`r`n", IniFile
				FileAppend "RegRead, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppV\Client, Enabled, 0`r`n", IniFile
				FileAppend "IfEqual, Enabled, 0, EnableAppV,, Enable the App-V client on the system when it is not enabled.`r`n", IniFile
				FileAppend "RegRead, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppV\Client\Scripting, EnablePackageScripts, 0`r`n", IniFile
				FileAppend "IfEqual, EnablePackageScripts, 0, EnablePackageScripts,, Enable the App-V Package Scripts on the system when it is not enabled.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[EnableAppV]`r`n", IniFile
				FileAppend "Powershell, Enable-Appv`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[EnablePackageScripts]`r`n", IniFile
				FileAppend "Powershell, Set-AppvClientConfiguration -EnablePackageScripts 1`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Check_Install]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install,, Register installation when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_Install]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %DisplayName%, 1`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyStart%", IniFile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", IniFile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyProgress%", IniFile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%DisplayName%", IniFile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", IniFile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishInstall", "UserScript"
				IniWrite "5", IniFile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", IniFile, "UnInstall", "LogFile"
				IniWrite WorkingDir, IniFile, "UnInstall", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageId%", IniFile, "UnInstall", "SetSourceDir"
				IniWrite "1", IniFile, "UnInstall", "DeleteSourceDir"				
				IniWrite "", IniFile, "UnInstall", "TaskKill"
				IniWrite ">Stop-AppVClientPackage -Name '%DisplayName%' -Global \n Remove-AppVClientPackage -PackageId %PackageId% -VersionId %VersionId%", IniFile, "UnInstall", "Command1"
				IniWrite "explorer.exe", IniFile, "UnInstall", "CheckProcess"
				IniWrite "", IniFile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", IniFile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", IniFile, "UnInstall", "NotifyFinish"
				IniWrite "", IniFile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", IniFile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_UnInstall]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall,, Register removal when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_UnInstall]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %DisplayName%, 0`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyStart%", IniFile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", IniFile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyProgress%", IniFile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%DisplayName%", IniFile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", IniFile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "5", IniFile, "NotifyFinishUnInstall", "Wait"
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected APPV File:`n`n" IniFile, ProductName " " ProductVersion, "64 T30"
				}
			If (Extension = "msix" OR Extension = "appx")
				{
				GetProductName := ""
				GetProductVersion := ""
				GetProcessorArchitecture := ""
				GetDisplayName := ""
				GetPublisher := ""
				GetDescription := ""
				AppxManifestXml := Read_AppxManifest(SetupFilePath)
				Loop parse, AppxManifestXml, "`n", "`r"
					{
					If InStr(A_LoopField, "<Identity")
						{
						Identity := Trim(A_LoopField)
						Identity := StrReplace(Identity, "<Identity")
						Identity := StrReplace(Identity, "/>")
						Identity := Trim(Identity)
						Identity := StrReplace(Identity, "Name=", "`nName=")
						Identity := StrReplace(Identity, "Publisher=", "`nPublisher=")
						Identity := StrReplace(Identity, "Version=", "`nVersion=")
						Identity := StrReplace(Identity, "ProcessorArchitecture=", "`nProcessorArchitecture=")
						Loop parse, Identity, "`n"
							{
							If InStr(A_LoopField, "Name=")
								{
								GetProductName := StrReplace(A_LoopField, "Name=")
								GetProductName := StrReplace(GetProductName, "`"")
								}
							If InStr(A_LoopField, "Version=")
								{
								GetProductVersion := StrReplace(A_LoopField , "Version=")
								GetProductVersion := StrReplace(GetProductVersion, "`"")
								}
							If InStr(A_LoopField, "ProcessorArchitecture=")
								{
								GetProcessorArchitecture := StrReplace(A_LoopField , "ProcessorArchitecture=")
								GetProcessorArchitecture := StrReplace(GetProcessorArchitecture, "`"")
								}
							}
						}
					If InStr(A_LoopField, "<DisplayName>")
						{
						GetDisplayName := Trim(A_LoopField)
						GetDisplayName := StrReplace(GetDisplayName , "<DisplayName>")
						GetDisplayName := StrReplace(GetDisplayName , "</DisplayName>")
						}
					If InStr(A_LoopField, "<PublisherDisplayName>")
						{
						GetPublisher := Trim(A_LoopField)
						GetPublisher := StrReplace(GetPublisher, "<PublisherDisplayName>")
						GetPublisher := StrReplace(GetPublisher, "</PublisherDisplayName>")
						}
					If InStr(A_LoopField, "<Description>")
						{
						GetDescription := Trim(A_LoopField)
						GetDescription := StrReplace(GetDescription, "<Description>")
						GetDescription := StrReplace(GetDescription, "</Description>")
						}
					}
				AppxManifestXml := ""
				LogFileName := GetDisplayName
				loop parse ReservedCharacters
					If InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")				
				IniWrite A_UserName, IniFile, "Variables", "ScriptAuthor"
				IniWrite FormatTime(, "ShortDate"), IniFile, "Variables", "ScriptDate"
				IniWrite GetProductName, IniFile, "Variables", "ProductName"
				IniWrite GetProductVersion, IniFile, "Variables", "ProductVersion"
				IniWrite GetProcessorArchitecture, IniFile, "Variables", "ProcessorArchitecture"
				IniWrite GetDisplayName, IniFile, "Variables", "DisplayName"
				IniWrite GetPublisher, IniFile, "Variables", "Publisher"
				IniWrite GetDescription, IniFile, "Variables", "PackageDescription"
				IniWrite SetupFileNameNoExt, IniFile, "Variables", "PackageName"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", IniFile, "Variables", "LogFolder"
				IniWrite LogFileName, IniFile, "Variables", "LogFileName"
				IniWrite SetupFileName, IniFile, "Variables", "SetupFileName"
				FileAppend "`r`n", IniFile
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", IniFile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed from your system. Save your work and close this application when it is active.", IniFile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", IniFile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", IniFile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being added to your system. Do not turn off your computer during the installation. Please wait...", IniFile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", IniFile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", IniFile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", IniFile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", IniFile, "0409", "DeferLinkText"
				IniWrite "Continue", IniFile, "0409", "ContinueLinkText"
				IniWrite "Close", IniFile, "0409", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", IniFile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", IniFile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", IniFile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt toegevoegd aan het systeem. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", IniFile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", IniFile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", IniFile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", IniFile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", IniFile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", IniFile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", IniFile, "0413", "CloseLinkText"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", IniFile, "Install", "LogFile"
				IniWrite WorkingDir, IniFile, "Install", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageName%", IniFile, "Install", "SetSourceDir"
				IniWrite "0", IniFile, "Install", "DeleteSourceDir"
				IniWrite "", IniFile, "Install", "TaskKill"
				IniWrite ">Add-AppProvisionedPackage -online -packagepath '%SetupFileName%' -skiplicense", IniFile, "Install", "Command1"
				IniWrite "explorer.exe", IniFile, "Install", "CheckProcess"
				IniWrite "", IniFile, "Install", "NotifyPendingReboot"
				IniWrite "", IniFile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", IniFile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", IniFile, "Install", "NotifyFinish"
				IniWrite "", IniFile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", IniFile, "Install", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_Install]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install,, Register installation when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_Install]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %DisplayName%, 1`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyStart%", IniFile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", IniFile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyProgress%", IniFile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%DisplayName%", IniFile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", IniFile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishInstall", "UserScript"
				IniWrite "5", IniFile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Execution", IniFile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", IniFile, "UnInstall", "LogFile"
				IniWrite WorkingDir, IniFile, "UnInstall", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageName%", IniFile, "UnInstall", "SetSourceDir"
				IniWrite "1", IniFile, "UnInstall", "DeleteSourceDir"				
				IniWrite "", IniFile, "UnInstall", "TaskKill"
				IniWrite ">Remove-AppPackage -AllUsers -package '%PackageName%'", IniFile, "UnInstall", "Command1"
				IniWrite "explorer.exe", IniFile, "UnInstall", "CheckProcess"
				IniWrite "", IniFile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", IniFile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", IniFile, "UnInstall", "NotifyFinish"
				IniWrite "", IniFile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", IniFile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", IniFile
				FileAppend "[Check_UnInstall]`r`n", IniFile
				FileAppend "IfEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall,, Register removal when the last return code is successful.`r`n", IniFile
				FileAppend "`r`n", IniFile
				FileAppend "[Register_UnInstall]`r`n", IniFile
				FileAppend "RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher, %DisplayName%, 0`r`n", IniFile
				FileAppend "`r`n", IniFile
				IniWrite "Notification", IniFile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyStart%", IniFile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", IniFile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", IniFile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", IniFile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", IniFile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", IniFile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", IniFile
				IniWrite "Progress", IniFile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%DisplayName%", IniFile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyProgress%", IniFile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", IniFile
				IniWrite "Message", IniFile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", IniFile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%DisplayName%", IniFile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", IniFile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", IniFile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", IniFile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "Status"
				IniWrite "", IniFile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "5", IniFile, "NotifyFinishUnInstall", "Wait"
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected MSIX/APPX File:`n`n" IniFile, ProductName " " ProductVersion, "64 T30"
				}
			}
		ExitApp 0
		}
	}

; --------------------------------------------------------------- GUI functions for clicking continue ---------------------------------------------------------------
CloseNotificationContinue(*)
	{
	Global UserInteraction
	UserInteraction := 1
	NotifyGui.Destroy()
	Send_NamedPipeMessage("0", UserPipeName)
	}

; --------------------------------------------------------------- GUI functions for clicking defer (cancel) ---------------------------------------------------------------
CloseNotificationRetry(*)
	{
	Global UserInteraction
	UserInteraction := 1
	ProductNameText := ShowProductName.Value
	ProductVersionText := ShowProductVersion.Value
	Section := NotifyGui.Title
	If ProductVersionText = ""
		ProductVersionText := "1"
	DeferredCount := RegRead("HKEY_CURRENT_USER\SOFTWARE\" . ProductName . "\DeferTimes", Section . " " . ProductNameText . " " . ProductVersionText, 0)
	DeferredCount++
	RegWrite DeferredCount, "REG_DWORD", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName . "\DeferTimes", Section . " " . ProductNameText . " " . ProductVersionText
	NotifyGui.Destroy()
	Send_NamedPipeMessage("1602", UserPipeName)
	}

; --------------------------------------------------------------- Timer (USER) to check if the SYSTEM process (PID) is up and running ---------------------------------------------------------------
SystemProcessCheck()
	{
	Global System_PID
	If System_PID
		{
		If !(ProcessExist(System_PID))
			{
			NotifyGui.Destroy()
			MsgBox "The communication is lost with the main (SYSTEM) process with PID: " System_PID, ProductName " " ProductVersion, "16 T30"
			ExitApp 1
			}
		}
	}

; --------------------------------------------------------------- Run and wait script item timer notification progress bar ---------------------------------------------------------------
RunWaitProgress(ProgressBar := 0)
	{
	NotifyProgress.Value := ProgressBar++
	If ProgressBar = 100
		ProgressBar := 0
	Sleep 50
	}

; --------------------------------------------------------------- Function for detecting system shutdown/logoff and allows the user to abort it ---------------------------------------------------------------
On_WM_QUERYENDSESSION(wParam, lParam, *)
	{
	Global ProductName
	ENDSESSION_LOGOFF := 0x80000000
	If (lParam & ENDSESSION_LOGOFF)
		EventType := "Logoff"
	Else
		EventType := "Shutdown"
	Try
		{
		BlockShutdown(ProductName . " attempting to prevent " . EventType ".")
		Return False
		}
	}

BlockShutdown(Reason)
	{
	DllCall("ShutdownBlockReasonCreate", "ptr", A_ScriptHwnd, "wstr", Reason)
	OnExit StopBlockingShutdown
	}

StopBlockingShutdown(*)
	{
	Global UserPipeName
	Send_NamedPipeMessage("ExitApp", UserPipeName)
	OnExit StopBlockingShutdown, 0
	DllCall("ShutdownBlockReasonDestroy", "ptr", A_ScriptHwnd)
	}

; --------------------------------------------------------------- Function for sending a Named Pipe message and wait for the delivering ---------------------------------------------------------------
Send_NamedPipeMessage(Message := "", PipeName := "PipeName")
	{
	hPipe := DllCall("CreateNamedPipe", "str", "\\.\pipe\" . PipeName, "uint", 3, "uint", 0, "uint", 255, "uint", 0, "uint", 0, "uint", 0, "ptr", 0, "ptr")
	If (hPipe = -1)
		Return hPipe
	DllCall("ConnectNamedPipe", "ptr", hPipe, "ptr", 0)
	f := FileOpen(hPipe, "h")
	f.Write(Message . "`n")
	f.ReadLine()
	DllCall("CloseHandle", "ptr", hPipe)
	Return "delivered"
	}

; --------------------------------------------------------------- Functions for receiving a Named Pipe message ---------------------------------------------------------------
Receive_NamedPipeMessage(PipeName := "PipeName")
	{
	While !DllCall("WaitNamedPipe", "Str", "\\.\pipe\" . PipeName, "UInt", 0xffffffff)
		Sleep 250
	f := FileOpen("\\.\pipe\" . PipeName, "r")
	PipeMessage := f.ReadLine()
	f.Close()
	Return PipeMessage
	}

; --------------------------------------------------------------- Timer for setting the "UserLogin" and "RebootRequired" session environment variables to value 1 or 0 during runtime ---------------------------------------------------------------
UserLoginPendingReboot()
	{
	If (PID := ProcessExist("explorer.exe"))
		EnvSet "UserLogin", 1
	Else
		EnvSet "UserLogin", 0
	EnvSet "RebootRequired", PendingReboot()
	}

; --------------------------------------------------------------- Function to check if a reboot is pending on the system. Returns 1 (true) or 0 (false) ---------------------------------------------------------------
PendingReboot()
	{
	RebootRequired := RegKeyExists("HKLM", "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
	RebootPending := RegKeyExists("HKLM", "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
	PendingFileRenameOperations := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations", 0)
	If (RebootRequired = 1 or RebootPending = 1 or PendingFileRenameOperations != 0)
		Return 1
	Else
		Return 0
	}

; --------------------------------------------------------------- Function to check if a registry key exists. Returns 1 (true) or 0 (false) ---------------------------------------------------------------
RegKeyExists(RootKey := 0x80000002, SubKey := "SOFTWARE")
	{
	If (RootKey = "HKEY_CLASSES_ROOT" or RootKey = "HKCR")
		RootKey := 0x80000000
	If (RootKey = "HKEY_CURENT_USER" or RootKey = "HKCU")
		RootKey := 0x80000001
	If (RootKey = "HKEY_LOCAL_MACHINE" or RootKey = "HKLM")
		RootKey := 0x80000002
	If (RootKey = "HKEY_USERS" or RootKey = "HKU")
		RootKey := 0x80000003
	If (RootKey = "HKEY_CURRENT_CONFIG" or RootKey = "HKCC")
		RootKey := 0x80000005
	hKey := "69.420.parrot"
	ERROR_SUCCESS := 0
	KEY_READ := 131097
	RegKeyExists := ERROR_SUCCESS == DllCall("RegOpenKeyExW", "Ptr", RootKey, "Wstr", SubKey, "Uint", 0, "Uint", KEY_READ, "Str", hKey)
	DllCall("RegCloseKey", "Str", hKey)
	Return RegKeyExists
	}

; --------------------------------------------------------------- Function for getting the system error message text using the error number ---------------------------------------------------------------
GetSysErrorText(errNr){
	bufferSize := 1024
	VarSetStrCapacity(&RetBuffer, bufferSize)
	DllCall("FormatMessage", "UInt", FORMAT_MESSAGE_FROM_SYSTEM := 0x1000, "UInt", 0, "UInt", errNr, "UInt", 0, "Str", Retbuffer, "UInt", bufferSize, "UInt", 0)
	Return Retbuffer
	}

; --------------------------------------------------------------- Function for setting NTFS file permissions ---------------------------------------------------------------
SetSecurityFile(File, Trustee, AccessMask, Flags)
	{
	dacl := ComObject("AccessControlList")
	sd := ComObject("SecurityDescriptor")
	newAce := ComObject("AccessControlEntry")
	sdutil := ComObject("ADsSecurityUtility")
	sd := sdUtil.GetSecurityDescriptor(File, 1, 1)
	dacl := sd.DiscretionaryAcl
	newAce.Trustee := Trustee
	newAce.AccessMask := AccessMask
	newAce.AceFlags := Flags
	newAce.AceType := 0
	dacl.AddAce(newAce)
	sdutil.SetSecurityDescriptor(File, 1, sd, 1)
	Return A_LastError
	}

; --------------------------------------------------------------- Function for retrieving property values using the EXE file name and the property name ("Comments", "CompanyName", "FileDescription", "FileVersion", "InternalName", "LegalCopyright", "LegalTrademarks", "OriginalFilename", "PrivateBuild", "ProductName", "ProductVersion", "SpecialBuild") ---------------------------------------------------------------
GetFileVersionInfo(ExeFile := "", Property := "ProductName")
    {
	If FileExist(ExeFile)
		{
		If (Size := DllCall("version\GetFileVersionInfoSizeW", "Str", ExeFile, "Ptr", 0, "UInt"))
			{
			Data := Buffer(Size)
			If (DllCall("version\GetFileVersionInfoW", "Str", ExeFile, "UInt", 0, "UInt", Data.Size, "Ptr", Data))
				{
				If (DllCall("version\VerQueryValueW", "Ptr", Data, "Str", "\VarFileInfo\Translation", "Ptr*", &Buf := 0, "UInt*", &Len := 0))
					{
					LangCP := Format("{:04X}{:04X}", NumGet(Buf + 0, "UShort"), NumGet(Buf + 2, "UShort"))
					FileInfo := Map()
					If (DllCall("version\VerQueryValueW", "Ptr", Data, "Str", "\StringFileInfo\" . LangCP . "\" . Property, "Ptr*", &Buf, "UInt*", &Len))
						{
						FileInfo[Property] := StrGet(Buf, Len, "UTF-16")
						Return Trim(FileInfo[Property])
						}
					}
				}
			}
		}
    }

; --------------------------------------------------------------- Function for retrieving the AppxManifest.xml file content from an App-V (*.appv) or MSIX (*.msix) package file ---------------------------------------------------------------
Read_AppxManifest(AppxFile := "")
	{
	If FileExist(AppxFile)
		{
		FileMove AppxFile, AppxFile . ".zip"
		psh := ComObject("Shell.Application")
		psh.Namespace(A_Temp).CopyHere(psh.Namespace(AppxFile . ".zip").items.Item("AppxManifest.xml"), 4|16)
		FileMove AppxFile  . ".zip", AppxFile
		While !FileExist(A_Temp . "\AppxManifest.xml")
			Sleep 50
		AppxManifestXml := FileRead(A_Temp . "\AppxManifest.xml")
		FileDelete A_Temp . "\AppxManifest.xml"
		Return AppxManifestXml
		}
	}

; --------------------------------------------------------------- Function for retrieving property values using the MSI file name and the property name ---------------------------------------------------------------
WindowsInstaller(MsiFile := "", Property := "")
	{
	If FileExist(MsiFile)
		{
		installer := ComObject("WindowsInstaller.Installer")
		If !database := installer.OpenDatabase(MsiFile, 0)
			Return A_LastError
		view := database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = '" Property "'")
		view.Execute
		record := view.Fetch
		Return record.StringData(1)
		}
  	}

; --------------------------------------------------------------- Function for directly installing or removing a Windows Installer MSI file ---------------------------------------------------------------
MsiInstallProduct(PackagePath := "", CommandLine := "", LogFile := "")
	{
	Installer := ComObject("WindowsInstaller.Installer")
	If LogFile
		Installer.EnableLog "voicewarmup", LogFile
	Installer.UILevel := 2
	Return DllCall("Msi.dll\MsiInstallProductW", "Wstr", PackagePath, "Wstr", CommandLine)
	}

; --------------------------------------------------------------- Function for creating a scheduled task at next logon ---------------------------------------------------------------
CreateTask(TaskName := "", Description := "", Path := "", Arguments := "", ExecutionTimeLimit := "12")
	{
	Local TriggerType := 9
	Local ActionTypeExec := 0
	Local LogonType := 5
	Local TaskCreateOrUpdate := 6
	service := ComObject("Schedule.Service")
	service.Connect()
	rootFolder := service.GetFolder("\")
	taskDefinition := service.NewTask(0)
	regInfo := taskDefinition.RegistrationInfo
	regInfo.Description := Description
	regInfo.Author := A_UserName
	principal := taskDefinition.Principal
	principal.LogonType := LogonType
	principal.UserId := "S-1-5-18"
	settings := taskDefinition.Settings
	settings.DisallowStartIfOnBatteries := True
	settings.StopIfGoingOnBatteries := False
	settings.AllowHardTerminate := True
	settings.StartWhenAvailable := False
	settings.RunOnlyIfNetworkAvailable := False
	settings.Enabled := True
	settings.Hidden := False
	settings.RunOnlyIfIdle := False
	settings.DisallowStartOnRemoteAppSession := False
	settings.UseUnifiedSchedulingEngine := True
	settings.WakeToRun := False
	settings.ExecutionTimeLimit := "PT" ExecutionTimeLimit "H"
	triggers := taskDefinition.Triggers
	trigger := triggers.Create(TriggerType)
	trigger.Id := "LogonTriggerId"
	trigger.Enabled := True
	Action := taskDefinition.Actions.Create(ActionTypeExec)
	Action.Path := Path
	Action.Arguments := Arguments
	Try rootFolder.RegisterTaskDefinition(TaskName, taskDefinition, TaskCreateOrUpdate , "", "", 4)
	Catch as err
		Return err.Message
	Else
		Return 0
  	}

; --------------------------------------------------------------- Function for removing a scheduled task ---------------------------------------------------------------
DeleteTask(TaskName := "")
	{
	service := ComObject("Schedule.Service")
	service.Connect()
	rootFolder := service.GetFolder("\")
	Try rootFolder.DeleteTask(TaskName, 0)
	Catch as err
		Return err.Message
	Else
		Return 0
	}

; --------------------------------------------------------------- Function for translating existing environment variables to their actual values ---------------------------------------------------------------
TransForm(Str) 
	{
	spo := 1
	out := ""
	Str := Trim(Str)
	While (fpo := RegexMatch(Str, "(%(.*?)%)|``(.)", &m, spo))
		{
		out .= SubStr(Str, spo, fpo-spo)
		spo := fpo + StrLen(m[0])
		If (m[1])
			out .= (env := EnvGet(m[2])) ? env : m[1]
		Else
 			out .= m[3]
		}
	return out SubStr(Str, spo)
	}

; --------------------------------------------------------------- Function for comparing two (version) numbers. If the CompareValue is higher the result will be greater. will return equal, greater or smaller depending on the (version) number values ---------------------------------------------------------------
CompareVersionString(CheckValue := 0, CompareValue := 0)
	{
	Result := VerCompare(CompareValue, CheckValue)
	If Result = 0
		Return "Equal"
	If Result = 1
		Return "Greater"
	If Result = -1
		Return "Smaller"
	}

; --------------------------------------------------------------- Function for creating a new (user) process using the token of an existing running process name. (if wait is 1 the process exit code will return or else, the PID will return) ---------------------------------------------------------------
CreateProcessAsUser(cmd, args := "", workingDir := "", processName := "", wait := 0) 
	{
	Global LogFile
	static token := 0
	static startupInfo := 0
	static processInfo := 0
	static env := 0
	startupInfo := Buffer(7 * A_PtrSize + 9 * 4 + 2 * 2, 0)
	NumPut("Int", startupInfo.Size, startupInfo)
	NumPut("Int", 1, startupInfo, 4 + 3 * A_PtrSize + 4 * 7)
	NumPut("UShort", 5, startupInfo, 4 + 3 * A_PtrSize + 5 * 7)
	processInfo := Buffer(2 * A_PtrSize + 2 * 4, 0)
	If !(ProcessId := ProcessExist(processName))
		{
		FileAppend A_Now "`t[ ERROR ]`t[ CreateProcessAsUser ] Process not found: " processName "`r`n" , LogFile
		Return -1
		}
	DllCall("ProcessIdToSessionId", "Uint", ProcessId, "Ptr*", &SessionId := 0)
	If !SessionId
		{
		FileAppend A_Now "`t[ ERROR ]`t[ CreateProcessAsUser ] ProcessIdToSessionId function failed with error: " A_LastError "`r`n" , LogFile
		Return -1
		}
	DllCall("wtsapi32\WTSQueryUserToken", "Int", SessionId, "Ptr*", &hToken := 0)
	DllCall("CloseHandle", "Ptr", SessionId)
	DllCall("Userenv.dll\CreateEnvironmentBlock", "Ptr*", &lpEnvironment := 0, "Ptr", hToken, "Int", 0)
	If args
		cmd := cmd " " args
	If !DllCall("Advapi32.dll\CreateProcessAsUserW", "Ptr", hToken, "Ptr", 0, "WStr", cmd, "Int", 0, "Int", 0, "Int", 0, "Int", 0x00000400, "Ptr", lpEnvironment, "WStr", workingDir, "Ptr", startupInfo, "Ptr", processInfo)
		{
		If FileExist(LogFile)
			FileAppend A_Now "`t[ ERROR ]`t[ CreateProcessAsUser ] CreateProcessAsUserW function failed with error: " A_LastError "`r`n" , LogFile
		DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
		DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
		Return -2
		}
	If wait = 1
		{
		loop
			{
			DllCall("GetExitCodeProcess", "Ptr", NumGet(processInfo, 0, "Ptr"), "UIntP" , &exitCode := 0)
			If exitCode = 259
				Sleep 500
			Else
				Break
			}
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, A_PtrSize, "Ptr"))
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, 0, "Ptr"))
		DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
		DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
		Return exitCode
		}
	Else
		{
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, A_PtrSize, "Ptr"))
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, 0, "Ptr"))
		DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
		DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
		Return NumGet(processInfo, 2 * A_PtrSize, "Int")
		}
	}

; --------------------------------------------------------------- Function to determine if other sessions of the Intune Win32 Launcher are running with the option to stop these sessions ---------------------------------------------------------------
CheckForOtherSessions(Stop := 0)
	{
	Global A_PID
	Global LogFile
	FoundSessions := 0
	s := 4096
	h := DllCall("OpenProcess", "UInt", 0x0400, "Int", false, "UInt", A_PID, "Ptr")
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
		If !h
			Continue
		n := Buffer(s, 0)
		e := DllCall("Psapi.dll\GetModuleBaseName", "Ptr", h, "Ptr", 0, "Ptr", n, "UInt", s//2)
		If !e
			e := DllCall("Psapi.dll\GetProcessImageFileName", "Ptr", h, "Ptr", n, "UInt", s//2)
		SplitPath StrGet(n), &n
		DllCall("CloseHandle", "Ptr", h)
		If (n && e)
			{
			If n = A_ScriptName
				{
				If id != A_PID
					{
					FoundSessions++
					If Stop = 1
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tStop notification process with Process Id: " id "`r`n" , LogFile
						ProcessClose id
						}
					}
				}
			}
		}
	DllCall("FreeLibrary", "Ptr", hModule)
	Return FoundSessions
	}

; --------------------------------------------------------------- Function for converting the number of seconds to the mm:ss time format ---------------------------------------------------------------
FormatSeconds(NumberOfSeconds) 
	{
	time := 19990101
	time := DateAdd(time, NumberOfSeconds, "Seconds")
	Return FormatTime(time, "mm:ss")
	}

; --------------------------------------------------------------- Function for setting session environment variables ---------------------------------------------------------------
SetEnv(Section)
	{
	Global IniFile
	Found := 0
	Counter := 0
	Loop read, IniFile
		{
		If Found = 1
			{
			If A_LoopReadLine
				{
				If InStr(A_LoopReadLine, "=")
					{
					EnvSetString := StrSplit(A_LoopReadLine, "=")
					EnvSetName := EnvSetString[1]
					EnvSetValue := EnvSetString[2]
					EnvSetValue := TransForm(EnvSetValue)
					EnvSet EnvSetName, EnvSetValue
					Counter++
					}
				Else
					Break
				}
			}
		If A_LoopReadLine = "[" Section "]"
			Found := 1
		}
	Return Counter
	}

; --------------------------------------------------------------- psScript Manager Class ---------------------------------------------------------------
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
		If !DirExist(SetInstallDir)
			DirCreate SetInstallDir
		If !FileExist(SetInstallDir . "\psScript.dll")
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
		Try
			this.ComInstance := ComObject(this.ClassName)
		Catch as Err
			{
			this.Unregister(SetInstallDir)
			return err.message
			}
		Return 0
		}
	static Unregister(SetInstallDir := A_Temp)
		{
		If (this.ComInstance is ComObject)
			this.ComInstance := unset
		RegDeleteKey this.RootKey . this.ClassName
		RegDeleteKey this.RootKey . "CLSID\" . this.CLSID
		RegDelete this.RootKey . "Component Categories\" . this.Component, "0"
		If FileExist(SetInstallDir . "\psScript.dll")
			Try FileDelete SetInstallDir . "\psScript.dll"
		Return 0
		}
	}

; --------------------------------------------------------------- Function for executing PreLaunch, PostExit and User Script items ---------------------------------------------------------------
Script(Section)
	{
	Global IniFile
	Global LogFile
	Global ScriptSections
	Global cpauPID
	Global SystemPipeName
	Global UserPipeName
	Global UserScript
	If UserScript
		ProgressBar := 0
	Found := 0
	Counter := 0
	If Instr(ScriptSections, Section "`r`n")
		{
		If FileExist(LogFile)
			FileAppend A_Now "`t[ WARN  ]`tScript section " Section " is already executed. Each section can be executed only once for each PreLaunch and PostExit script section. Please check the if/else statements.`r`n" , LogFile
		Return Counter
		}
	ScriptSections := ScriptSections . Section "`r`n"
	Loop read, IniFile
		{
		If Found = 1
			{
			If A_LoopReadLine
				{
				If InStr(A_LoopReadLine, ",")
					{
					If UserScript
						{
						NotifyProgress.Value := ProgressBar++
						If ProgressBar = 100
							ProgressBar := 0
						}
					ScriptItemString := StrSplit(A_LoopReadLine, ",")
					ScriptItemName := ScriptItemString[1]
					ScriptItemName := Trim(ScriptItemName)
					If (ScriptItemName = "IfEqual" or ScriptItemName = "IfNotEqual" or ScriptItemName = "IfGreater" or ScriptItemName = "IfGreaterOrEqual" or ScriptItemName = "IfNotGreater" or ScriptItemName = "IfNotGreaterOrEqual" or ScriptItemName = "IfSmaller" or ScriptItemName = "IfSmallerOrEqual" or ScriptItemName = "IfNotSmaller" or ScriptItemName = "IfNotSmallerOrEqual")
						{
						CheckVariable := ""
						CheckValue := ""
						GotoIfSection := ""
						GotoElseSection := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								CheckVariable := ScriptItemString[2]
							If A_Index = 3
								CheckValue := ScriptItemString[3]
							If A_Index = 4
								GotoIfSection := ScriptItemString[4]
							If A_Index = 5
								GotoElseSection := ScriptItemString[5]
							If A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						CheckVariable := Trim(CheckVariable)
						CheckValue := Transform(CheckValue)
						GotoIfSection := Trim(GotoIfSection)
						GotoElseSection := Trim(GotoElseSection)
						UserDescription := TransForm(UserDescription)
						If ScriptItemName = "IfEqual"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							CheckValue := StrReplace(CheckValue, A_Space . "OR" . A_Space, "`n|`t")
							CheckValue := "`n|`t" . CheckValue
							Result := False
							Loop parse, CheckValue, "`n"
								{
								If A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									If Check_array[1] = "|"
										{
										If GetValue = Check_array[2]
											Result := True
										}
									}
								}
							If Result = True
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotEqual"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							CheckValue := StrReplace(CheckValue, A_Space . "OR" . A_Space, "`n|`t")
							CheckValue := "`n|`t" . CheckValue
							Result := False
							Loop parse, CheckValue, "`n"
								{
								If A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									If Check_array[1] = "|"
										{
										If GetValue != Check_array[2]
											Result := True
										}
									}
								}
							If Result = True
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfGreater"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If Result = "Greater"
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfGreaterOrEqual"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If (Result = "Greater" or Result = "Equal")
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotGreater"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If (Result = "Smaller" or Result = "Equal")
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotGreaterOrEqual"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If Result = "Smaller"
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfSmaller"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If Result = "Smaller"
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfSmallerOrEqual"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If (Result = "Smaller" or Result = "Equal")
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotSmaller"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If (Result = "Greater" or Result = "Equal")
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotSmallerOrEqual"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							If Result = "Greater"
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						}
					If (ScriptItemName = "IfExist" or ScriptItemName = "IfNotExist" or ScriptItemName = "IfProcessExist" or ScriptItemName = "IfNotProcessExist")
						{
						CheckFileFolderRegKey := ""
						GotoIfSection := ""
						GotoElseSection := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								CheckFileFolderRegKey := ScriptItemString[2]
							If A_Index = 3
								GotoIfSection := ScriptItemString[3]
							If A_Index = 4
								GotoElseSection := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						CheckFileFolderRegKey := TransForm(CheckFileFolderRegKey)
						GotoIfSection := Trim(GotoIfSection)
						GotoElseSection := Trim(GotoElseSection)
						UserDescription := TransForm(UserDescription)
						If ScriptItemName = "IfExist"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "OR" . A_Space, "`n|`t")
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "AND" . A_Space, "`n&`t")
							CheckFileFolderRegKey := "`n|`t" . CheckFileFolderRegKey
							Result := False
							Loop parse, CheckFileFolderRegKey, "`n"
								{
								If A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									If Check_array[1] = "|"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										If (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											If RegKeyExists(Path_Array[1], Check_array[2]) = 1
												Result := True
											}
										Else
											{
											SplitPath Check_array[2],,, &Extension
											If Extension
												{
												If FileExist(Check_array[2])
													Result := True
												}
											Else
												{
												If DirExist(Check_array[2])
													Result := True
												}
											}
										}
									If Check_array[1] = "&"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										If (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											If RegKeyExists(Path_Array[1], Check_array[2]) = 0
												Result := false
											}
										Else
											{
											SplitPath Check_array[2],,, &Extension
											If Extension
												{
												If Not FileExist(Check_array[2])
													Result := False
												}
											Else
												{
												If Not DirExist(Check_array[2])
													Result := False
												}
											}
										}
									}
								}
							If Result = true
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotExist"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "OR" . A_Space, "`n|`t")
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "AND" . A_Space, "`n&`t")
							CheckFileFolderRegKey := "`n|`t" . CheckFileFolderRegKey
							Result := False
							Loop parse, CheckFileFolderRegKey, "`n"
								{
								If A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									If Check_array[1] = "|"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										If (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											If RegKeyExists(Path_Array[1], Check_array[2]) = 0
												Result := True
											}
										Else
											{
											SplitPath Check_array[2],,, &Extension
											If Extension
												{
												If Not FileExist(Check_array[2])
													Result := True
												}
											Else
												{
												If Not DirExist(Check_array[2])
													Result := True
												}
											}
										}
									If Check_array[1] = "&"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										If (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											If RegKeyExists(Path_Array[1], Check_array[2]) = 1
												Result := false
											}
										Else
											{
											SplitPath Check_array[2],,, &Extension
											If Extension
												{
												If FileExist(Check_array[2])
													Result := False
												}
											Else
												{
												If DirExist(Check_array[2])
													Result := False
												}
											}
										}
									}
								}
							If Result = true
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfProcessExist"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfProcessExist statement for process (ID) name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							If (PID := ProcessExist(CheckFileFolderRegKey))
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotProcessExist"
							{
							If GotoElseSection
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotProcessExist statement for process (ID) name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'."
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tUsing IfNotProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
								Else
									If UserScript
										If UserDescription
											ShowInfo.Value := UserDescription
										Else
											ShowInfo.Value := "Using IfNotProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue..."
								}
							If Not (PID := ProcessExist(CheckFileFolderRegKey))
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
									   		ShowInfo.Value := "Performing script section: " GotoIfSection
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Performing script section: " GotoElseSection
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						}
					If ScriptItemName = "FileCopy"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Source := ScriptItemString[2]
							If A_Index = 3
								Destination := ScriptItemString[3]
							If A_Index = 4
								OverWrite := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFileCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FileCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						If FileExist(Source)
							{
							Try
								FileCopy Source, Destination, OverWrite
							Catch as Err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFileCopy failed for " err.Extra " files.`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tFileCopy failed for " err.Extra " files.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "FileCopy failed for " err.Extra " files."
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FileCopy successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "FileMove"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Source := ScriptItemString[2]
							If A_Index = 3
								Destination := ScriptItemString[3]
							If A_Index = 4
								OverWrite := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFileMove (File rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FileMove (File rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						If FileExist(Source)
							{
							Try
								FileMove Source, Destination, OverWrite
							Catch as Err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFileMove failed for " err.Extra " files.`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tFileMove failed for " err.Extra " files.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "FileMove failed for " err.Extra " files."
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileMove successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FileMove successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "FolderCopy"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Source := ScriptItemString[2]
							If A_Index = 3
								Destination := ScriptItemString[3]
							If A_Index = 4
								OverWrite := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFolderCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FolderCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						If (DirExist(Source) or FileExist(Source))
							{
							Try
								DirCopy Source, Destination, OverWrite
							Catch as err
								{
								If FileExist(LogFile)
									{
									If OverWrite = 1
										FileAppend A_Now "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`r`n" , LogFile
									Else
										If !UserDescription
											FileAppend A_Now "`t[ INFO  ]`tFolderCopy successfully executed.`r`n" , LogFile
									}
								Else
									If UserScript
										{
										If OverWrite = 1
											{
											FileAppend A_Now "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
											ShowInfo.Value := "FolderCopy failed with Error: " err.Message
											Sleep 10000
											}
										Else
											If !UserDescription
												ShowInfo.Value := "FolderCopy successfully executed."
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFolderCopy successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FolderCopy successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "FolderMove"
						{
						Source := ""
						Destination := ""
						OverWrite := "R"
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Source := ScriptItemString[2]
							If A_Index = 3
								Destination := ScriptItemString[3]
							If A_Index = 4
								OverWrite := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFolderMove (Folder rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FolderMove (Folder rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						If DirExist(Source)
							{
							Try
								DirMove Source, Destination, OverWrite
							Catch as err
								{
								If FileExist(LogFile)
									{
									If OverWrite = 1
										FileAppend A_Now "`t[ WARN  ]`tFolderMove failed with Error: " err.Message "`r`n" , LogFile
									Else
										If !UserDescription
											FileAppend A_Now "`t[ INFO  ]`tFolderMove successfully executed.`r`n" , LogFile
									}
								Else
									If UserScript
										{
										If OverWrite = 1
											{
											FileAppend A_Now "`t[ WARN  ]`tFolderMove failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
											ShowInfo.Value := "FolderMove failed with Error: " err.Message
											Sleep 10000
											}
										Else
											If !UserDescription
												ShowInfo.Value := "FolderMove successfully executed."
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFolderMove successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FolderMove successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "FolderCreate"
						{
						Destination := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Destination := ScriptItemString[2]
							If A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						Destination := TransForm(Destination)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFolderCreate: " Destination "`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FolderCreate: " Destination
						If Not FileExist(Destination)
							{
							Try
								DirCreate Destination
							Catch as Err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFolderCreate failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tFolderCreate failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "FolderCreate failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFolderCreate successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FolderCreate successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" Destination "' already exists."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "FileDelete"
						{
						FilePattern := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								FilePattern := ScriptItemString[2]
							If A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						FilePattern := TransForm(FilePattern)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFileDelete: " FilePattern "`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FileDelete: " FilePattern
						If FileExist(FilePattern)
							{
							Try
								FileDelete FilePattern
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFileDelete failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tFileDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "FileDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileDelete successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FileDelete successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" FilePattern "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" FilePattern "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" FilePattern "' not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "FolderDelete"
						{
						DirName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								DirName := ScriptItemString[2]
							If A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						DirName := TransForm(DirName)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFolderDelete: " DirName "`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FolderDelete: " DirName
						If DirExist(DirName)
							{
							Try
								DirDelete DirName, 1
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFolderDelete failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tFolderDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "FolderDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFolderDelete successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "FolderDelete successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" DirName "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" DirName "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" DirName "' not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "EnvSet"
						{
						EnvSetName := ""
						EnvSetValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								EnvSetName := ScriptItemString[2]
							If A_Index = 3
								EnvSetValue := ScriptItemString[3]
							If A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						EnvSetName := Trim(EnvSetName)
						EnvSetValue := TransForm(EnvSetValue)
						UserDescription := TransForm(UserDescription)
						If EnvSetValue
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvSet setting system environment variable name '" EnvSetName "' with value: " EnvSetValue "`r`n" , LogFile
							Else
								If UserScript
									If UserDescription
										ShowInfo.Value := UserDescription
									Else
										ShowInfo.Value := "EnvSet setting user environment variable name '" EnvSetName "' with value: " EnvSetValue
							If UserScript
								{
								Try
									RegWrite EnvSetValue, "REG_EXPAND_SZ", "HKEY_CURRENT_USER\Environment", EnvSetName
								Catch as err
									{
									ShowInfo.Value := "User environment variable registry (HKEY_CURRENT_USER\Environment) failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message
									Sleep 10000
									}
								Else
									{
									Try
										EnvSet EnvSetName, EnvSetValue
									Catch as err
										{
										ShowInfo.Value := "User environment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message
										Sleep 10000
										}
									Else
										{
										SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
										If !UserDescription
											ShowInfo.Value := "User environment variable successfully set for '" EnvSetName "' with value: " EnvSetValue
										}
									}
								}
							Else
								{
								Try
									RegWrite EnvSetValue, "REG_EXPAND_SZ", "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", EnvSetName
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tSystem environment variable registry (HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment) failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									Try
										EnvSet EnvSetName, EnvSetValue
									Catch as err
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ WARN  ]`tSystem environment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
										}
									Else
										{
										SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`tSystem environment variable successfully set for '" EnvSetName "' with value: " EnvSetValue "`r`n" , LogFile
										}
									}
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvSet remove system environment variable name: " EnvSetName "`r`n" , LogFile
							Else
								If UserScript
									If UserDescription
										ShowInfo.Value := UserDescription
									Else
										ShowInfo.Value := "EnvSet remove user environment variable name: " EnvSetName
							If UserScript
								{
								EnvGetValue := RegRead("HKEY_CURRENT_USER\Environment", EnvSetName, 0)
								If EnvGetValue != 0
									{
									Try
										RegDelete "HKEY_CURRENT_USER\Environment", EnvSetName
									Catch as err
										{
										ShowInfo.Value := "User environment variable registry (HKEY_CURRENT_USER\Environment) failed to remove for '" EnvSetName "' Error: " err.Message
										Sleep 10000
										}
									Else
										{
										Try
											EnvSet EnvSetName
										Catch as err
											{
											ShowInfo.Value := "User environment variable failed to remove for '" EnvSetName "' Error: " err.Message
											Sleep 10000
											}
										Else
											{
											SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
											If !UserDescription
												ShowInfo.Value := "User environment variable successfully removed for variable name: " EnvSetName
											}
										}
									}
								Else
									{
									ShowInfo.Value := "User environment variable '" EnvSetName "' not found."
									Sleep 10000										
									}
								}
							Else
								{
								EnvGetValue := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", EnvSetName, 0)
								If EnvGetValue != 0
									{
									Try
										RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", EnvSetName
									Catch as err
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ WARN  ]`tSystem environment variable registry (HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment) failed to remove for '" EnvSetName "' Error: " err.Message "`r`n" , LogFile
										}
									Else
										{
										Try
											EnvSet EnvSetName
										Catch as err
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ WARN  ]`tSystem environment variable failed to remove for '" EnvSetName "' Error: " err.Message "`r`n" , LogFile
											}
										Else
											{
											SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`tSystem environment variable successfully removed for variable name: " EnvSetName "`r`n" , LogFile
											}
										}
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tSystem environment variable '" EnvSetName "' not found.`r`n" , LogFile
									}
								}
							}
						}
					If ScriptItemName = "RegRead"
						{
						KeyName := ""
						ValueName := ""
						DefaultValue := ""
						EnvSetValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								KeyName := ScriptItemString[2]
							If A_Index = 3
								ValueName := ScriptItemString[3]
							If A_Index = 4
								DefaultValue := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						KeyName := TransForm(KeyName)
						ValueName := Trim(ValueName)
						DefaultValue := TransForm(DefaultValue)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tRegRead registry item '" ValueName "' in '" KeyName "' with default value '" DefaultValue "'`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "RegRead registry item '" ValueName "' in '" KeyName "' with default value '" DefaultValue "'"
						Try
							EnvSetValue := RegRead(KeyName, ValueName, DefaultValue)
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tRegRead failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tRegRead failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "RegRead failed with Error: " err.Message
									Sleep 10000
									}
							EnvSetValue := DefaultValue
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegRead successfully executed with value: " EnvSetValue "`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "RegRead successfully executed with value: " EnvSetValue
							}
						Try
							EnvSet ValueName, EnvSetValue
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "Environment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" ValueName "' with value: " EnvSetValue "`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "Environment variable successfully set for '" ValueName "' with value: " EnvSetValue
							}
						}
					If ScriptItemName = "RegWrite"
						{
						ValueType := ""
						KeyName := ""
						ValueName := ""
						Value := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								ValueType := ScriptItemString[2]
							If A_Index = 3
								KeyName := ScriptItemString[3]
							If A_Index = 4
								ValueName := ScriptItemString[4]
							If A_Index = 5
								Value := ScriptItemString[5]
							If A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						ValueType := TransForm(ValueType)
						KeyName := TransForm(KeyName)
						ValueName := TransForm(ValueName)
						Value := TransForm(Value)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tRegWrite create '" ValueType "' registry item '" ValueName "' in '" KeyName "' with value '" Value "'`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "RegWrite create '" ValueType "' registry item '" ValueName "' in '" KeyName "' with value '" Value "'"
						Try
							RegWrite Value, ValueType, KeyName, ValueName
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tRegWrite failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tRegWrite failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "RegWrite failed with Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegWrite successfully executed.`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "RegWrite successfully executed."
							}
						}
					If ScriptItemName = "RegDelete"
						{
						KeyName := ""
						ValueName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								KeyName := ScriptItemString[2]
							If A_Index = 3
								ValueName := ScriptItemString[3]
							If A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						KeyName := TransForm(KeyName)
						ValueName := TransForm(ValueName)
						UserDescription := TransForm(UserDescription)
						If ValueName
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegDelete registry value '" ValueName "' from '" KeyName "'`r`n" , LogFile
							Else
								If UserScript
									If UserDescription
										ShowInfo.Value := UserDescription
									Else
										ShowInfo.Value := "RegDelete registry value '" ValueName "' from '" KeyName "'"
							Try
								RegDelete KeyName, ValueName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "RegDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegDelete successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "RegDelete successfully executed."
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegDelete registry key: " KeyName "`r`n" , LogFile
							Else
								If UserScript
									If UserDescription
										ShowInfo.Value := UserDescription
									Else
										ShowInfo.Value := "RegDelete registry key: " KeyName
							Try
								RegDeleteKey KeyName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "RegDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegDelete successfully executed.`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "RegDelete successfully executed."
								}
							}
						}
					If ScriptItemName = "FileVersion"
						{
						Filename := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Filename := ScriptItemString[2]
							If A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						Filename := TransForm(Filename)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tRetrieve the version for file name: " Filename "'`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "Retrieve the version for file name: " Filename
						If FileExist(Filename)
							{
							FileVersion := FileGetVersion(Filename)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileVersion successfully executed with value: " FileVersion "`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "FileVersion successfully executed with value: " FileVersion
							Try
								EnvSet "FileVersion", FileVersion
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for 'FileVersion' with value '" FileVersion "' Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for 'FileVersion' with value '" FileVersion "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "Environment variable failed to set for 'FileVersion' with value '" FileVersion "' Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for 'FileVersion' with value: " FileVersion "`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Environment variable successfully set for 'FileVersion' with value: " FileVersion
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tFile not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tFile not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "File not found."
									Sleep 10000
									}
							}
						}
					If ScriptItemName = "IniRead"
						{
						Filename := ""
						Section := ""
						KeyName := ""
						DefaultValue := ""
						IniValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Filename := ScriptItemString[2]
							If A_Index = 3
								Section := ScriptItemString[3]
							If A_Index = 4
								KeyName := ScriptItemString[4]
							If A_Index = 5
								DefaultValue := ScriptItemString[5]
							If A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						Filename := TransForm(Filename)
						Section := Trim(Section)
						KeyName := Trim(KeyName)
						DefaultValue := TransForm(DefaultValue)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tIniRead from file name '" Filename "' using section name '" Section "' for key name '" KeyName "' with default value '" DefaultValue "'`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "IniRead from file name '" Filename "' using section name '" Section "' for key name '" KeyName "' with default value '" DefaultValue "'"
						Try
							IniValue := IniRead(Filename, Section, KeyName, DefaultValue)
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tIniRead failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tIniRead failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "IniRead failed with Error: " err.Message
									Sleep 10000
									}
							IniValue := DefaultValue
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniRead successfully executed with value: " IniValue "`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "IniRead successfully executed with value: " IniValue
							}
						Try
							EnvSet KeyName, IniValue
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "Environment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" KeyName "' with value: " IniValue "`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "Environment variable successfully set for '" KeyName "' with value: " IniValue
							}
						}
					If ScriptItemName = "IniWrite"
						{
						Value := ""
						Filename := ""
						Section := ""
						KeyName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Value := ScriptItemString[2]
							If A_Index = 3
								Filename := ScriptItemString[3]
							If A_Index = 4
								Section := ScriptItemString[4]
							If A_Index = 5
								KeyName := ScriptItemString[5]
							If A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						Value := TransForm(Value)
						Filename := TransForm(Filename)
						Section := Trim(Section)
						KeyName := Trim(KeyName)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tIniWrite value '" Value "' in file name '" Filename "' for section name '" Section "' with key name '" KeyName "'`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "IniWrite value '" Value "' in file name '" Filename "' for section name '" Section "' with key name '" KeyName "'"
						Try
							IniWrite Value, Filename, Section, KeyName
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tIniWrite failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tIniWrite failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "IniWrite failed with Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniWrite successfully executed.`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "IniWrite successfully executed."
							}
						}
					If ScriptItemName = "IniDelete"
						{
						Filename := ""
						Section := ""
						KeyName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Filename := ScriptItemString[2]
							If A_Index = 3
								Section := ScriptItemString[3]
							If A_Index = 4
								KeyName := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Filename := TransForm(Filename)
						Section := Trim(Section)
						KeyName := Trim(KeyName)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tIniDelete using file name '" Filename "' for section name '" Section "' (with key name '" KeyName "')`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "IniDelete using file name '" Filename "' for section name '" Section "' (with key name '" KeyName "')"
						Try
							IniDelete Filename, Section, KeyName
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tIniDelete failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tIniDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "IniDelete failed with Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniDelete successfully executed.`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "IniDelete successfully executed."
							}
						}
					If ScriptItemName = "FileCreateShortcut"
						{
						ShortcutTarget := ""
						ShortcutLinkFile := ""
						ShortcutWorkingDir := ""
						ShortcutArgs := ""
						ShortcutDescription := ""
						ShortcutIconFile := ""
						ShortcutKey := ""
						ShortcutIconNumber := "1"
						ShortcutRunState := "1"
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								ShortcutTarget := ScriptItemString[2]
							If A_Index = 3
								ShortcutLinkFile := ScriptItemString[3]
							If A_Index = 4
								ShortcutWorkingDir := ScriptItemString[4]
							If A_Index = 5
								ShortcutArgs := ScriptItemString[5]
							If A_Index = 6
								ShortcutDescription := ScriptItemString[6]
							If A_Index = 7
								ShortcutIconFile := ScriptItemString[7]
							If A_Index = 8
								ShortcutKey := ScriptItemString[8]
							If A_Index = 9
								ShortcutIconNumber := ScriptItemString[9]
							If A_Index = 10
								ShortcutRunState := ScriptItemString[10]
							If A_Index = 11
								UserDescription := ScriptItemString[11]
							}
						ShortcutTarget := TransForm(ShortcutTarget)
						ShortcutLinkFile := TransForm(ShortcutLinkFile)
						ShortcutWorkingDir := TransForm(ShortcutWorkingDir)
						ShortcutArgs := TransForm(ShortcutArgs)
						ShortcutDescription := TransForm(ShortcutDescription)
						ShortcutIconFile := TransForm(ShortcutIconFile)
						ShortcutKey := Trim(ShortcutKey)
						ShortcutIconNumber := Trim(ShortcutIconNumber)
						ShortcutRunState := Trim(ShortcutRunState)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFileCreateShortcut using the following properties: Target = '" ShortcutTarget "' LinkFile = '" ShortcutLinkFile "' WorkingDir = '" ShortcutWorkingDir "' Args = '" ShortcutArgs "' Description = '" ShortcutDescription "' IconFile = '" ShortcutIconFile "' ShortcutKey = '" ShortcutKey "' IconNumber = '" ShortcutIconNumber "' RunState = '" ShortcutRunState "'`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FileCreateShortcut using the following properties: Target = '" ShortcutTarget "' LinkFile = '" ShortcutLinkFile "' WorkingDir = '" ShortcutWorkingDir "' Args = '" ShortcutArgs "' Description = '" ShortcutDescription "' IconFile = '" ShortcutIconFile "' ShortcutKey = '" ShortcutKey "' IconNumber = '" ShortcutIconNumber "' RunState = '" ShortcutRunState "'"
						Try
							FileCreateShortcut ShortcutTarget, ShortcutLinkFile, ShortcutWorkingDir, ShortcutArgs, ShortcutDescription, ShortcutIconFile, ShortcutKey, ShortcutIconNumber, ShortcutRunState
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tFileCreateShortcut failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ WARN  ]`tFileCreateShortcut failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "FileCreateShortcut failed with Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileCreateShortcut successfully executed.`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "FileCreateShortcut successfully executed."
							}
						}
					If ScriptItemName = "Run"
						{
						Target := ""
						WorkingDir := ""
						Options := ""
						NoWait := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Target := ScriptItemString[2]
							If A_Index = 3
								WorkingDir := ScriptItemString[3]
							If A_Index = 4
								Options := ScriptItemString[4]
							If A_Index = 5
								NoWait := ScriptItemString[5]
							If A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						Target := TransForm(Target)
						If WorkingDir
							WorkingDir := TransForm(WorkingDir)
						Else
							{
							If !UserScript
								WorkingDir := A_ScriptDir
							}
						Options := TransForm(Options)
						NoWait := TransForm(NoWait)
						UserDescription := TransForm(UserDescription)
						If NoWait = 0
							{
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tRun and wait for Target: " Target "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
								}
							Else
								If UserScript
									If UserDescription
										ShowInfo.Value := UserDescription
									Else
										ShowInfo.Value := "Run and wait for Target: " Target "`r`nWorking directory: " WorkingDir
							If UserScript
								SetTimer RunWaitProgress, 50
							Try
								ReturnCode := RunWait(Target, WorkingDir, Options, &PID)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										SetTimer RunWaitProgress, 0
										FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "Run failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully stopped with return code: " ReturnCode "`r`n" , LogFile
								Else
									If UserScript
										{
										SetTimer RunWaitProgress, 0
										If !UserDescription
											ShowInfo.Value := "Run target with Process Id '" PID "' successfully stopped with return code: " ReturnCode
										}
								Try
									EnvSet "ReturnCode", ReturnCode
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , LogFile
									Else
										If UserScript
											{
											FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
											ShowInfo.Value := "Environment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message
											Sleep 10000
											}
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for 'ReturnCode' with value: " ReturnCode "`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Environment variable successfully set for 'ReturnCode' with value: " ReturnCode
									}
								}
							}
						Else
							{
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tRun Target: " Target "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
								}
							Else
								If UserScript
									If UserDescription
										ShowInfo.Value := UserDescription
									Else
										ShowInfo.Value := "Run Target: " Target "`r`nWorking directory: " WorkingDir
							Try
								ReturnCode := Run(Target, WorkingDir, Options, &PID)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := "Run failed with Error: " err.Message
										Sleep 10000
										}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully started with return code: " ReturnCode "`r`n" , LogFile
								Else
									If UserScript
										If !UserDescription
											ShowInfo.Value := "Run target with Process Id '" PID "' successfully started with return code: " ReturnCode
								Try
									EnvSet "ReturnCode", ReturnCode
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , LogFile
									Else
										If UserScript
											{
											FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
											ShowInfo.Value := "Environment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message
											Sleep 10000
											}
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for 'ReturnCode' with value: " ReturnCode "`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "Environment variable successfully set for 'ReturnCode' with value: " ReturnCode
									}
								}
							}
						}
					If ScriptItemName = "StringReplace"
						{
						Source := ""
						FindText := ""
						ReplaceText := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Source := ScriptItemString[2]
							If A_Index = 3
								FindText := ScriptItemString[3]
							If A_Index = 4
								ReplaceText := ScriptItemString[4]
							If A_Index = 5
								Destination := ScriptItemString[5]
							If A_Index = 6
								OverWrite := ScriptItemString[6]
							If A_Index = 7
								UserDescription := ScriptItemString[7]
							}
						Source := TransForm(Source)
						FindText := TransForm(FindText)
						ReplaceText := TransForm(ReplaceText)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tStringReplace text '" FindText "' with '" ReplaceText "' using source file '" Source "' for destination file: '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "StringReplace text '" FindText "' with '" ReplaceText "' using source file '" Source "' for destination file: '" Destination "' (OverWrite=" OverWrite ")"
						If FileExist(Source)
							{
							If Source = Destination
								{
								Contents := FileRead(Source)
								Contents := StrReplace(Contents, FindText, ReplaceText)
								FileDelete Source
								Try
									FileAppend Contents, Source
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , LogFile
									Else
										If UserScript
											{
											FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
											ShowInfo.Value := "StringReplace failed with Error: " err.Message
											Sleep 10000
											}
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile
									Else
										If UserScript
											If !UserDescription
												ShowInfo.Value := "StringReplace successfully executed."
									}
								Contents := ""
								}
							Else
								{
								If OverWrite = 1
									{
									If FileExist(Destination)
										FileDelete Destination
									Contents := FileRead(Source)
									Contents := StrReplace(Contents, FindText, ReplaceText)
									Try
										FileAppend Contents, Destination
									Catch as err
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , LogFile
										Else
											If UserScript
												{
												FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
												ShowInfo.Value := "StringReplace failed with Error: " err.Message
												Sleep 10000
												}
										}
									Else
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile
										Else
											If UserScript
												If !UserDescription
													ShowInfo.Value := "StringReplace successfully executed."
										}
									Contents := ""
									}
								Else
									{
									If Not FileExist(Destination)
										{
										Contents := FileRead(Source)
										Contents := StrReplace(Contents, FindText, ReplaceText)
										Try
											FileAppend Contents, Destination	
										Catch as err
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , LogFile
											Else
												If UserScript
													{
													FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
													ShowInfo.Value := "StringReplace failed with Error: " err.Message
													Sleep 10000
													}
											}
										Else
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile
											Else
												If UserScript
													If !UserDescription
														ShowInfo.Value := "StringReplace successfully executed."
											}
										Contents := ""
										}
									Else
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , LogFile
										Else
											If UserScript
												ShowInfo.Value := "'" Destination "' already exists."
										}
									}
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}	
					If ScriptItemName = "FileAppend"
						{
						Text := ""
						FileName := ""
						Encoding := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								Text := ScriptItemString[2]
							If A_Index = 3
								FileName := ScriptItemString[3]
							If A_Index = 4
								Encoding := ScriptItemString[4]
							If A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Text := TransForm(Text)
						FileName := TransForm(FileName)
						Encoding := TransForm(Encoding)
						UserDescription := TransForm(UserDescription)
						If Encoding = ""
							Encoding := A_FileEncoding
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tFileAppend write text '" Text "' to file: " FileName " (" Encoding ")`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "FileAppend write text '" Text "' to file: " FileName " (" Encoding ")"
						Text := Text "`r`n"
						Try	
							FileAppend Text, FileName, "`r`n " Encoding
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`tFileAppend failed with Error: " err.Message "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`tFileAppend failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "FileAppend failed with Error: " err.Message
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileAppend successfully executed.`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "FileAppend successfully executed."
							}
						}
					If ScriptItemName = "Download"
						{
						URL := ""
						FileName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								URL := ScriptItemString[2]
							If A_Index = 3
								FileName := ScriptItemString[3]
							If A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						URL := Trim(URL)
						FileName := TransForm(FileName)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tDownload '" URL "' for file: " FileName "`r`n" , LogFile
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "Download '" URL "' for file: " FileName
						Download URL, FileName
						If Not FileExist(FileName)
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`tDownload failed with Error: " A_LastError "`r`n" , LogFile
							Else
								If UserScript
									{
									FileAppend A_Now "`t[ ERROR ]`tDownload failed with Error: " A_LastError "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := "Download failed with Error: " A_LastError
									Sleep 10000
									}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tDownload successfully executed.`r`n" , LogFile
							Else
								If UserScript
									If !UserDescription
										ShowInfo.Value := "Download successfully executed."
							}
						}
					If ScriptItemName = "Powershell"
						{
						psCommand := ""
						WorkingDir := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								psCommand := ScriptItemString[2]
							If A_Index = 3
								WorkingDir := ScriptItemString[3]
							If A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						psCommand := TransForm(psCommand)
						If WorkingDir
							WorkingDir := TransForm(WorkingDir)
						UserDescription := TransForm(UserDescription)
						If FileExist(LogFile)
							{
							FileAppend A_Now "`t[ INFO  ]`tExecute Powershell command: " psCommand "`r`n" , LogFile
							FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
							}
						Else
							If UserScript
								If UserDescription
									ShowInfo.Value := UserDescription
								Else
									ShowInfo.Value := "Execute Powershell command: " psCommand
						SplitPath psCommand,,, &Extension
						If Extension = "ps1"
							{
							Try
								psCommand := FileRead(psCommand)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , LogFile
								Else
									If UserScript
										{
										FileAppend A_Now "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
										ShowInfo.Value := err.Message
										Sleep 10000
										}
								}
							}
						Else
							psCommand := StrReplace(psCommand, A_Space . "\n" . A_Space, "`r`n")
						If WorkingDir
							psCommand := "Set-Location -Path `"" WorkingDir "`"`r`n" psCommand
						psCommand := Format(psCommand)
						If UserScript
							{
							Try 
								ps := ComObject("psScript")
							Catch as err
								{
								FileAppend A_Now "`t[ ERROR ]`tpsScript error:`r`n" err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
								ShowInfo.Value := err.Message
								Sleep 10000
								}
							Else
								{
								Try
									psReturn := ps.PS_Script(psCommand)
								Catch as err
									{
									FileAppend A_Now "`t[ ERROR ]`tPowershell command stopped with error:`r`n" err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									ShowInfo.Value := err.Message
									Sleep 10000
									}
								Else
									{
									FileAppend A_Now "`t[ INFO  ]`tPowershell command successfully stopped with return message:`r`n" psReturn "`r`n" , A_AppData . "\Intune Win32 Launcher error.log"
									If !UserDescription
										ShowInfo.Value := "Powershell command successfully executed."
									}
								}
							}
						Else
							{
							Try 
								psReturn := PsScriptManager.ComInstance.PS_Script(psCommand)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tPowershell command stopped with error:`r`n" err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tPowershell command successfully stopped with return message:`r`n" psReturn "`r`n" , LogFile
								}
							}
						}
					If ScriptItemName = "CreateTask"
						{
						If !UserScript
							{
							TaskName := ""
							Description := ""
							Path := ""
							Arguments := ""
							ExecutionTimeLimit := 12
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									TaskName := ScriptItemString[2]
								If A_Index = 3
									Description := ScriptItemString[3]
								If A_Index = 4
									Path := ScriptItemString[4]
								If A_Index = 5
									Arguments := ScriptItemString[5]
								If A_Index = 6
									ExecutionTimeLimit := ScriptItemString[6]
								}
							TaskName := TransForm(TaskName)
							Description := TransForm(Description)
							Path := TransForm(Path)
							Arguments := TransForm(Arguments)
							ExecutionTimeLimit := Trim(ExecutionTimeLimit)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tCreateTask (system) for next logon with task name '" TaskName "' and description '" Description "' using executable path '" Path "' with the following arguments '" Arguments "'. Set limited runtime execution for " ExecutionTimeLimit " hours.`r`n" , LogFile
							ReturnMsg := CreateTask(TaskName, Description, Path, Arguments, ExecutionTimeLimit)
							If ReturnMsg = 0
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tCreateTask successfully executed.`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tCreateTask failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					If ScriptItemName = "DeleteTask"
						{
						If !UserScript
							{
							TaskName := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									TaskName := ScriptItemString[2]
								}
							TaskName := TransForm(TaskName)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tDeleteTask (system) with task name '" TaskName "'.`r`n" , LogFile
							ReturnMsg := DeleteTask(TaskName)
							If ReturnMsg = 0
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tDeleteTask successfully executed.`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tDeleteTask failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					If ScriptItemName = "Notify"
						{
						If !UserScript
							{
							MessageSection := ""
							ExitCode := 0
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									MessageSection := ScriptItemString[2]
								}
							MessageSection := TransForm(MessageSection)
							SectionType := IniRead(IniFile, MessageSection, "SectionType", "")
							SectionType := Trim(SectionType)
							If (SectionType = "Message")
								{
								If (cpauPID := ProcessExist(cpauPID))
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing Message notification section: " MessageSection " ]`r`n" , LogFile
									Send_NamedPipeMessage(MessageSection . ":" . ExitCode, SystemPipeName)
									ReturnCode := Receive_NamedPipeMessage(UserPipeName)
									If ReturnCode = 0
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
										}
									Else
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ ERROR ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
										}
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tNotification process not running. Skipping message notification section: " MessageSection "`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tMessage notification section not found: " MessageSection "`r`n" , LogFile
								}
							}
						}
					If ScriptItemName = "ExitApp"
						{
						If !UserScript
							{
							ExitCode := 0
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									ExitCode := ScriptItemString[2]
								}
							ExitCode := Trim(ExitCode)
							If (cpauPID := ProcessExist(cpauPID))
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tStopping notification process...`r`n" , LogFile
								Send_NamedPipeMessage("*:0", SystemPipeName)
								ReturnCode := Receive_NamedPipeMessage(UserPipeName)
								}
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`t" ProductName " stopped successfully. (Return Code = " ExitCode ")`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
							ExitApp ExitCode
							}
						}
					If ScriptItemName = "Reboot"
						{
						If !UserScript
							{
							ExitCode := 350
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									ExitCode := ScriptItemString[2]
								}
							ExitCode := Trim(ExitCode)
							If (cpauPID := ProcessExist(cpauPID))
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tStopping notification process...`r`n" , LogFile
								Send_NamedPipeMessage("*:0", SystemPipeName)
								ReturnCode := Receive_NamedPipeMessage(UserPipeName)
								}
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRun Target: " A_Comspec " /c Shutdown /r`r`n" , LogFile
							ReturnCode := Run(A_Comspec " /c Shutdown /r", "", "", &PID)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`t" ProductName " stopped successfully. (Return Code = " ExitCode ")`r`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
							ExitApp ExitCode
							}
						}
					If ScriptItemName = "FileAccess"
						{
						If !UserScript
							{
							FileName := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									FileName := ScriptItemString[2]
								}
							FileName := TransForm(FileName)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileAccess (everyone read and execute permission) for filename: " FileName "`r`n" , LogFile
							If FileExist(FileName)
								{
								Try
									ReturnValue := SetSecurityFile(FileName, "S-1-1-0", "1179817", "1")
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tFileAccess failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tFileAccess successfully executed. Return value: " ReturnValue "`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`t'" FileName "' not found.`r`n" , LogFile
								}
							}
						}
					Counter++
					}
				Else
					Break
				}
			}
		If A_LoopReadLine = "[" Section "]"
			Found := 1
		}
	Return Counter
	}
