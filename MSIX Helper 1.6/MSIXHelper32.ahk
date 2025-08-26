;@Ahk2Exe-SetName MSIX Helper 32-bit
;@Ahk2Exe-SetOrigFilename MSIXHelper32.exe
;@Ahk2Exe-SetDescription MSIX Helper 1.6
;@Ahk2Exe-SetVersion 1.6.0.0
;@Ahk2Exe-SetCompanyName Provolve B.V.
;@Ahk2Exe-SetCopyright Ferry van Gelderen
;@Ahk2Exe-SetMainIcon MainIcon.ico
;@Ahk2Exe-AddResource 160.ico, 160
;@Ahk2Exe-AddResource 206.ico, 206
;@Ahk2Exe-AddResource 207.ico, 207
;@Ahk2Exe-AddResource 208.ico, 208

#Requires AutoHotkey v2.0
#SingleInstance off
#NoTrayIcon

ProductName := "MSIX Helper 1.6"
AppUserModelId := GetCurrentApplicationUserModelId()
If AppUserModelId
	{
	A_PID := ProcessExist()
	for process in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_Process")
		{
		If process.ProcessId != 0
			{
			GetAppUserModelId := GetCurrentApplicationUserModelId(process.ProcessId)
			If GetAppUserModelId
				{
				If AppUserModelId = GetAppUserModelId
					{
					If process.ProcessId != A_PID
						ExitApp 89
					}
				}
			}
		}
	AppUserModelIdAppId := StrSplit(AppUserModelId, "!")
	PackageFamilyName := AppUserModelIdAppId[1]
	AppId := AppUserModelIdAppId[2]
	PackagePath := GetCurrentPackagePath()
	}
Else
	{
	ArgumentsNumber := A_Args.Length
	If ArgumentsNumber > 0
		{
		AppId := A_Args[1]
		PackageFamilyName := AppId
		PackagePath := A_ScriptDir
		A_Args[1] := ""
		}
	Else
		{
		AppId := ""
		PackageFamilyName := ""
		PackagePath := ""
		}	
	}
If AppId
	{
	EnvSet "PackageFamilyName", PackageFamilyName
	EnvSet "AppId", AppId
	EnvSet "PackagePath", PackagePath
	EnvSet "WorkingDir", A_ScriptDir
	EnvSet "StartMenuDir", A_StartMenu
	EnvSet "ProgramsDir", A_Programs
	EnvSet "DesktopDir", A_Desktop
	EnvSet "StartupDir", A_Startup
	EnvSet "DocumentsDir", A_MyDocuments
	FileLength := StrLen(A_ScriptName)
	FileLength := FileLength - 4
	FileName := SubStr(A_ScriptName, 1, FileLength)
	If FileExist(A_Temp "\" PackageFamilyName ".ini")
		IniFile := A_Temp "\" PackageFamilyName ".ini"
	Else
		IniFile := A_ScriptDir "\" FileName ".ini"
	LogFile := A_Temp "\" PackageFamilyName ".log"
	If FileExist(LogFile)
		{
		FileAppend A_Now "`t[ INFO  ]`t" ProductName " started for Application Id: " AppId "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tComputer name: " A_ComputerName "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tOperating System version: " A_OSVersion " (64-bit=" A_Is64bitOS ")`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tOperating System language code: " A_Language "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tUser name: " A_UserName "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tWorking directory: " A_ScriptDir "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tMSIX PackagePath: " PackagePath "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tMSIX PackageFamilyName: " PackageFamilyName "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tMSIX AppId: " AppId "`r`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tINI file: " IniFile "`r`n" , LogFile
		}
	Target := IniRead(IniFile, AppId, "Target", "")
	If Target
		{
		SetEnvSection := IniRead(IniFile, AppId, "SetEnv", "")
		If SetEnvSection
			{
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Setting Session Environment Variables for section: " SetEnvSection " ]`r`n" , LogFile				
			SetEnvCounter := SetEnv(SetEnvSection)
			}
		PreLaunchSection := IniRead(IniFile, AppId, "PreLaunch", "")
		If PreLaunchSection
			{
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Performing PreLaunch script section: " PreLaunchSection " ]`r`n" , LogFile
			StartGUI(AppId)
			ScriptSections := ""			
			PreLaunchCounter := Script(PreLaunchSection)
			SetTimer StopGUI, -1000
			}
		Target := TransForm(Target)
		WorkingDir := IniRead(IniFile, AppId, "WorkingDir", A_ScriptDir)
		If WorkingDir
			WorkingDir := TransForm(WorkingDir)
		Options := IniRead(IniFile, AppId, "Options", A_ScriptDir)
		If Options
			Options := TransForm(Options)
		Parameters := ""
		ArgumentsNumber := A_Args.Length
		If ArgumentsNumber > 0
			{
			Loop
				{
				Argument := IniRead(IniFile, AppId, "Arg" A_Index, "")
				If Argument = ""
					Break
				Else
					{
					Argument := TransForm(Argument)
					If InStr(Argument, A_Space)
						Argument := "`"" Argument "`""
					Parameters := Parameters . Argument . A_Space					
					}
				}
			for index, Argument in A_Args
				{
				If InStr(Argument, A_Space)
					Argument := "`"" Argument "`""
				Parameters := Parameters . Argument . A_Space
				}
			}
		Else
			{
			Loop
				{
				Argument := IniRead(IniFile, AppId, "Param" A_Index, "")
				If Argument = ""
					Break
				Else
					{
					Argument := TransForm(Argument)
					If InStr(Argument, A_Space)
						Argument := "`"" Argument "`""
					Parameters := Parameters . Argument . A_Space 
					}
				}
			}
		Parameters := Trim(Parameters)
		Target := Target . A_Space . Parameters
		If FileExist(LogFile)
			FileAppend A_Now "`t[ INFO  ]`t[ Executing target section: " AppId " ]`r`n" , LogFile		
		PostExitSection := IniRead(IniFile, AppId, "PostExit", "")
		If PostExitSection
			{
			If FileExist(LogFile)
				{
				FileAppend A_Now "`t[ INFO  ]`tRun and wait for Target: " Target "`r`n" , LogFile
				FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
				}
			Try
				ExitCode := RunWait(Target, WorkingDir, Options, &PID)
			Catch as err
				{
				If FileExist(LogFile)
					{
					FileAppend A_Now "`t[ ERROR ]`tError: " err.Message "`r`n" , LogFile
					FileAppend A_Now "`t[ WARN  ]`t" ProductName " stopped with error.`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
					}
				ExitApp 1
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tTarget with Process Id '" PID "' stopped with return code: " ExitCode "`r`n" , LogFile
				}
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Performing PostExit script section: " PostExitSection " ]`r`n" , LogFile
			StartGUI(AppId)
			ScriptSections := ""
			PostExitCounter := Script(PostExitSection)
			StopGUI()
			}
		Else
			{
			If FileExist(LogFile)
				{
				FileAppend A_Now "`t[ INFO  ]`tRun Target: " Target "`r`n" , LogFile
				FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
				}
			Try
				ExitCode := Run(Target, WorkingDir, Options, &PID)
			Catch as err
				{
				If FileExist(LogFile)
					{
					FileAppend A_Now "`t[ ERROR ]`tError: " err.Message "`r`n" , LogFile
					FileAppend A_Now "`t[ WARN  ]`t" ProductName " stopped with error.`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
					}
				ExitApp 1
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tTarget with Process Id '" PID "' started successfully.`r`n" , LogFile
				}
			}
		If FileExist(LogFile)
			FileAppend A_Now "`t[ INFO  ]`t" ProductName " stopped successfully.`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
		ExitApp 0	
		}
	Else
		{
		If FileExist(LogFile)
			{
			FileAppend A_Now "`t[ ERROR ]`tNo target found for the MSIX Aplication Id: " AppId "`r`n" , LogFile
			FileAppend A_Now "`t[ WARN  ]`tPlease make sure the MSIX Application Id exist as a section and has a Target key value in the INI file: " IniFile "`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
			}
		MsgBox "No target found for the MSIX Aplication Id: " AppId "`nPlease make sure the MSIX Application Id exist as a section and has a Target key value in the INI file:`n`n" IniFile, ProductName, 48
		ExitApp 1
		}
	}
Else
	{
	MsgBox "No MSIX Application Id found!`n`nPlease make sure this process is running from the MSIX package environment or add the Application Id section name that is present in the INI file as a startup parameter.", ProductName, 48
	ExitApp 15703
	}

; --------------------------------------------------------------- Functions for collecting the application user model ID and PackagePath for the current process or other process ID ---------------------------------------------------------------
GetCurrentApplicationUserModelId(PID := "")
	{
	ERROR_SUCCESS := 0
	ERROR_INSUFFICIENT_BUFFER := 122
	APPMODEL_ERROR_NO_APPLICATION := 15703
	sizeofwchar_t := 2
	If PID
		{
		PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
		hProc := DllCall("OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "UInt", 0, "Ptr", PID, "Ptr")
		RC := DllCall("GetApplicationUserModelId", "Uint", hProc, "UInt*", &Length := 0, "Ptr", 0)
		If (RC != ERROR_INSUFFICIENT_BUFFER)
			Return ""
		CurrentApplicationUserModelId := Buffer(Length * sizeofwchar_t)
		RC := DllCall("GetApplicationUserModelId", "Uint", hProc, "UInt*", &Length, "Ptr", CurrentApplicationUserModelId)
		If (RC != ERROR_SUCCESS)
			Return ""
		}
	Else
		{
		RC := DllCall("GetCurrentApplicationUserModelId", "UInt*", &Length := 0, "Ptr", 0)
		If (RC != ERROR_INSUFFICIENT_BUFFER)
			Return ""
		CurrentApplicationUserModelId := Buffer(Length * sizeofwchar_t)
		RC := DllCall("GetCurrentApplicationUserModelId", "UInt*", &Length, "Ptr", CurrentApplicationUserModelId)
		If (RC != ERROR_SUCCESS)
			Return ""
		}
	AppUserModelId := StrGet(CurrentApplicationUserModelId)
	DllCall("CloseHandle", "Ptr", CurrentApplicationUserModelId)
	Return AppUserModelId
	}

GetPackageFullName(PID := "")
	{
	ERROR_SUCCESS := 0
	ERROR_INSUFFICIENT_BUFFER := 122
	APPMODEL_ERROR_NO_PACKAGE := 15700
	sizeofwchar_t := 2
	If PID
		{
		PROCESS_QUERY_LIMITED_INFORMATION := 0x1000
		hProc := DllCall("OpenProcess", "UInt", PROCESS_QUERY_LIMITED_INFORMATION, "UInt", 0, "Ptr", PID, "Ptr")
		RC := DllCall("GetPackageFullName", "Uint", hProc, "UInt*", &Length := 0, "Ptr", 0)
		If (RC != ERROR_INSUFFICIENT_BUFFER)
			Return ""
		CurrentPackageFullName := Buffer(Length * sizeofwchar_t)
		RC := DllCall("GetPackageFullName", "Uint", hProc, "UInt*", &Length, "Ptr", CurrentPackageFullName)
		If (RC != ERROR_SUCCESS)
			Return ""
		}
	Else
		{
		RC := DllCall("GetCurrentPackageFullName", "UInt*", &Length := 0, "Ptr", 0)
		If (RC != ERROR_INSUFFICIENT_BUFFER)
			Return ""
		CurrentPackageFullName := Buffer(Length * sizeofwchar_t)
		RC := DllCall("GetCurrentPackageFullName", "UInt*", &Length, "Ptr", CurrentPackageFullName)
		If (RC != ERROR_SUCCESS)
			Return ""
		}
	PackageFullName := StrGet(CurrentPackageFullName)
	DllCall("CloseHandle", "Ptr", CurrentPackageFullName)
	Return PackageFullName
	}

GetCurrentPackagePath(PackageFullName := "")
	{
	ERROR_SUCCESS := 0
	ERROR_INSUFFICIENT_BUFFER := 122
	APPMODEL_ERROR_NO_PACKAGE := 15700
	sizeofwchar_t := 2
	If PackageFullName
		{
		RC := DllCall("GetPackagePathByFullName", "WStr", PackageFullName, "UInt*", &Length := 0, "Ptr", 0)
		If (RC != ERROR_INSUFFICIENT_BUFFER)
			Return ""
		CurrentPackagePath := Buffer(Length * sizeofwchar_t)
		RC := DllCall("GetPackagePathByFullName", "WStr", PackageFullName, "UInt*", &Length, "Ptr", CurrentPackagePath)
		If (RC != ERROR_SUCCESS)
			Return ""
		}
	Else
		{
		RC := DllCall("GetCurrentPackagePath", "UInt*", &Length := 0, "Ptr", 0)
		If (RC != ERROR_INSUFFICIENT_BUFFER)
			Return ""
		CurrentPackagePath := Buffer(Length * sizeofwchar_t)
		RC := DllCall("GetCurrentPackagePath", "UInt*", &Length, "Ptr", CurrentPackagePath)
		If (RC != ERROR_SUCCESS)
			Return ""
		}
	PackagePath := StrGet(CurrentPackagePath)
	DllCall("CloseHandle", "Ptr", CurrentPackagePath)
	Return PackagePath
	}

; --------------------------------------------------------------- Function for translating existing environment variables ---------------------------------------------------------------
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

; --------------------------------------------------------------- Function for setting environment variables ---------------------------------------------------------------
SetEnv(Section)
	{
	Global IniFile
	Global LogFile
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
					Try
						EnvSet EnvSetName, EnvSetValue
					Catch as err
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile							
						}
					Else
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" EnvSetName "' with value: " EnvSetValue "`r`n" , LogFile
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

; --------------------------------------------------------------- Function for showing and stopping the progress GUI during the Pre-Launch and Post-Exit script execution ---------------------------------------------------------------
StartGUI(AppId := "")
	{
	Global NotifyProgress
	Global MSIXHelperGUI
	CoordMode "Mouse", "Screen"
	MouseGetPos &xpos, &ypos
	BackColor := "2C2C2C"
	MSIXHelperGUI := Gui("+DpiScale -Caption +AlwaysOnTop +E0x08000000", AppId)
	MSIXHelperGUI.MarginX := 0
	MSIXHelperGUI.MarginY := 0
	MSIXHelperGUI.BackColor := BackColor
	If AppId
		{
		If FileExist(A_WorkingDir . "\Assets\" . AppId . "-Square44x44Logo.scale-400.png")
			MSIXHelperGUI.Add("Picture", "w50 h-1 BackgroundTrans", A_WorkingDir . "\Assets\" . AppId . "-Square44x44Logo.scale-400.png")
		Else
			MSIXHelperGUI.Add("Picture", "w50 h-1 BackgroundTrans Icon1", A_ScriptFullPath)
		}
	Else
		MSIXHelperGUI.Add("Picture", "w50 h-1 BackgroundTrans Icon1", A_ScriptFullPath)
	NotifyProgress := MSIXHelperGUI.Add("Progress", "xp y+m h5 w50 0x8")
	NotifyProgress.Opt("c3399FF")
	NotifyProgress.Opt("Background" . BackColor . "")
	MSIXHelperGUI.Show("X" . xpos . " Y" . ypos . "")
	WinSetTransColor BackColor, AppId
	SetTimer RunWaitProgress, 50
	}

StopGUI()
	{
	SetTimer RunWaitProgress, 0
	NotifyProgress.Opt("-0x8")
	NotifyProgress.Value := 100
	Sleep 1000
	MSIXHelperGUI.Destroy()
	}

RunWaitProgress(ProgressBar := 0)
	{
	NotifyProgress.Value := ProgressBar++
	If ProgressBar = 100
		ProgressBar := 0
	Sleep 100
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

; --------------------------------------------------------------- psScript Manager Class ---------------------------------------------------------------
class PsScriptManager
	{
	static RootKey        := "HKEY_CURRENT_USER\Software\Classes\"
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

; --------------------------------------------------------------- Function for executing PreLaunch and PostExit script items ---------------------------------------------------------------
Script(Section)
	{
	Global IniFile
	Global LogFile
	Global ScriptSections
	Found := 0
	Counter := 0
	If FileExist(IniFile)
		{
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
						ScriptItemString := StrSplit(A_LoopReadLine, ",")
						ScriptItemName := ScriptItemString[1]
						ScriptItemName := Trim(ScriptItemName)
						If (ScriptItemName = "IfEqual" or ScriptItemName = "IfNotEqual")
							{
							CheckVariable := ""
							CheckValue := ""
							GotoIfSection := ""
							GotoElseSection := ""
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
								}
							CheckVariable := Trim(CheckVariable)
							CheckValue := Trim(CheckValue)
							GotoIfSection := Trim(GotoIfSection)
							GotoElseSection := Trim(GotoElseSection)
							If ScriptItemName = "IfEqual"
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`r`n" , LogFile
									}
								GetValue := EnvGet(CheckVariable)
								If CheckValue = GetValue
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
									StatementCounter := Script(GotoIfSection)
									}
								Else
									{
									If GotoElseSection
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
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
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
									}
								GetValue := EnvGet(CheckVariable)
								If CheckValue != GetValue
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
									StatementCounter := Script(GotoIfSection)
									}
								Else
									{
									If GotoElseSection
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
										StatementCounter := Script(GotoElseSection)
										}
									}
								}
							}
						If (ScriptItemName = "IfExist" or ScriptItemName = "IfNotExist")
							{
							CheckFileFolderRegKey := ""
							GotoIfSection := ""
							GotoElseSection := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									CheckFileFolderRegKey := ScriptItemString[2]
								If A_Index = 3
									GotoIfSection := ScriptItemString[3]
								If A_Index = 4
									GotoElseSection := ScriptItemString[4]
								}
							CheckFileFolderRegKey := TransForm(CheckFileFolderRegKey)
							GotoIfSection := Trim(GotoIfSection)
							GotoElseSection := Trim(GotoElseSection)
							If ScriptItemName = "IfExist"
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
									}
								Path_Array := StrSplit(CheckFileFolderRegKey, "\",)
								If (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
									{
									CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, Path_Array[1] . "\", "")
									If RegKeyExists(Path_Array[1], CheckFileFolderRegKey) = 1
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
										StatementCounter := Script(GotoIfSection)
										}
									Else
										{
										If GotoElseSection
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
											StatementCounter := Script(GotoElseSection)
											}
										}
									}
								Else
									{
									If FileExist(CheckFileFolderRegKey)
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
										StatementCounter := Script(GotoIfSection)
										}
									Else
										{
										If GotoElseSection
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
											StatementCounter := Script(GotoElseSection)
											}
										}
									}
								}
							If ScriptItemName = "IfNotExist"
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else goto script section '" GotoElseSection "'.`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tUsing IfNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoIfSection "'. Else, continue...`r`n" , LogFile
									}
								Path_Array := StrSplit(CheckFileFolderRegKey, "\",)
								If (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
									{
									CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, Path_Array[1] . "\", "")
									If RegKeyExists(Path_Array[1], CheckFileFolderRegKey) = 0
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
										StatementCounter := Script(GotoIfSection)
										}
									Else
										{
										If GotoElseSection
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
											StatementCounter := Script(GotoElseSection)
											}
										}
									}
								Else
									{
									If Not FileExist(CheckFileFolderRegKey)
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`r`n" , LogFile
										StatementCounter := Script(GotoIfSection)
										}
									Else
										{
										If GotoElseSection
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`r`n" , LogFile
											StatementCounter := Script(GotoElseSection)
											}
										}
									}							
								}
							}
						If ScriptItemName = "FileCopy"
							{
							Source := ""
							Destination := ""
							OverWrite := 0
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									Source := ScriptItemString[2]
								If A_Index = 3
									Destination := ScriptItemString[3]
								If A_Index = 4
									OverWrite := ScriptItemString[4]
								}	
							Source := TransForm(Source)
							Destination := TransForm(Destination)
							OverWrite := TransForm(OverWrite)	
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile						
							If FileExist(Source)
								{
								Try
									FileCopy Source, Destination, OverWrite
								Catch as Err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tFileCopy failed for " err.Extra " files.`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
								}
							}
						If ScriptItemName = "FolderCopy"
							{
							Source := ""
							Destination := ""
							OverWrite := 0
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									Source := ScriptItemString[2]
								If A_Index = 3
									Destination := ScriptItemString[3]
								If A_Index = 4
									OverWrite := ScriptItemString[4]
								}
							Source := TransForm(Source)
							Destination := TransForm(Destination)
							OverWrite := TransForm(OverWrite)	
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFolderCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile						
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
											FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`r`n" , LogFile
										}
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
								}
							}
						If ScriptItemName = "FolderCreate"
							{
							Destination := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									Destination := ScriptItemString[2]
								}							
							Destination := TransForm(Destination)
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFolderCreate: " Destination "`r`n" , LogFile
							If Not FileExist(Destination)
								{
								Try						
									DirCreate Destination
								Catch as Err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tFolderCreate failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tFolderCreate successfully executed.`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , LogFile
								}
							}	
						If ScriptItemName = "FileDelete"
							{
							FilePattern := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									FilePattern := ScriptItemString[2]
								}
							FilePattern := TransForm(FilePattern)												
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileDelete: " FilePattern "`r`n" , LogFile						
							If FileExist(FilePattern)
								{
								Try
									FileDelete FilePattern
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tFileDelete failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tFileDelete successfully executed.`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`t'" FilePattern "' not found.`r`n" , LogFile
								}
							}						
						If ScriptItemName = "FolderDelete"
							{
							DirName := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									DirName := ScriptItemString[2]
								}
							DirName := TransForm(DirName)													
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFolderDelete: " DirName "`r`n" , LogFile
							If DirExist(DirName)
								{
								Try
									DirDelete DirName, 1
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tFolderDelete failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tFolderDelete successfully executed.`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`t'" DirName "' not found.`r`n" , LogFile
								}
							}
						If ScriptItemName = "EnvSet"
							{
							EnvSetName := ""
							EnvSetValue := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									EnvSetName := ScriptItemString[2]
								If A_Index = 3
									EnvSetValue := ScriptItemString[3]
								}
							EnvSetName := Trim(EnvSetName)
							EnvSetValue := TransForm(EnvSetValue)							
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvSet setting session environment variable name '" EnvSetName "' with value '" EnvSetValue "'`r`n" , LogFile
							Try
								EnvSet EnvSetName, EnvSetValue
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tSession environment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile							
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tSession environment variable successfully set for '" EnvSetName "' with value: " EnvSetValue "`r`n" , LogFile
								}
							}
						If ScriptItemName = "RegRead"
							{
							KeyName := ""
							ValueName := ""
							DefaultValue := ""
							EnvSetValue := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									KeyName := ScriptItemString[2]
								If A_Index = 3
									ValueName := ScriptItemString[3]
								If A_Index = 4
									DefaultValue := ScriptItemString[4]
								}
							KeyName := TransForm(KeyName)
							ValueName := Trim(ValueName)
							DefaultValue := TransForm(DefaultValue)							
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegRead registry item '" ValueName "' in '" KeyName "' with default value '" DefaultValue "'`r`n" , LogFile
							Try
								EnvSetValue := RegRead(KeyName, ValueName, DefaultValue)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tRegRead failed with Error: " err.Message "`r`n" , LogFile
								EnvSetValue := DefaultValue
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegRead successfully executed with value: " EnvSetValue "`r`n" , LogFile
								}
							Try
								EnvSet ValueName, EnvSetValue
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tSession environment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile							
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tSession environment variable successfully set for '" ValueName "' with value: " EnvSetValue "`r`n" , LogFile
								}
							}
						If ScriptItemName = "RegWrite"
							{
							ValueType := ""
							KeyName := ""
							ValueName := ""
							Value := ""
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
								}
							ValueType := TransForm(ValueType)
							KeyName := TransForm(KeyName)
							ValueName := TransForm(ValueName)
							Value := TransForm(Value)							
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegWrite create '" ValueType "' registry item '" ValueName "' in '" KeyName "' with value '" Value "'`r`n" , LogFile
							Try
								RegWrite Value, ValueType, KeyName, ValueName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tRegWrite failed with Error: " err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegWrite successfully executed.`r`n" , LogFile
								}
							}	
						If ScriptItemName = "RegDelete"
							{
							KeyName := ""
							ValueName := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									KeyName := ScriptItemString[2]
								If A_Index = 3
									ValueName := ScriptItemString[3]
								}
							KeyName := TransForm(KeyName)
							ValueName := TransForm(ValueName)				
							If ValueName
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegDelete registry value '" ValueName "' from '" KeyName "'`r`n" , LogFile	
								Try							
									RegDelete KeyName, ValueName
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tRegDelete successfully executed.`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegDelete registry key: " KeyName "`r`n" , LogFile
								Try
									RegDeleteKey KeyName
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tRegDelete successfully executed.`r`n" , LogFile
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
								}
							Filename := TransForm(Filename)
							Section := Trim(Section)
							KeyName := Trim(KeyName)
							DefaultValue := TransForm(DefaultValue)							
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniRead from file name '" Filename "' using section name '" Section "' for key name '" KeyName "' with default value '" DefaultValue "'`r`n" , LogFile
							Try
								IniValue := IniRead(Filename, Section, KeyName, DefaultValue)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tIniRead failed with Error: " err.Message "`r`n" , LogFile
								IniValue := DefaultValue
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tIniRead successfully executed with value: " IniValue "`r`n" , LogFile
								}
							Try
								EnvSet KeyName, IniValue
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tSession environment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message "`r`n" , LogFile							
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tSession environment variable successfully set for '" KeyName "' with value: " IniValue "`r`n" , LogFile
								}
							}
						If ScriptItemName = "IniWrite"
							{
							Value := ""
							Filename := ""
							Section := ""
							KeyName := ""
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
								}
							Value := TransForm(Value)
							Filename := TransForm(Filename)
							Section := Trim(Section)
							KeyName := Trim(KeyName)							
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniWrite value '" Value "' in file name '" Filename "' for section name '" Section "' with key name '" KeyName "'`r`n" , LogFile
							Try
								IniWrite Value, Filename, Section, KeyName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tIniWrite failed with Error: " err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tIniWrite successfully executed.`r`n" , LogFile
								}
							}	
						If ScriptItemName = "IniDelete"
							{
							Filename := ""
							Section := ""
							KeyName := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									Filename := ScriptItemString[2]
								If A_Index = 3
									Section := ScriptItemString[3]
								If A_Index = 4
									KeyName := ScriptItemString[4]
								}
							Filename := TransForm(Filename)
							Section := Trim(Section)
							KeyName := Trim(KeyName)				
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniDelete using file name '" Filename "' for section name '" Section "' (with key name '" KeyName "')`r`n" , LogFile	
							Try							
								IniDelete Filename, Section, KeyName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tIniDelete failed with Error: " err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tIniDelete successfully executed.`r`n" , LogFile
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
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileCreateShortcut using the following properties: Target = '" ShortcutTarget "' LinkFile = '" ShortcutLinkFile "' WorkingDir = '" ShortcutWorkingDir "' Args = '" ShortcutArgs "' Description = '" ShortcutDescription "' IconFile = '" ShortcutIconFile "' ShortcutKey = '" ShortcutKey "' IconNumber = '" ShortcutIconNumber "' RunState = '" ShortcutRunState "'`r`n" , LogFile
							Try
								FileCreateShortcut ShortcutTarget, ShortcutLinkFile, ShortcutWorkingDir, ShortcutArgs, ShortcutDescription, ShortcutIconFile, ShortcutKey, ShortcutIconNumber, ShortcutRunState
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFileCreateShortcut failed with Error: " err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileCreateShortcut successfully executed.`r`n" , LogFile
								}
							}
						If ScriptItemName = "Run"
							{
							Target := ""
							WorkingDir := ""
							Options := ""
							NoWait := 0
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
								}
							Target := TransForm(Target)
							WorkingDir := TransForm(WorkingDir)
							Options := TransForm(Options)
							NoWait := TransForm(NoWait)															
							If NoWait = 0
								{
								If FileExist(LogFile)
									{
									FileAppend A_Now "`t[ INFO  ]`tRun and wait for Target: " Target "`r`n" , LogFile
									FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile									
									}
								Try
									ExitCode := RunWait(Target, WorkingDir, Options, &PID)
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully stopped with return code: " ExitCode "`r`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									{
									FileAppend A_Now "`t[ INFO  ]`tRun Target: " Target "`r`n" , LogFile
									FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
									}
								Try							
									ExitCode := Run(Target, WorkingDir, Options, &PID)
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , LogFile
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully started with return code: " ExitCode "`r`n" , LogFile
									}
								}
							}
						If ScriptItemName = "Powershell"
							{
							psCommand := ""
							WorkingDir := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									psCommand := ScriptItemString[2]
								If A_Index = 3
									WorkingDir := ScriptItemString[3]
								}
							psCommand := TransForm(psCommand)
							If WorkingDir
								WorkingDir := TransForm(WorkingDir)
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tExecute Powershell command: " psCommand "`r`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
								}
							SplitPath psCommand,,, &Extension
							If Extension = "ps1"
								{
								Try
									psCommand := FileRead(psCommand)
								Catch as err
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , LogFile
									}
								}
							Else
								psCommand := StrReplace(psCommand, A_Space . "\n" . A_Space, "`r`n")
							If WorkingDir
								psCommand := "Set-Location -Path `"" WorkingDir "`"`r`n" psCommand
							psCommand := Format(psCommand)
							Result := PsScriptManager.Register()
							If !Result = 0
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tPowershell COM registration failed with Error: " Result "`r`n" , LogFile	
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
						If ScriptItemName = "StringReplace"
							{
							Source := ""
							FindText := ""
							ReplaceText := ""
							Destination := ""
							OverWrite := 0
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
								}
							Source := TransForm(Source)
							FindText := TransForm(FindText)
							ReplaceText := TransForm(ReplaceText)
							Destination := TransForm(Destination)
							OverWrite := TransForm(OverWrite)					
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tStringReplace text '" FindText "' with '" ReplaceText "' using source file '" Source "' for destination file: '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
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
										}
									Else
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile										
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
											}
										Else
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile										
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
												}
											Else
												{
												If FileExist(LogFile)
													FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile									
												}
											Contents := ""
											}
										Else
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , LogFile										
											}
										}									
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile								
								}
							}	
						If ScriptItemName = "FileAppend"
							{
							Text := ""
							FileName := ""
							Encoding := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									Text := ScriptItemString[2]
								If A_Index = 3
									FileName := ScriptItemString[3]
								If A_Index = 4
									Encoding := ScriptItemString[4]
								}
							Text := TransForm(Text)
							FileName := TransForm(FileName)
							Encoding := TransForm(Encoding)
							If Encoding = ""
								Encoding := A_FileEncoding
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileAppend write text '" Text "' to file: " FileName " (" Encoding ")`r`n" , LogFile
							Text := Text "`r`n"
							Try					
								FileAppend Text, FileName, "`n " Encoding
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tFileAppend failed with Error: " err.Message "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileAppend successfully executed.`r`n" , LogFile
								}
							}	
						If ScriptItemName = "Download"
							{
							URL := ""
							FileName := ""
							Loop ScriptItemString.Length
								{
								If A_Index = 2
									URL := ScriptItemString[2]
								If A_Index = 3
									FileName := ScriptItemString[3]
								}
							URL := Trim(URL)
							FileName := TransForm(FileName)								
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tDownload '" URL "' for file: " FileName "`r`n" , LogFile
							Download URL, FileName
							If Not FileExist(FileName)
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tDownload failed with Error: " A_LastError "`r`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tDownload successfully executed.`r`n" , LogFile
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
		}
	Else
		{
		If FileExist(LogFile)
			FileAppend A_Now "`t[ ERROR ]`tConfiguration file not found: " IniFile "`r`n" , LogFile		
		}
	Return Counter
	}
