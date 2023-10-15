;@Ahk2Exe-SetName MSIX Helper 32-bit
;@Ahk2Exe-SetOrigFilename MSIXHelper32.exe
;@Ahk2Exe-SetDescription MSIX Helper 1.2
;@Ahk2Exe-SetVersion 1.2.0.0
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

ProductName := "MSIX Helper 1.2"
AppUserModelId := GetCurrentApplicationUserModelId()
If AppUserModelId
	{
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
		FileAppend A_Now "`t[ INFO  ]`t" ProductName " started for Application Id: " AppId "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tComputer name: " A_ComputerName "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tOperating System version: " A_OSVersion " (64-bit=" A_Is64bitOS ")`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tOperating System language code: " A_Language "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tUser name: " A_UserName "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tWorking directory: " A_ScriptDir "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tMSIX PackagePath: " PackagePath "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tMSIX PackageFamilyName: " PackageFamilyName "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tMSIX AppId: " AppId "`n" , LogFile
		FileAppend A_Now "`t[ INFO  ]`tINI file: " IniFile "`n" , LogFile
		}
	Target := IniRead(IniFile, AppId, "Target", "")
	If Target
		{
		SetEnvSection := IniRead(IniFile, AppId, "SetEnv", "")
		If SetEnvSection
			{
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Setting Environment Variables for section: " SetEnvSection " ]`n" , LogFile				
			SetEnvCounter := SetEnv(SetEnvSection)
			}
		PreLaunchSection := IniRead(IniFile, AppId, "PreLaunch", "")
		If PreLaunchSection
			{
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Performing PreLaunch script section: " PreLaunchSection " ]`n" , LogFile
			ScriptSections := ""			
			PreLaunchCounter := Script(PreLaunchSection)
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
			FileAppend A_Now "`t[ INFO  ]`t[ Executing target section: " AppId " ]`n" , LogFile		
		PostExitSection := IniRead(IniFile, AppId, "PostExit", "")
		If PostExitSection
			{
			If FileExist(LogFile)
				{
				FileAppend A_Now "`t[ INFO  ]`tRun and wait for Target: " Target "`n" , LogFile
				FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`n" , LogFile
				}
			Try
				ExitCode := RunWait(Target, WorkingDir, Options, &PID)
			Catch as err
				{
				If FileExist(LogFile)
					{
					FileAppend A_Now "`t[ ERROR ]`tError: " err.Message "`n" , LogFile
					FileAppend A_Now "`t[ WARN  ]`t" ProductName " stopped with error.`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`n" , LogFile
					}
				ExitApp 1
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tTarget with Process Id '" PID "' stopped with return code: " ExitCode "`n" , LogFile
				}
			If FileExist(LogFile)
				FileAppend A_Now "`t[ INFO  ]`t[ Performing PostExit script section: " PostExitSection " ]`n" , LogFile
			ScriptSections := ""
			PostExitCounter := Script(PostExitSection)
			}
		Else
			{
			If FileExist(LogFile)
				{
				FileAppend A_Now "`t[ INFO  ]`tRun Target: " Target "`n" , LogFile
				FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`n" , LogFile
				}
			Try
				ExitCode := Run(Target, WorkingDir, Options, &PID)
			Catch as err
				{
				If FileExist(LogFile)
					{
					FileAppend A_Now "`t[ ERROR ]`tError: " err.Message "`n" , LogFile
					FileAppend A_Now "`t[ WARN  ]`t" ProductName " stopped with error.`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`n" , LogFile
					}
				ExitApp 1
				}
			Else
				{
				If FileExist(LogFile)
					FileAppend A_Now "`t[ INFO  ]`tTarget with Process Id '" PID "' started successfully.`n" , LogFile
				}
			}
		If FileExist(LogFile)
			FileAppend A_Now "`t[ INFO  ]`t" ProductName " stopped successfully.`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`n" , LogFile
		ExitApp 0	
		}
	Else
		{
		If FileExist(LogFile)
			{
			FileAppend A_Now "`t[ ERROR ]`tNo target found for the MSIX Aplication Id: " AppId "`n" , LogFile
			FileAppend A_Now "`t[ WARN  ]`tPlease make sure the MSIX Application Id exist as a section and has a Target key value in the INI file: " IniFile "`n--------------`t---------`t----------------------------------------------------------------------------------------------------------------------------`n" , LogFile
			}
		MsgBox "No target found for the MSIX Aplication Id: " AppId "`nPlease make sure the MSIX Application Id exist as a section and has a Target key value in the INI file:`n`n" IniFile, ProductName, 48
		ExitApp 1
		}
	}
Else
	{
	MsgBox "No MSIX Application Id found!`n`nPlease make sure this process is running from the MSIX package environment.", ProductName, 48
	ExitApp 15703
	}

; --------------------------------------------------------------- Functions for collecting the application user model ID and PackagePath for the current process ---------------------------------------------------------------
GetCurrentApplicationUserModelId()
	{
	ERROR_SUCCESS := 0
	ERROR_INSUFFICIENT_BUFFER := 122
	APPMODEL_ERROR_NO_APPLICATION := 15703
	sizeofwchar_t := 2
	RC := DllCall("GetCurrentApplicationUserModelId", "UInt*", &Length := 0, "Ptr", 0)
	If (RC != ERROR_INSUFFICIENT_BUFFER)
		{
		AppUserModelId := ""
		Return AppUserModelId
		}
	CurrentApplicationUserModelId := Buffer(Length * sizeofwchar_t)
	RC := DllCall("GetCurrentApplicationUserModelId", "UInt*", &Length, "Ptr", CurrentApplicationUserModelId)
	If (RC != ERROR_SUCCESS)
		{
		AppUserModelId := ""
		Return AppUserModelId
		}
	AppUserModelId := StrGet(CurrentApplicationUserModelId)
	DllCall("CloseHandle", "Ptr", CurrentApplicationUserModelId)
	Return AppUserModelId
	}
GetCurrentPackagePath()
	{
	ERROR_SUCCESS := 0
	ERROR_INSUFFICIENT_BUFFER := 122
	APPMODEL_ERROR_NO_PACKAGE := 15700
	sizeofwchar_t := 2
	RC := DllCall("GetCurrentPackagePath", "UInt*", &Length := 0, "Ptr", 0)
	If (RC != ERROR_INSUFFICIENT_BUFFER)
		{
		PackagePath := ""
		Return PackagePath
		}
	CurrentPackagePath := Buffer(Length * sizeofwchar_t)
	RC := DllCall("GetCurrentPackagePath", "UInt*", &Length, "Ptr", CurrentPackagePath)
	If (RC != ERROR_SUCCESS)
		{
		PackagePath := ""
		Return PackagePath
		}
	PackagePath := StrGet(CurrentPackagePath)
	DllCall("CloseHandle", "Ptr", CurrentPackagePath)
	Return PackagePath
	}

; --------------------------------------------------------------- Function for translating existing environment variables ---------------------------------------------------------------
TransForm(String)
	{
	String := Trim(String)
	If InStr(String, "%")
		{
		StringLength := StrLen(String)
		StartPos := 1	
		FoundPair := 0
		FirstPos := 0
		SecondPos := 0
		OutputValue := ""
		Loop 
			{
			StringAdd := ""
			FoundPos := InStr(String, "%",,, A_Index)
			If FoundPos = 0
				{
				StartPos := SecondPos + 1
				StringLength++
				FoundLength := StringLength - StartPos
				StringAdd := SubStr(String, StartPos, FoundLength)
				OutputValue := OutputValue . StringAdd
				Break
				}
			Else
				{
				If FoundPair = 1
					{
					StartPos := SecondPos + 1
					FoundPair := 0
					FirstPos := 0
					SecondPos := 0
					}
				If FirstPos = 0
					FirstPos := FoundPos
				Else
					{
					If SecondPos = 0
						{
						FoundPair := 1
						SecondPos := FoundPos
						If StartPos != FirstPos
							{
							FoundLength := FirstPos - StartPos
							StringAdd := SubStr(String, StartPos, FoundLength)
							}
						FirstPos++	
						FoundLength := SecondPos - FirstPos
						EnvVarName := SubStr(String, FirstPos, FoundLength)
						EnvVarValue := EnvGet(EnvVarName)
						If EnvVarValue		
							OutputValue := OutputValue . StringAdd . EnvVarValue
						Else
							OutputValue := OutputValue . StringAdd . "%" EnvVarName "%"
						}
					}
				}
			}
		Return OutputValue
		}
	Else
		Return String
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
							FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`n" , LogFile							
						}
					Else
						{
						If FileExist(LogFile)
							FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" EnvSetName "' with value: " EnvSetValue "`n" , LogFile
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

; --------------------------------------------------------------- Function for executing PreLaunch and PostExit script items ---------------------------------------------------------------
Script(Section)
	{
	Global IniFile
	Global LogFile
	Global ScriptSections
	Found := 0
	Counter := 0
	If Instr(ScriptSections, Section "`n")
		{
		If FileExist(LogFile)
			FileAppend A_Now "`t[ WARN  ]`tScript section " Section " is already executed. Each section can be executed only once for each PreLaunch and PostExit script section. Please check the if/else statements.`n" , LogFile
		Return Counter
		}
	ScriptSections := ScriptSections . Section "`n"
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
					If (ScriptItemName = "If" or ScriptItemName = "IfNot")
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
						If ScriptItemName = "If"
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tUsing If statement for variable '" CheckVariable "' with value '" CheckValue "'. When true goto script section: " GotoIfSection " (Else goto script section: " GotoElseSection ")`n" , LogFile
							GetValue := EnvGet(CheckVariable)
							If CheckValue = GetValue
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`n" , LogFile
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`n" , LogFile
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNot"
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tUsing IfNot statement for variable '" CheckVariable "' with value '" CheckValue "'. When true goto script section: " GotoIfSection " (Else goto script section: " GotoElseSection ")`n" , LogFile
							GetValue := EnvGet(CheckVariable)
							If CheckValue != GetValue
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`n" , LogFile
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`n" , LogFile
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						}
					If (ScriptItemName = "IfExist" or ScriptItemName = "IfNotExist")
						{
						CheckFileFolder := ""
						GotoIfSection := ""
						GotoElseSection := ""
						Loop ScriptItemString.Length
							{
							If A_Index = 2
								CheckFileFolder := ScriptItemString[2]
							If A_Index = 3
								GotoIfSection := ScriptItemString[3]
							If A_Index = 4
								GotoElseSection := ScriptItemString[4]
							}
						CheckFileFolder := TransForm(CheckFileFolder)
						GotoIfSection := Trim(GotoIfSection)
						GotoElseSection := Trim(GotoElseSection)
						If ScriptItemName = "IfExist"
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tUsing IfExist statement for file or folder '" CheckFileFolder "'. When true goto script section: " GotoIfSection " (Else goto script section: " GotoElseSection ")`n" , LogFile
							If FileExist(CheckFileFolder)
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`n" , LogFile
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`n" , LogFile
									StatementCounter := Script(GotoElseSection)
									}
								}
							}
						If ScriptItemName = "IfNotExist"
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tUsing IfNotExist statement for file or folder '" CheckFileFolder "'. When true goto script section: " GotoIfSection " (Else goto script section: " GotoElseSection ")`n" , LogFile
							If Not FileExist(CheckFileFolder)
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoIfSection " ]`n" , LogFile
								StatementCounter := Script(GotoIfSection)
								}
							Else
								{
								If GotoElseSection
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`t[ Performing script section: " GotoElseSection " ]`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tFileCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`n" , LogFile						
						If FileExist(Source)
							{
							Try
								FileCopy Source, Destination, OverWrite
							Catch as Err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFileCopy failed for " err.Extra " files.`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tFolderCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`n" , LogFile						
						If DirExist(Source)
							{
							Try
								DirCopy Source, Destination, OverWrite
							Catch as err
								{
								If FileExist(LogFile)
									{
									If OverWrite = 1
										FileAppend A_Now "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`n" , LogFile
									Else
										FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`n" , LogFile
									}
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileCopy successfully executed.`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tFolderCreate: " Destination "`n" , LogFile
						If Not FileExist(Destination)
							{
							Try						
								DirCreate Destination
							Catch as Err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFolderCreate failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFolderCreate successfully executed.`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tFileDelete: " FilePattern "`n" , LogFile						
						If FileExist(FilePattern)
							{
							Try
								FileDelete FilePattern
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFileDelete failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFileDelete successfully executed.`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" FilePattern "' not found.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tFolderDelete: " DirName "`n" , LogFile
						If DirExist(DirName)
							{
							Try
								DirDelete DirName, 1
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tFolderDelete failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tFolderDelete successfully executed.`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" DirName "' not found.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tEnvSet setting environment variable name '" EnvSetName "' with value '" EnvSetValue "'`n" , LogFile
						Try
							EnvSet EnvSetName, EnvSetValue
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`n" , LogFile							
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" EnvSetName "' with value: " EnvSetValue "`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tRegRead registry item '" ValueName "' in '" KeyName "' with default value '" DefaultValue "'`n" , LogFile
						Try
							EnvSetValue := RegRead(KeyName, ValueName, DefaultValue)
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tRegRead failed with Error: " err.Message "`n" , LogFile
							EnvSetValue := DefaultValue
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegRead successfully executed with value: " EnvSetValue "`n" , LogFile
							}
						Try
							EnvSet ValueName, EnvSetValue
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`n" , LogFile							
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" ValueName "' with value: " EnvSetValue "`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tRegWrite create '" ValueType "' registry item '" ValueName "' in '" KeyName "' with value '" Value "'`n" , LogFile
						Try
							RegWrite Value, ValueType, KeyName, ValueName
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tRegWrite failed with Error: " err.Message "`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegWrite successfully executed.`n" , LogFile
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
								FileAppend A_Now "`t[ INFO  ]`tRegDelete registry value '" ValueName "' from '" KeyName "'`n" , LogFile	
							Try							
								RegDelete KeyName, ValueName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegDelete successfully executed.`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tRegDelete registry key: " KeyName "`n" , LogFile
							Try
								RegDeleteKey KeyName
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRegDelete successfully executed.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tIniRead from file name '" Filename "' using section name '" Section "' for key name '" KeyName "' with default value '" DefaultValue "'`n" , LogFile
						Try
							IniValue := IniRead(Filename, Section, KeyName, DefaultValue)
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tIniRead failed with Error: " err.Message "`n" , LogFile
							IniValue := DefaultValue
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniRead successfully executed with value: " IniValue "`n" , LogFile
							}
						Try
							EnvSet KeyName, IniValue
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tEnvironment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message "`n" , LogFile							
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tEnvironment variable successfully set for '" KeyName "' with value: " IniValue "`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tIniWrite value '" Value "' in file name '" Filename "' for section name '" Section "' with key name '" KeyName "'`n" , LogFile
						Try
							IniWrite Value, Filename, Section, KeyName
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tIniWrite failed with Error: " err.Message "`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniWrite successfully executed.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tIniDelete using file name '" Filename "' for section name '" Section "' (with key name '" KeyName "')`n" , LogFile	
						Try							
							IniDelete Filename, Section, KeyName
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ WARN  ]`tIniDelete failed with Error: " err.Message "`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tIniDelete successfully executed.`n" , LogFile
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
								FileAppend A_Now "`t[ INFO  ]`tRun and wait for Target: " Target "`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`n" , LogFile									
								}
							Try
								ExitCode := RunWait(Target, WorkingDir, Options, &PID)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully stopped with return code: " ExitCode "`n" , LogFile
								}
							}
						Else
							{
							If FileExist(LogFile)
								{
								FileAppend A_Now "`t[ INFO  ]`tRun Target: " Target "`n" , LogFile
								FileAppend A_Now "`t[ INFO  ]`tWorking directory: " WorkingDir "`n" , LogFile
								}
							Try							
								ExitCode := Run(Target, WorkingDir, Options, &PID)
							Catch as err
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ ERROR ]`tRun failed with Error: " err.Message "`n" , LogFile
								}
							Else
								{
								If FileExist(LogFile)
									FileAppend A_Now "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully started with return code: " ExitCode "`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tStringReplace text '" FindText "' with '" ReplaceText "' using source file '" Source "' for destination file: '" Destination "' (OverWrite=" OverWrite ")`n" , LogFile
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
										FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`n" , LogFile	
									}
								Else
									{
									If FileExist(LogFile)
										FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`n" , LogFile										
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
											FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`n" , LogFile	
										}
									Else
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`n" , LogFile										
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
												FileAppend A_Now "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`n" , LogFile							
											}
										Else
											{
											If FileExist(LogFile)
												FileAppend A_Now "`t[ INFO  ]`tStringReplace successfully executed.`n" , LogFile									
											}
										Contents := ""
										}
									Else
										{
										If FileExist(LogFile)
											FileAppend A_Now "`t[ INFO  ]`t'" Destination "' already exists.`n" , LogFile										
										}
									}									
								}
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`t'" Source "' not found.`n" , LogFile								
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
							FileAppend A_Now "`t[ INFO  ]`tFileAppend write text '" Text "' to file: " FileName " (" Encoding ")`n" , LogFile
						Text := Text "`n"
						Try					
							FileAppend Text, FileName, "`n " Encoding
						Catch as err
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`tFileAppend failed with Error: " err.Message "`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tFileAppend successfully executed.`n" , LogFile
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
							FileAppend A_Now "`t[ INFO  ]`tDownload '" URL "' for file: " FileName "`n" , LogFile
						Download URL, FileName
						If Not FileExist(FileName)
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ ERROR ]`tDownload failed with Error: " A_LastError "`n" , LogFile
							}
						Else
							{
							If FileExist(LogFile)
								FileAppend A_Now "`t[ INFO  ]`tDownload successfully executed.`n" , LogFile
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
