;@Ahk2Exe-SetName MSIX Helper 64-bit
;@Ahk2Exe-SetOrigFilename MSIXHelper64.exe
;@Ahk2Exe-SetDescription MSIX Helper 1.0
;@Ahk2Exe-SetVersion 1.0.0.0
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

ProductName := "MSIX Helper 1.0"
AppUserModelId := GetCurrentApplicationUserModelId()
If AppUserModelId
	{
	AppUserModelIdAppId := StrSplit(AppUserModelId, "!")
	PackageFamilyName := AppUserModelIdAppId[1]
	AppId := AppUserModelIdAppId[2]
	}
Else
	{
	ArgumentsNumber := A_Args.Length
	If ArgumentsNumber > 0
		{
		AppId := A_Args[1]
		PackageFamilyName := AppId
		}
	Else
		{
		AppId := ""
		PackageFamilyName := ""
		}	
	}

If AppId
	{
	PackagePath := GetCurrentPackagePath()
	EnvSet "PackageFamilyName", PackageFamilyName
	EnvSet "AppId", AppId
	EnvSet "PackagePath", PackagePath
	EnvSet "WorkingDir", A_ScriptDir
	FileLength := StrLen(A_ScriptName)
	FileLength := FileLength - 4
	FileName := SubStr(A_ScriptName, 1, FileLength)
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
		}
	Target := IniRead(IniFile, AppId, "Target", "")
	If Target
		{
		Target := TransForm(Target)
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
			PreLaunchCounter := Script(PreLaunchSection)
			}
		WorkingDir := IniRead(IniFile, AppId, "WorkingDir", A_ScriptDir)
		If WorkingDir
			WorkingDir := TransForm(WorkingDir)
		Options := IniRead(IniFile, AppId, "Options", A_ScriptDir)
		If Options
			Options := TransForm(Options)
		Parameters := ""
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
		ArgumentsNumber := A_Args.Length
		If ArgumentsNumber > 0
			{
			for index, Argument in A_Args
				{
				If InStr(Argument, A_Space)
					Argument := "`"" Argument "`""
				Parameters := Parameters . Argument . A_Space
				}
			}
		ParametersLenght := StrLen(Parameters)
		ParametersLenght := ParametersLenght - 1
		Parameters := SubStr(Parameters, 1, ParametersLenght)
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
					ScriptItemString := StrSplit(A_LoopReadLine, "=")
					ScriptItemName := ScriptItemString[1]
					ScriptItemValue := ScriptItemString[2]
					ScriptItemValueString := StrSplit(ScriptItemValue, ",")
					If ScriptItemName = "FileCopy"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								Source := ScriptItemValueString[1]
							If A_Index = 2
								Destination := ScriptItemValueString[2]
							If A_Index = 3
								OverWrite := ScriptItemValueString[3]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								Source := ScriptItemValueString[1]
							If A_Index = 2
								Destination := ScriptItemValueString[2]
							If A_Index = 3
								OverWrite := ScriptItemValueString[3]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								Destination := ScriptItemValueString[1]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								FilePattern := ScriptItemValueString[1]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								DirName := ScriptItemValueString[1]
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
					If ScriptItemName = "RegWrite"
						{
						ValueType := ""
						KeyName := ""
						ValueName := ""
						Value := ""
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								ValueType := ScriptItemValueString[1]
							If A_Index = 2
								KeyName := ScriptItemValueString[2]
							If A_Index = 3
								ValueName := ScriptItemValueString[3]
							If A_Index = 4
								Value := ScriptItemValueString[4]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								KeyName := ScriptItemValueString[1]
							If A_Index = 2
								ValueName := ScriptItemValueString[2]
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
					If ScriptItemName = "Run"
						{
						Target := ""
						WorkingDir := ""
						Options := ""
						NoWait := 0
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								Target := ScriptItemValueString[1]
							If A_Index = 2
								WorkingDir := ScriptItemValueString[2]
							If A_Index = 3
								Options := ScriptItemValueString[3]
							If A_Index = 4
								NoWait := ScriptItemValueString[4]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								Source := ScriptItemValueString[1]
							If A_Index = 2
								FindText := ScriptItemValueString[2]
							If A_Index = 3
								ReplaceText := ScriptItemValueString[3]
							If A_Index = 4
								Destination := ScriptItemValueString[4]
							If A_Index = 5
								OverWrite := ScriptItemValueString[5]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								Text := ScriptItemValueString[1]
							If A_Index = 2
								FileName := ScriptItemValueString[2]
							If A_Index = 3
								Encoding := ScriptItemValueString[3]
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
						Loop ScriptItemValueString.Length
							{
							If A_Index = 1
								URL := ScriptItemValueString[1]
							If A_Index = 2
								FileName := ScriptItemValueString[2]
							}
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
