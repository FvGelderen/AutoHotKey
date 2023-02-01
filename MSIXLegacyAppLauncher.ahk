/*
Language:			English
Platform:			Windows 10/11 32-bit or 64-bit
Author:				Ferry van Gelderen (ferry@provolve.nl)
Script Function:	MSIX Legacy Application Launcher
*/

#SingleInstance off
#NoTrayIcon
SendMode Input
SetWorkingDir %A_ScriptDir%
SetBatchLines -1
#Include %A_ScriptDir%\Anchor.ahk

AppName = MSIX Legacy Application Launcher
Version = 1.0.0.2

StringTrimRight, FileName, A_ScriptName, 4
IfExist, %A_Temp%\%FileName%.log
	FileDelete, %A_Temp%\%FileName%.log

IfNotExist, %A_ScriptDir%\%FileName%.ini
	{
	MsgBox, 64, %AppName% %Version%, Runs one or more external programs with the ability to set environments variables and perform basic script actions.`n`nPlease read the "%AppName%" document file for information on how to use this program.`n`nThis software is provided "as is". Use of the software is free and at your own risk.`n`nCopyright © 2023 Provolve IT B.V.
	ExitApp, 0
	}

IniRead, NoLog, %A_ScriptDir%\%FileName%.ini, MAIN, NoLog
If NoLog != 1
	{
	FileAppend, %A_Now%`tInfo`t%AppName% %Version% started for username: "%A_UserName%" on computer name: "%A_ComputerName%"`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tOperating System version: %A_OSVersion% (64-bit=%A_Is64bitOS%)`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tMain working directory: %A_ScriptDir%`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tMain executable name: %A_ScriptName%`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tMain INI file name: %FileName%.ini`n, %A_Temp%\%FileName%.log
	}

; --------------------------------------------------------------- Dark Mode GUI Settings ---------------------------------------------------------------
RegRead, AppsUseLightTheme, HKCU, SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize, AppsUseLightTheme
If AppsUseLightTheme = 0
	{
	ColorGUI = 2C2C2C
	ColorLVb = 191919
	ColorLVt = FFFFFF
	ColorEb = 202020
	}
Else
	{
	ColorGUI =
	ColorLVb =
	ColorLVt =
	ColorEb =
	}
if (A_OSVersion >= "10.0.17763" && SubStr(A_OSVersion, 1, 3) = "10.") 
	{
	attr := 19
	if (A_OSVersion >= "10.0.18985") 
		attr := 20
	}

; --------------------------------------------------------------- Collect MSIX Package Information ---------------------------------------------------------------
PackagePath := GetCurrentPackagePath(256)
If PackagePath
	AppxManifestFile := PackagePath . "\AppxManifest.xml"
Else
	AppxManifestFile =
If NoLog != 1
	{
	FileAppend, %A_Now%`tInfo`tMSIX PackagePath: %PackagePath%`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tMSIX ManifestFile: %AppxManifestFile%`n, %A_Temp%\%FileName%.log
	}		
IfExist, %AppxManifestFile%
	{
	If NoLog != 1
		FileAppend, %A_Now%`tInfo`tMSIX AppxManifest.xml file found.`n, %A_Temp%\%FileName%.log		
	}
Else
	{
	If NoLog != 1
		FileAppend, %A_Now%`tWarning`tMSIX AppxManifest.xml file not found.`n, %A_Temp%\%FileName%.log			
	}
PackageFullName := GetCurrentPackageFullName(128)
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tMSIX PackageFullName: %PackageFullName%`n, %A_Temp%\%FileName%.log
PackageFamilyName := GetCurrentPackageFamilyName(128)
If PackageFamilyName
	{
	StringSplit, PackageNamePublisherId, PackageFamilyName, _
	PackageName := PackageNamePublisherId1
	PublisherId := PackageNamePublisherId2
	}
Else
	{
	PublisherId =
	PackageName =
	}
If NoLog != 1
	{
	FileAppend, %A_Now%`tInfo`tMSIX PackageName: %PackageName%`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tMSIX PublisherId: %PublisherId%`n, %A_Temp%\%FileName%.log
	FileAppend, %A_Now%`tInfo`tMSIX PackageFamilyName: %PackageFamilyName%`n, %A_Temp%\%FileName%.log
	}
AppUserModelId := GetCurrentApplicationUserModelId(256)
If AppUserModelId
	{
	StringSplit, AppUserModelIdAppiD, AppUserModelId, !
	AppID :=  AppUserModelIdAppiD2	
	If NoLog != 1
		{
		FileAppend, %A_Now%`tInfo`tMSIX AppUserModelId: %AppUserModelId%`n, %A_Temp%\%FileName%.log
		FileAppend, %A_Now%`tInfo`tMSIX AppId: %AppId%`n, %A_Temp%\%FileName%.log
		}
	}
Else
	{
	AppId =
	If NoLog != 1
		FileAppend, %A_Now%`tWarning`tAppId not found.`n, %A_Temp%\%FileName%.log		
	}

; --------------------------------------------------------------- Skip MAIN GUI for Single App ---------------------------------------------------------------
IniRead, ButtonOK, %A_ScriptDir%\%FileName%.ini, MAIN, ButtonOK, OK
ButtonOK := " " . ButtonOK . " "
IniRead, App2, %A_ScriptDir%\%FileName%.ini, MAIN, App2
If App2 = ERROR
	{
	IniRead, Name, %A_ScriptDir%\%FileName%.ini, MAIN, App1
	If Name != ERROR
		{
		Transform, Name, Deref, %Name%
		ShowMAINGUI = 0
		AppNumber = App1
		Gosub, SelectShortcut
		Gosub, MAINGuiClose
		}
	}

; --------------------------------------------------------------- Skip MAIN GUI for registered AppId ---------------------------------------------------------------
Loop
	{
	App := "App" . A_Index
	IniRead, Name, %A_ScriptDir%\%FileName%.ini, MAIN, %App%
	If Name = ERROR
		Break
	Else
		{
		If Name = %AppId%
			{
			Transform, Name, Deref, %Name%
			ShowMAINGUI = 0
			AppNumber = %App%
			Gosub, SelectShortcut
			Gosub, MAINGuiClose			
			}
		}
	}

; --------------------------------------------------------------- Creating MAIN GUI ---------------------------------------------------------------
IniRead, Title, %A_ScriptDir%\%FileName%.ini, MAIN, Title, %AppName%
IniRead, Icon, %A_ScriptDir%\%FileName%.ini, MAIN, Icon, %A_Space%
IniRead, Height, %A_ScriptDir%\%FileName%.ini, MAIN, Height, 200
IniRead, Width, %A_ScriptDir%\%FileName%.ini, MAIN, Width, 330
IniRead, IconSize, %A_ScriptDir%\%FileName%.ini, MAIN, IconSize, 48
IconSize := IconSize*(A_ScreenDPI/96)

Gui, MAIN:New, +HwndstrHwnd
Gui, MAIN:+Resize -MaximizeBox -MinimizeBox
Gui, MAIN:+LastFound +DPIScale
Gui, MAIN:Margin, 5, 5
If ColorGUI
	Gui, MAIN:Color, %ColorGUI%
If (ColorLVb = "" and ColorLVt = "")
	Gui, MAIN:Add, ListView, h%Height% w%Width% gSelectShortcut vListView HWNDLV_LView -Theme +Icon -E0x200 +0x100 -Border -Multi, Shortcuts
Else
	Gui, MAIN:Add, ListView, h%Height% w%Width% gSelectShortcut vListView HWNDLV_LView -Theme +Icon -E0x200 +0x100 -Border -Multi +BackGround%ColorLVb% c%ColorLVt%, Shortcuts
Gui, MAIN:Add, Button, Hidden Default gSelectShortcut
Gosub, ReadMAIN
If AppsUseLightTheme = 0
	{
	GuiControlGet, ListViewHwnd, Hwnd, ListView
	DllCall("uxtheme\SetWindowTheme", "ptr", ListViewHwnd, "str", "DarkMode_Explorer", "ptr", 0)	
	DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", strHwnd, "Int", attr, "Int*", true, "Int", 4)
	}
Gui, MAIN:Show, AutoSize, %Title%
If Icon
	{
	Transform, Icon, Deref, %Icon%
	hIcon := DllCall("LoadImage", uint, 0, str, Icon, uint, 1, int, 0, int, 0, uint, 0x10)
	SendMessage, 0x80, 0, hIcon 
	SendMessage, 0x80, 1, hIcon	
	}
ShowMAINGUI = 1
WaitForClose = 0
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tShow Main GUI.`n, %A_Temp%\%FileName%.log
Return

; --------------------------------------------------------------- Exit Application ---------------------------------------------------------------
MAINGuiClose:
while WaitForClose = 1
	Sleep, 100
DllCall("DestroyIcon", "ptr", hIcon)
If NoLog != 1
	FileAppend, %A_Now%`tInfo`t%AppName% %Version% stopped.`n, %A_Temp%\%FileName%.log
ExitApp, %ErrorNumber%
Return

MAINGuiSize:
Anchor("SysListView321", "wh")
Return

ReadMAIN:
IconColor := 0x20
Gui, MAIN:Default
Gui, MAIN:+Disabled
LV_Delete()
hIml := ImageList_Create(IconSize, IconSize, IconColor, 100, 100)
LV_SetImageList(hIml) 
Loop
	{
	App := "App" . A_Index
	IniRead, Name, %A_ScriptDir%\%FileName%.ini, MAIN, %App%
	If Name = ERROR
		Break
	Else
		{
		Transform, Name, Deref, %Name%
		IniRead, AppIcon, %A_ScriptDir%\%FileName%.ini, %App%, Icon, %A_Space%
		StringSplit, icon_array, AppIcon, `,
		If icon_array2 =
			icon_array2 = 1
		hIcon := LoadIcon(icon_array1, icon_array2, IconSize)
		i := ImageList_AddIcon( hIml, hIcon )
		LV_Add("Icon" i+1, Name)
		}
	}
Gui, MAIN:-Disabled
Return

SelectShortcut:
If ShowMAINGUI = 1
	{
	ControlGet, Line, List, Count Focused, SysListView321,
	ControlGet, Name, List, Focused, SysListView321,
	If (Line = "" or Name = "")
		Return
	Gui, MAIN:+LastFound
	Gui, MAIN:+OwnDialogs
	Gui, MAIN:+Disabled
	AppNumber := "App" . Line
	}
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tRunning %AppNumber%: %Name%...`n, %A_Temp%\%FileName%.log
; --------------------------------------------------------------- Set ENVironment Variables ---------------------------------------------------------------
SetEnv("SetEnv", AppNumber)
; --------------------------------------------------------------- Get ENVironment Variables ---------------------------------------------------------------
IniRead, AppIcon, %A_ScriptDir%\%FileName%.ini, %AppNumber%, Icon, %A_Space%
Gosub, GetEnv
If AbortStart = 1
	{
	If ShowMAINGUI = 1
		Gui, MAIN:-Disabled
	Else
		ErrorNumber = 1223
	AbortStart = 0
	Return
	}
; --------------------------------------------------------------- Execute Pre-Launch SCRIPT actions ---------------------------------------------------------------
Script("PreLaunch", AppNumber)
; --------------------------------------------------------------- Get Shortcut information ---------------------------------------------------------------
IniRead, Target, %A_ScriptDir%\%FileName%.ini, %AppNumber%, Target, %A_Space%
Transform, Target, Deref, %Target%
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tTarget=%Target%`n, %A_Temp%\%FileName%.log
IniRead, WorkingDir, %A_ScriptDir%\%FileName%.ini, %AppNumber%, WorkingDir, %A_Space%
Transform, WorkingDir, Deref, %WorkingDir%
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tWorkingDir=%WorkingDir%`n, %A_Temp%\%FileName%.log
IniRead, Options, %A_ScriptDir%\%FileName%.ini, %AppNumber%, Options, %A_Space%
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tOptions=%Options%`n, %A_Temp%\%FileName%.log
ParamString =
Loop
	{
	Param := "Param" . A_Index
	IniRead, CurrentParam, %A_ScriptDir%\%FileName%.ini, %AppNumber%, %Param%
	If CurrentParam = ERROR
		Break
	If CurrentParam
		{
		Transform, CurrentParam, Deref, %CurrentParam%
		If NoLog != 1
			FileAppend, %A_Now%`tInfo`t%Param%=%CurrentParam%`n, %A_Temp%\%FileName%.log
		ParamString = %ParamString% %CurrentParam%
		}
	}
IniRead, Wait, %A_ScriptDir%\%FileName%.ini, %AppNumber%, Wait
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tWait=%Wait%`n, %A_Temp%\%FileName%.log
IniRead, RunInPackage, %A_ScriptDir%\%FileName%.ini, %AppNumber%, RunInPackage, 0
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tRunInPackage=%RunInPackage%`n, %A_Temp%\%FileName%.log
; --------------------------------------------------------------- Execute Target ---------------------------------------------------------------
If NoLog != 1
	FileAppend, %A_Now%`tInfo`tStarting Target...`n, %A_Temp%\%FileName%.log
EmptyMem()
If Target = ERROR
	{
	If NoLog != 1
		FileAppend, %A_Now%`tError`tTarget property is not found.`n, %A_Temp%\%FileName%.log
	}
Else
	{
	If Wait = 1
		{
		If RunInPackage = 0
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tRun and wait: %Target% %ParamString%`n, %A_Temp%\%FileName%.log			
			RunWait, %Target% %ParamString%, %WorkingDir%, %Options% UseErrorLevel, pID
			}
		If RunInPackage = 1
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tRun and wait: powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command "%Target%" -Args '%ParamString%' -PreventBreakaway`n, %A_Temp%\%FileName%.log				
			RunWait, powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command "%Target%" -Args '%ParamString%' -PreventBreakaway, %WorkingDir%, %Options% UseErrorLevel, pID
			}
		If RunInPackage = 2
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tRun and wait (Elevated): powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command "%Target%" -Args '%ParamString%' -PreventBreakaway`n, %A_Temp%\%FileName%.log			
			RunWait, *RunAs powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command "%Target%" -Args '%ParamString%' -PreventBreakaway, %WorkingDir%, %Options% UseErrorLevel, pID
			}
		If ErrorLevel = ERROR
			{
			ErrorNumber = %A_LastError%
			ErrorMessage := GetSysErrorText(ErrorNumber)
			If NoLog != 1
				FileAppend, %A_Now%`tError`tTarget launch failed with return code: %ErrorNumber% Message: %ErrorMessage%`n, %A_Temp%\%FileName%.log
			}
		Else
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tTarget with process ID %pID% successfully stopped with return code: %A_LastError%`n, %A_Temp%\%FileName%.log
 ;--------------------------------------------------------------- Execute Post-Exit SCRIPT actions ---------------------------------------------------------------
			WaitForClose = 1
			Script("PostExit", AppNumber)
			WaitForClose = 0
			}
		If ShowMAINGUI = 1
			WinActivate, %Title%
		}
	Else
		{
		If RunInPackage = 0
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tRun: %Target% %ParamString%`n, %A_Temp%\%FileName%.log			
			Run, %Target% %ParamString%, %WorkingDir%, %Options% UseErrorLevel, pID
			}
		If RunInPackage = 1
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tRun: powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command "%Target%" -Args '%ParamString%' -PreventBreakaway`n, %A_Temp%\%FileName%.log			
			Run, powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command '%Target%' -Args '%ParamString%' -PreventBreakaway, %WorkingDir%, %Options% UseErrorLevel, pID
			}
		If RunInPackage = 2
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tRun (Elevated): powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command "%Target%" -Args '%ParamString%' -PreventBreakaway`n, %A_Temp%\%FileName%.log				
			Run, *RunAs powershell.exe Invoke-CommandInDesktopPackage -AppId '%AppId%' -PackageFamilyName '%PackageFamilyName%' -Command '%Target%' -Args '%ParamString%' -PreventBreakaway, %WorkingDir%, %Options% UseErrorLevel, pID
			}
		If ErrorLevel = ERROR
			{
			ErrorNumber = %A_LastError%
			ErrorMessage := GetSysErrorText(ErrorNumber)
			If NoLog != 1
				FileAppend, %A_Now%`tError`tError Running: %Target% from %WorkingDir%. Return code: %ErrorNumber% Message: %ErrorMessage%`n, %A_Temp%\%FileName%.log
			}
		Else
			{
			If NoLog != 1
				FileAppend, %A_Now%`tInfo`tTarget successfully started with process ID: %pID%`n, %A_Temp%\%FileName%.log
			}
		If ShowMAINGUI = 0
			{
			WinWait, ahk_pid %pID%
			WinActivate, ahk_pid %pID%
			}
		}
	}
If ShowMAINGUI = 1
	{
	Gui, MAIN:-Disabled
	Gui, MAIN:Default
	}
Return

; --------------------------------------------------------------- Subroutine Get Environment Variables ---------------------------------------------------------------
GetEnv:
GETENVGUI = 0
Loop
	{
	EnvNumber := "GetEnv" . A_Index
	IniRead, CurrentEnv, %A_ScriptDir%\%FileName%.ini, %AppNumber%, %EnvNumber%
	If CurrentEnv = ERROR
		Break
	Loop, Parse, CurrentEnv, `t
		{
		If A_Index = 1
			EnvName = %A_LoopField%
		If A_Index = 2
			EnvValue = %A_LoopField%
		If A_Index = 3
			Hide = %A_LoopField%
		If A_Index > 2
			Break
		}
	If EnvName
		{
		If GETENVGUI = 0
			{
			Gui, GETENV:New, +HwndstrHwnd
			If ShowMAINGUI = 1
				Gui, GETENV:+ownerMAIN
			Gui, GETENV:-MaximizeBox -MinimizeBox
			Gui, GETENV:+LastFound +DPIScale
			Gui, GETENV:Margin, 5, 5			
			If ColorGUI
				Gui, GETENV:Color, %ColorGUI%			
			GETENVGUI = 1
			}
		EnvCount = %A_Index%
		Transform, EnvValue, Deref, %EnvValue%
		Gui, GETENV:Add, Text, vName%A_Index% xm y+5 R1 w120 c%ColorLVt% Right, %EnvName%
		Gui, GETENV:Add, Text, x+0 R1 w5 c%ColorLVt% Right, :
		If AppsUseLightTheme = 0
			{
			If Hide != 1
				Gui, GETENV:Add, Edit, vValue%A_Index% x+5 R1 w200 -Theme -E0x200 +0x100 -Border -Multi c%ColorLVt%, %EnvValue%
			Else
				Gui, GETENV:Add, Edit, vValue%A_Index% x+5 R1 w200 -Theme -E0x200 +0x100 -Border -Multi Password c%ColorLVt%, %EnvValue%
			Gui, GETENV:color,, %ColorEb%
			}
		Else
			{
			If Hide != 1
				Gui, GETENV:Add, Edit, vValue%A_Index% x+5 R1 w200 -Theme -E0x200 +0x100 -Border -Multi, %EnvValue%
			Else
				Gui, GETENV:Add, Edit, vValue%A_Index% x+5 R1 w200 -Theme -E0x200 +0x100 -Border -Multi Password, %EnvValue%			
			}
		}
	EnvName =
	EnvValue =
	Hide =
	}
If GETENVGUI = 1
	{
	Gui, GETENV:Add, Button, xm Default gGETENVButtonOK h20 -Border, %ButtonOK%
	If AppsUseLightTheme = 0
		{
		GuiControlGet, ButtonHwnd, Hwnd, %ButtonOK%
		DllCall("uxtheme\SetWindowTheme", "ptr", ButtonHwnd, "str", "DarkMode_Explorer", "ptr", 0)
		DllCall("dwmapi\DwmSetWindowAttribute", "Ptr", strHwnd, "Int", attr, "Int*", true, "Int", 4)
		}
	Gui, GETENV:Show, AutoSize, %Name%
	StringSplit, icon_array, AppIcon, `,
	If icon_array2 =
		icon_array2 = 1	
	hIcon := LoadIcon(icon_array1, icon_array2, IconSize)
	SendMessage, 0x80, 0, hIcon 
	SendMessage, 0x80, 1, hIcon
	If NoLog != 1
		FileAppend, %A_Now%`tInfo`tShow GetEnv GUI.`n, %A_Temp%\%FileName%.log	
	}
Loop
	{
	If GETENVGUI = 0
		Break
	Else
		Sleep, 500
	}
Return

GETENVGuiEscape:
GETENVGuiClose:
AbortStart = 1
Gui, GETENV:Destroy
If ShowMAINGUI = 1
	{
	Gui, MAIN:Default
	WinActivate, %Title%
	}
GETENVGUI = 0
Return	

GETENVButtonOK:
AbortStart = 0
Loop, %EnvCount%
	{
	GuiControlGet, EnvName,, Name%A_Index%
	GuiControlGet, EnvValue,, Value%A_Index%
	Transform, EnvValue, Deref, %EnvValue%
	EnvSet, %EnvName%, %EnvValue%
	If ErrorLevel
		{
		If NoLog != 1
			{
			ErrorNumber = %A_LastError%
			ErrorMessage := GetSysErrorText(ErrorNumber)			
			FileAppend, %A_Now%`tError`tError setting user environment: "%EnvName%" with value "%EnvValue%". Error: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log
			}
		}
	Else
		{
		If NoLog != 1
			FileAppend, %A_Now%`tInfo`tUser environment variable "%EnvName%" with value "%EnvValue%" set successfully.`n, %A_Temp%\%FileName%.log
		EnvSet = 1
		}	
	}
EnvUpdate
EnvCount =
EnvName =
EnvValue =
Hide =
Gui, GETENV:Destroy
If ShowMAINGUI = 1
	{
	Gui, MAIN:Default
	WinActivate, %Title%
	}
GETENVGUI = 0
Return

; --------------------------------------------------------------- Functions for showing ListView icons ---------------------------------------------------------------
ListView_SetImageList( hwnd, hIml, iImageList=0) 
	{ 
	SendMessage, 0x1000+3, iImageList, hIml, , ahk_id %hwnd% 
	Return ErrorLevel 
	} 
ImageList_Create(cx,cy,flags,cInitial,cGrow)
	{ 
	Return DllCall("comctl32.dll\ImageList_Create", "int", cx, "int", cy, "uint", flags, "int", cInitial, "int", cGrow) 
	} 
ImageList_Add(hIml, hbmImage, hbmMask="")
	{ 
	Return DllCall("comctl32.dll\ImageList_Add", "uint", hIml, "uint",hbmImage, "uint", hbmMask) 
	} 
ImageList_AddIcon(hIml, hIcon) 
	{ 
	Return DllCall("comctl32.dll\ImageList_ReplaceIcon", "uint", hIml, "int", -1, "uint", hIcon) 
	} 
API_ExtractIcon(Icon, Idx=0)
	{ 
	Return DllCall("shell32\ExtractIcon", "UInt", 0, "Str", Icon, "UInt",Idx) 
	} 
API_LoadImage(pPath, uType, cxDesired, cyDesired, fuLoad) 
	{ 
	Return DllCall("LoadImage", "uint", 0, "str", pPath, "uint", uType, "int", cxDesired, "int", cyDesired, "uint", fuLoad) 
	} 
LoadIcon(Filename, IconNumber, IconSize) 
	{ 
	DllCall("PrivateExtractIcons", "str", Filename, "int", IconNumber-1, "int", IconSize, "int", IconSize, "uint*", h_icon, "uint*", 0, "uint", 1, "uint", 0, "int") 
	If !ErrorLevel 
		Return h_icon 
	}	 

; --------------------------------------------------------------- Function Get Error Message ---------------------------------------------------------------
GetSysErrorText(errNr)
	{
	bufferSize = 1024
	VarSetCapacity(buffer, bufferSize)
	DllCall("FormatMessage"
		, "UInt", FORMAT_MESSAGE_FROM_SYSTEM := 0x1000
		, "UInt", 0
		, "UInt", errNr
		, "UInt", 0  ;LANG_USER_DEFAULT := 0x20000 ; LANG_SYSTEM_DEFAULT := 0x10000
		, "Str", buffer
		, "UInt", bufferSize
		, "UInt", 0)
	Return buffer
	}

; --------------------------------------------------------------- Function Free Memory ---------------------------------------------------------------
EmptyMem(pID="AHK Rocks")
	{
	pid := (pid="AHK Rocks") ? DllCall("GetCurrentProcessId") : pid
	h := DllCall("OpenProcess", "UInt", 0x001F0FFF, "Int", 0, "Int", pid)
	DllCall("SetProcessWorkingSetSize", "UInt", h, "Int", -1, "Int", -1)
	DllCall("CloseHandle", "Int", h)
	}

; --------------------------------------------------------------- Functions for Collecting MSIX Package Information ---------------------------------------------------------------
GetCurrentPackagePath(pathLength)
	{
	bytes_per_char := A_IsUnicode ? 2 : 1
	GetCurrentPackagePathLength := pathLength * bytes_per_char
	VarSetCapacity(GetCurrentPackagePath, GetCurrentPackagePathLength)
	DllCall("GetCurrentPackagePath", "Wstr", GetCurrentPackagePathLength, "Ptr", &GetCurrentPackagePath)
	PackagePath := StrGet(&GetCurrentPackagePath)
	Return PackagePath
	}
GetCurrentPackageFullName(packageFullNameLength)
	{
	bytes_per_char := A_IsUnicode ? 2 : 1
	GetCurrentPackageFullNameLength := packageFullNameLength * bytes_per_char
	VarSetCapacity(GetCurrentPackageFullName, GetCurrentPackageFullNameLength)
	DllCall("GetCurrentPackageFullName", "Wstr", GetCurrentPackageFullNameLength, "Ptr", &GetCurrentPackageFullName)
	PackageFullName := StrGet(&GetCurrentPackageFullName)
	Return PackageFullName
	}
GetCurrentPackageFamilyName(packageFamilyNameLength)
	{
	bytes_per_char := A_IsUnicode ? 2 : 1
	GetCurrentPackageFamilyNameLength := packageFamilyNameLength * bytes_per_char
	VarSetCapacity(GetCurrentPackageFamilyName, GetCurrentPackageFamilyNameLength)
	DllCall("GetCurrentPackageFamilyName", "Wstr", GetCurrentPackageFamilyNameLength, "Ptr", &GetCurrentPackageFamilyName)
	PackageFamilyName := StrGet(&GetCurrentPackageFamilyName)
	Return PackageFamilyName
	}
GetCurrentApplicationUserModelId(applicationUserModelIdLength)
	{
	bytes_per_char := A_IsUnicode ? 2 : 1
	GetCurrentApplicationUserModelIdLength := applicationUserModelIdLength * bytes_per_char		
	VarSetCapacity(GetCurrentApplicationUserModelId, GetCurrentApplicationUserModelIdLength)
	DllCall("GetCurrentApplicationUserModelId", "Wstr", GetCurrentApplicationUserModelIdLength, "Ptr", &GetCurrentApplicationUserModelId)
	AppUserModelId := StrGet(&GetCurrentApplicationUserModelId)
	Return AppUserModelId
	}

; --------------------------------------------------------------- Function Set Environment Variables ---------------------------------------------------------------
SetEnv(SetEnv, AppNumber)
	{
	global FileName
	EnvSet = 0
	Loop
		{
		EnvNumber := SetEnv . A_Index
		IniRead, CurrentEnv, %A_ScriptDir%\%FileName%.ini, %AppNumber%, %EnvNumber%
		If CurrentEnv = ERROR
			Break
		Loop, Parse, CurrentEnv, `t
			{
			If A_Index = 1
				EnvName = %A_LoopField%
			If A_Index = 2
				EnvValue = %A_LoopField%
			If A_Index = 3
				NoDeref = %A_LoopField%
			If A_Index > 2
				Break
			}
		If EnvName
			{
			If NoDeref != 1
				Transform, EnvValue, Deref, %EnvValue%
			EnvSet, %EnvName%, %EnvValue%
			If ErrorLevel
				{
				If NoLog != 1
					{
					ErrorNumber = %A_LastError%
					ErrorMessage := GetSysErrorText(ErrorNumber)			
					FileAppend, %A_Now%`tError`tError setting user environment: "%EnvName%" with value "%EnvValue%". Error: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log
					}
				}
			Else
				{
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tUser environment variable "%EnvName%" with value "%EnvValue%" set successfully.`n, %A_Temp%\%FileName%.log
				EnvSet = 1
				}
			}
		EnvName =
		EnvValue =
		NoDeref =
		}
	If EnvSet = 1
		EnvUpdate
	}

; --------------------------------------------------------------- Function Execute Pre-Launch and Post-Exit Basic Script Actions ---------------------------------------------------------------
Script(Line, AppNumber)
	{
	global FileName
	Loop
		{
		LineNumber := Line . A_Index
		IniRead, CurrentLine, %A_ScriptDir%\%FileName%.ini, %AppNumber%, %LineNumber%
		If CurrentLine = ERROR
			Break
		If CurrentLine
			{
			Loop, Parse, CurrentLine, `,
				{
				If A_Index = 1
					ScriptItem = %A_LoopField%
				If A_Index = 2
					FirstItem = %A_LoopField%
				If A_Index = 3
					SecondItem = %A_LoopField%
				If A_Index = 4
					ThirdItem = %A_LoopField%
				If A_Index = 5
					FourthItem = %A_LoopField%
				If A_Index > 5
					Break			
				}
			If ScriptItem = FileCopy
				{
				Transform, FirstItem, Deref, %FirstItem%
				Transform, SecondItem, Deref, %SecondItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tCopy file(s) from "%FirstItem%" to "%SecondItem%" overwrites existing file(s) is "%ThirdItem%"...`n, %A_Temp%\%FileName%.log
				FileCopy, %FirstItem%, %SecondItem%, %ThirdItem%
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)				
						FileAppend, %A_Now%`tWarning`tCopy file(s) stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log
						}
					IfExist, %SecondItem%
						FileAppend, %A_Now%`tInfo`t"%SecondItem%" exist.`n, %A_Temp%\%FileName%.log
					Else	
						FileAppend, %A_Now%`tWarning`t"%SecondItem%" not found.`n, %A_Temp%\%FileName%.log	
					}
				}
			If ScriptItem = FolderCopy
				{
				Transform, FirstItem, Deref, %FirstItem%
				Transform, SecondItem, Deref, %SecondItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tCopy folder from "%FirstItem%" to "%SecondItem%" overwrites existing file(s) and folder(s) is "%ThirdItem%"...`n, %A_Temp%\%FileName%.log
				FileCopyDir, %FirstItem%, %SecondItem%, %ThirdItem%
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)				
						FileAppend, %A_Now%`tWarning`tCopy folder stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					IfExist, %SecondItem%
						FileAppend, %A_Now%`tInfo`t"%SecondItem%" exist.`n, %A_Temp%\%FileName%.log
					Else	
						FileAppend, %A_Now%`tWarning`t"%SecondItem%" not found.`n, %A_Temp%\%FileName%.log						
					}
				}
			If ScriptItem = FolderCreate
				{
				Transform, FirstItem, Deref, %FirstItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tCreate folder "%FirstItem%"...`n, %A_Temp%\%FileName%.log
				FileCreateDir, %FirstItem%
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)				
						FileAppend, %A_Now%`tWarning`tCreate folder stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					IfExist, %FirstItem%
						FileAppend, %A_Now%`tInfo`t"%FirstItem%" exist.`n, %A_Temp%\%FileName%.log
					Else	
						FileAppend, %A_Now%`tWarning`t"%FirstItem%" not found.`n, %A_Temp%\%FileName%.log
					}
				}	
			If ScriptItem = FileDelete
				{
				Transform, FirstItem, Deref, %FirstItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tDelete file "%FirstItem%"...`n, %A_Temp%\%FileName%.log
				FileDelete, %FirstItem%
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)				
						FileAppend, %A_Now%`tWarning`tDelete file stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					IfNotExist, %FirstItem%
						FileAppend, %A_Now%`tInfo`t"%FirstItem%" not found.`n, %A_Temp%\%FileName%.log
					Else	
						FileAppend, %A_Now%`tWarning`t"%FirstItem%" still exist.`n, %A_Temp%\%FileName%.log						
					}
				}						
			If ScriptItem = FolderDelete
				{
				Transform, FirstItem, Deref, %FirstItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tDelete folder "%FirstItem%" removing all files and subdirectories...`n, %A_Temp%\%FileName%.log
				FileRemoveDir, %FirstItem%, 1
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)				
						FileAppend, %A_Now%`tWarning`tDelete folder stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					IfNotExist, %FirstItem%
						FileAppend, %A_Now%`tInfo`t"%FirstItem%" not found.`n, %A_Temp%\%FileName%.log
					Else	
						FileAppend, %A_Now%`tWarning`t"%FirstItem%" still exist.`n, %A_Temp%\%FileName%.log	
					}
				}			
			If ScriptItem = RegWrite
				{
				Transform, ThirdItem, Deref, %ThirdItem%
				Transform, FourthItem, Deref, %FourthItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tCreate "%FirstItem%" registry item "%ThirdItem%" in "%SecondItem%" with value "%FourthItem%"...`n, %A_Temp%\%FileName%.log
				RegWrite, %FirstItem%, %SecondItem%, %ThirdItem%, %FourthItem%
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)
						FileAppend, %A_Now%`tWarning`tCreate register item stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					RegRead, CheckItem, %SecondItem%, %ThirdItem%
					If CheckItem
						FileAppend, %A_Now%`tInfo`t"%SecondItem%\%ThirdItem%" exist.`n, %A_Temp%\%FileName%.log
					Else
						FileAppend, %A_Now%`tWarning`t"%SecondItem%\%ThirdItem%" not found.`n, %A_Temp%\%FileName%.log	
					}
				}
			If ScriptItem = RegDelete
				{
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tDelete registry item "%FirstItem%" with value name "%SecondItem%"...`n, %A_Temp%\%FileName%.log
				RegDelete, %FirstItem%, %SecondItem%
				If NoLog != 1
					{
					If (A_LastError or ErrorLevel)
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)				
						FileAppend, %A_Now%`tWarning`tDelete register item stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					RegRead, CheckItem, %FirstItem%, %SecondItem%
					If CheckItem
						FileAppend, %A_Now%`tWarning`t"%FirstItem%\%SecondItem%" still exist.`n, %A_Temp%\%FileName%.log
					Else
						FileAppend, %A_Now%`tInfo`t"%FirstItem%\%SecondItem%" not found.`n, %A_Temp%\%FileName%.log							
					}
				}
			If ScriptItem = Run
				{
				Transform, FirstItem, Deref, %FirstItem%
				Transform, SecondItem, Deref, %SecondItem%
				If NoLog != 1
					FileAppend, %A_Now%`tInfo`tRun "%FirstItem%" with working directory "%SecondItem%"...`n, %A_Temp%\%FileName%.log
				If FourthItem != 1
					RunWait, %FirstItem%, %SecondItem%, %ThirdItem% UseErrorLevel
				Else
					Run, %FirstItem%, %SecondItem%, %ThirdItem% UseErrorLevel
				If NoLog != 1
					{
					If ErrorLevel
						{
						ErrorNumber = %A_LastError%
						ErrorMessage := GetSysErrorText(ErrorNumber)
						FileAppend, %A_Now%`tWarning`tRunning "%FirstItem%" stopped with return code: %ErrorNumber% Message: %ErrorMessage%, %A_Temp%\%FileName%.log			
						}
					}
				}
			ScriptItem =
			FirstItem=
			SecondItem=
			ThirdItem=
			FourthItem=
			}
		}
	}
