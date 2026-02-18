;@Ahk2Exe-SetName Intune Win32 Launcher 64-bit
;@Ahk2Exe-SetOrigFilename IntuneLauncher.exe
;@Ahk2Exe-SetDescription Intune Win32 Launcher 2.0
;@Ahk2Exe-SetVersion 2.0.0.0
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
ProductVersion := "2.0.0.0"
SystemPipeName := "Intune Win32 Launcher\SYSTEM"
UserPipeName := "Intune Win32 Launcher\USER"

OSVersion := StrSplit(A_OSVersion, ".")
if Not OSVersion[1] >= 10
	{
	MsgBox "The operating system version is not supported for this product. Microsoft Windows 10 or higher is required.", ProductName " " ProductVersion, "16 T30"
	ExitApp 1150
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
			ClientToScreen(hWnd, lpPoint) => DllCall("User32\ClientToScreen", "ptr", hWnd, "ptr", lpPoint, "int")
			CreateSolidBrush(crColor) => DllCall('Gdi32\CreateSolidBrush', 'uint', crColor, 'ptr')
			DestroyIcon(hIcon) => DllCall("DestroyIcon", "ptr", hIcon)
			DWMSetWindowAttribute(hwnd, dwAttribute, pvAttribute, cbAttribute := 4) => DllCall("Dwmapi\DwmSetWindowAttribute", "Ptr" , hwnd, "UInt", dwAttribute, "Ptr*", &pvAttribute, "UInt", cbAttribute)
			DeleteObject(hObject) => DllCall('Gdi32\DeleteObject', 'ptr', hObject, 'int')
			EnumThreadWindows(dwThreadId, lpfn, lParam) => DllCall("User32\EnumThreadWindows", "uint", dwThreadId, "ptr", lpfn, "uptr", lParam, "int")
			FillRect(hDC, lprc, hbr) => DllCall("User32\FillRect", "ptr", hDC, "ptr", lprc, "ptr", hbr, "int")
			GetClientRect(hWnd, lpRect) => DllCall("User32\GetClientRect", "ptr", hWnd, "ptr", lpRect, "int")
			GetCurrentThreadId() => DllCall("Kernel32\GetCurrentThreadId", "uint")
			GetWindowDC(hwnd) => DllCall("User32\GetWindowDC", "ptr", hwnd, "ptr")
			GetWindowRect(hWnd, lpRect) => DllCall("User32\GetWindowRect", "ptr", hWnd, "ptr", lpRect, "uptr")
			GetWindowRgn(hWnd, hRgn) => DllCall("User32\GetWindowRgn", "ptr", hWnd, "ptr", hRgn, "int")
			GetWindowRgnBox(hWnd, hRgn) => DllCall("User32\GetWindowRgnBox", "ptr", hWnd, "ptr", hRgn, "int")
			ReleaseDC(hWnd, hDC) => DllCall("User32\ReleaseDC", "ptr", hWnd, "ptr", hDC, "int")
			ScreenToClient(hWnd, lpPoint) => DllCall("User32\ScreenToClient", "ptr", hWnd, "ptr", lpPoint, "int")
			SelectObject(hdc, hgdiobj) => DllCall('Gdi32\SelectObject', 'ptr', hdc, 'ptr', hgdiobj, 'ptr')
			SetBkColor(hdc, crColor) => DllCall('Gdi32\SetBkColor', 'ptr', hdc, 'uint', crColor, 'uint')
			SetBkMode(hdc, iBkMode) => DllCall('Gdi32\SetBkMode', 'ptr', hdc, 'int', iBkMode, 'int')
			SetTextColor(hdc, crColor) => DllCall('Gdi32\SetTextColor', 'ptr', hdc, 'uint', crColor, 'uint')
			SetThreadDpiAwarenessContext(dpiContext) => DllCall("SetThreadDpiAwarenessContext", "ptr", dpiContext, "ptr")
			SetWindowTheme(hwnd, pszSubAppName, pszSubIdList := "") => (!DllCall("uxtheme\SetWindowTheme", "ptr", hwnd, "ptr", StrPtr(pszSubAppName), "ptr", pszSubIdList ? StrPtr(pszSubIdList) : 0) ? true : false)
			}
		}
	}
else
	AppsUseLightTheme := 1
SplitPath A_ScriptFullPath,,,, &FileName
Inifile := A_ScriptDir "\" FileName ".ini"
TitleIcon := A_ScriptDir "\" FileName ".ico"
SetAltSubmit := False
if FileExist(A_ScriptDir "\" FileName ".png")
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
if (PID := ProcessExist("explorer.exe"))
	EnvSet "UserLogin", 1
else
	EnvSet "UserLogin", 0
ArgumentsNumber := A_Args.Length
if ArgumentsNumber > 0
	{
	if FileExist(Inifile)
		{
		SetEnvSection := "Variables"
		SetEnvCounter := SetEnv(SetEnvSection)
		if SetEnv(A_Language) = 0
			SetEnv("[0409")
		}
	Argument := A_Args[1]
	if InStr(Argument, ":")
		{
		Arg_array := StrSplit(Argument, ":")
		Section := Arg_array[1]
		NotifyCode := Arg_array[2]
		}
	else
		{
		Section := Argument
		NotifyCode := ""
		}
	if NotifyCode = ""
		{
		if A_IsAdmin = 0
			{
			MsgBox "Administrative privileges are required to execute the system operations.", ProductName " " ProductVersion, "16 T30"
			ExitApp 5
			}
		LogFile := IniRead(Inifile, Section, "LogFile", A_Temp "\" FileName ".log")
		if LogFile != A_Temp "\" FileName ".log"
			{
			LogFile := TransForm(LogFile)
			SplitPath LogFile,, &LogPath
			if !DirExist(LogPath)
				{
				try
					DirCreate LogPath
				catch as err
					LogFile := A_Temp "\" FileName ".log"
				}
			}
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" ProductName " " ProductVersion " started with Process Id " A_PID " for section: " Section "`r`n" , LogFile, "UTF-16"
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tComputer name: " A_ComputerName "`r`n" , LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tOperating System version: " A_OSVersion " (64-bit=" A_Is64bitOS ")`r`n" , LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tOperating System language code: " A_Language "`r`n" , LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUser name: " A_UserName "`r`n" , LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRebootRequired: " (EnvGet("RebootRequired")) "`r`n" , LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUserLogin: " (EnvGet("UserLogin")) "`r`n" , LogFile
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tWorking directory: " A_ScriptDir "`r`n" , LogFile
		}
	else
		{
		if NotifyCode = "*"
			{
			if FileExist(Inifile)
				{
				; --------------------------------------------------------------- USER Notification process ---------------------------------------------------------------
				LogFile := ""
				SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
				DllCall("kernel32.dll\SetProcessShutdownParameters", "UInt", 0x4FF, "UInt", 0)
				OnMessage(0x0011, On_WM_QUERYENDSESSION)
				ActiveProgress := 0
				Loop
					{
					SetCursorHwndList := ""
					GetMessage := Receive_NamedPipeMessage(SystemPipeName)
					if ActiveProgress = 1
						{
						SetTimer RunWaitProgress, 0
						NotifyGui.Destroy()
						ActiveProgress := 0
						}
					Arg_array := StrSplit(GetMessage, ":")
					Section := Arg_array[1]
					NotifyCode := Arg_array[2]
					SectionType := IniRead(Inifile, Section, "SectionType", "")
					if (SectionType = "Notification" or SectionType = "Progress" or SectionType = "Message" or Section = "*" or Section = "?")
						{
						if SectionType = "Notification"
							{
							RegWrite ProductVersion, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, ProductVersion
							RegWrite FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss"), "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, "TimeStamp"
							if AppsUseLightTheme = 0
								{
								BackColor := "2C2C2C"
								TextColor := "cF0F0F0"
								TimeColor := "cE78F79"
								LinkOptions := "s7 underline c6D99EE"
								}
							else
								{
								BackColor := "F0F0F0"
								TextColor := "c"
								TimeColor := "c721C24"
								LinkOptions := "s7 underline c194499"
								}
							ProductNameText := IniRead(Inifile, Section, "ProductNameText", "")
							ProductNameText := TransForm(ProductNameText)
							ProductVersionText := IniRead(Inifile, Section, "ProductVersionText", "")
							ProductVersionText := TransForm(ProductVersionText)
							InformationText := IniRead(Inifile, Section, "InformationText", "")
							InformationText := TransForm(InformationText)
							TimerText := IniRead(Inifile, Section, "TimerText", "")
							TimerText := TransForm(TimerText)
							DeferLinkText := IniRead(Inifile, Section, "DeferLinkText", "Defer")
							DeferLinkText := TransForm(DeferLinkText)
							ContinueLinkText := IniRead(Inifile, Section, "ContinueLinkText", "Continue")
							ContinueLinkText := TransForm(ContinueLinkText)
							Wait := IniRead(Inifile, Section, "Wait", 60)
							HideDeferLink := 0
							DeferTimes := IniRead(Inifile, Section, "DeferTimes", "")
							DeferTimes := Trim(DeferTimes)
							if (DeferTimes = "" or DeferTimes = 0)
								HideDeferLink := 1
							else
								{
								DeferredCount := RegRead("HKEY_CURRENT_USER\SOFTWARE\" . ProductName . "\DeferTimes", Section . " " . ProductNameText . " " . ProductVersionText, 0)
								if DeferTimes <= DeferredCount
									HideDeferLink := 1					
								}
							NotifyGui := Gui("+DpiScale -Caption +AlwaysOnTop +Owner +E0x08000000", Section)
							Hwnd := NotifyGui.Hwnd
							NotifyGui.MarginX := 10
							NotifyGui.MarginY := 0
							NotifyGui.BackColor := BackColor
							ShowBorder := NotifyGui.Add("Text", "x0 y0 w64 h128 vLineL Background646464")
							if FileExist(TitleIcon)
								{
								if SetAltSubmit = True
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans AltSubmit vIcon", TitleIcon)
								else
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon", TitleIcon)
								}
							else
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
							if AppsUseLightTheme = 0
								{
								EditHwnd := ShowInfo.Hwnd
								ShowInfo.Opt("Background2C2C2C")
								DllCall("uxtheme\SetWindowTheme", "ptr", EditHwnd, "str", "DarkMode_Explorer", "ptr", 0)
								}
							ShowTextTime := NotifyGui.Add("Text", "xp y100 wp-35 R1 BackgroundTrans Left", TimerText)
							ShowTextTime.Opt(TextColor)
							ShowTime := NotifyGui.Add("Text", "x+5 yp R1 BackgroundTrans Right", "00:00")
							ShowTime.Opt(TimeColor)
							if HideDeferLink = 0
								{
								DeferLink := NotifyGui.Add("Text", "x70 yp+13 w128", DeferLinkText)
								DeferLink.SetFont(LinkOptions)
								DeferLink.OnEvent("Click", CloseNotificationRetry)
								DeferLink.hCursor := LoadCursor(IDC_HAND := 32649)
								SetCursorHwndList .= DeferLink.hwnd . "|"
								ContinueLink := NotifyGui.Add("Text", "xp+128 yp w128 Right", ContinueLinkText)
								}
							else
								ContinueLink := NotifyGui.Add("Text", "x198 yp+13 w128 Right", ContinueLinkText)
							ContinueLink.SetFont(LinkOptions)
							ContinueLink.OnEvent("Click", CloseNotificationContinue)
							ContinueLink.hCursor := LoadCursor(IDC_HAND := 32649)
							SetCursorHwndList .= ContinueLink.hwnd . "|"
							Dummy := NotifyGui.Add("Edit", "xp yp w0 h0 ReadOnly", "")
							VirtualScreenHeight := SysGet(79)
							NotifyGui.Show("x0 y" . VirtualScreenHeight)
							OnMessage(0x20, WM_SETCURSOR)
							RoundCorners(NotifyGui.Hwnd)
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
								if (Counter = 0 or UserInteraction = 1)
									break
								}
							if UserInteraction = 0
								{
								OnMessage(0x20, WM_SETCURSOR, 0)
								NotifyGui.Destroy()
								Send_NamedPipeMessage("0", UserPipeName)
								}
							}
						if SectionType = "Progress"
							{
							RegWrite ProductVersion, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, ProductVersion
							RegWrite FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss"), "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, "TimeStamp"
							if AppsUseLightTheme = 0
								{
								BackColor := "2C2C2C"
								TextColor := "cF0F0F0"
								TimeColor := "cE78F79"
								LinkOptions := "s7 underline c6D99EE"
								}
							else
								{
								BackColor := "F0F0F0"
								TextColor := "c"
								TimeColor := "c721C24"
								LinkOptions := "s7 underline c194499"
								}
							ProductNameText := IniRead(Inifile, Section, "ProductNameText", "")
							ProductNameText := TransForm(ProductNameText)
							ProductVersionText := IniRead(Inifile, Section, "ProductVersionText", "")
							ProductVersionText := TransForm(ProductVersionText)
							InformationText := IniRead(Inifile, Section, "InformationText", "")
							InformationText := TransForm(InformationText)
							NotifyGui := Gui("+DpiScale -Caption +AlwaysOnTop +Owner +E0x08000000", Section)
							Hwnd := NotifyGui.Hwnd
							NotifyGui.MarginX := 10
							NotifyGui.MarginY := 0
							NotifyGui.BackColor := BackColor
							ShowBorder := NotifyGui.Add("Text", "x0 y0 w64 h128 vLineL Background3399FF")
							if FileExist(TitleIcon)
								{
								if SetAltSubmit = True
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans AltSubmit vIcon", TitleIcon)
								else
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans vIcon", TitleIcon)
								}
							else
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
							if AppsUseLightTheme = 0
								{
								EditHwnd := ShowInfo.Hwnd
								ShowInfo.Opt("Background2C2C2C")
								DllCall("uxtheme\SetWindowTheme", "ptr", EditHwnd, "str", "DarkMode_Explorer", "ptr", 0)	
								}
							NotifyProgress := NotifyGui.Add("Progress", "x70 y100 h8 w256 0x8")
							NotifyProgress.Opt("c3399FF")
							if AppsUseLightTheme = 0
								NotifyProgress.Opt("Background2C2C2C")
							Dummy := NotifyGui.Add("Edit", "xp yp w0 h0 ReadOnly", "")
							VirtualScreenHeight := SysGet(79)
							NotifyGui.Show("x0 y" . VirtualScreenHeight)
							RoundCorners(NotifyGui.Hwnd)
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
						if SectionType = "Message"
							{
							if NotifyCode = ""
								NotifyCode := 0
							RegWrite ProductVersion, "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, ProductVersion
							RegWrite FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss"), "REG_SZ", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName, "TimeStamp"
							if AppsUseLightTheme = 0
								{
								BackColor := "2C2C2C"
								TextColor := "cF0F0F0"
								TimeColor := "cE78F79"
								LinkOptions := "s7 underline c6D99EE"
								}
							else
								{
								BackColor := "F0F0F0"
								TextColor := "c"
								TimeColor := "c721C24"
								LinkOptions := "s7 underline c194499"
								}
							ProductNameText := IniRead(Inifile, Section, "ProductNameText", "")
							ProductNameText := TransForm(ProductNameText)
							ProductVersionText := IniRead(Inifile, Section, "ProductVersionText", "")
							ProductVersionText := TransForm(ProductVersionText)
							ContinueLinkText := IniRead(Inifile, Section, "ContinueLinkText", "Continue")
							ContinueLinkText := TransForm(ContinueLinkText)
							Status := IniRead(Inifile, Section, "Status", "Info")
							Status := TransForm(Status)
							Wait := IniRead(Inifile, Section, "Wait", 30)
							ShowWarning := 1
							ReturnCodes := IniRead(Inifile, Section, "ReturnCodes", "0, 1707")
							For index, ReturnCode in StrSplit(ReturnCodes, ",")
								{
								ReturnCode := Trim(ReturnCode)
								if NotifyCode = ReturnCode
									ShowWarning := 0
								}
							if ShowWarning = 0
								{
								UserScript := IniRead(Inifile, Section, "UserScript", "")
								UserScript := Trim(UserScript)
								InformationText := IniRead(Inifile, Section, "InformationText", "")
								InformationText := TransForm(InformationText)
								if (Status = "Warning" or Status = "Error")
									{
									if Status = "Warning"
										{
										ShowIconNumber := "Icon4"
										BarColor := "BackgroundEBB800"
										ProgressColor := "cEBB800"
										}
									if Status = "Error"
										{
										ShowIconNumber := "Icon2"
										BarColor := "BackgroundE40000"
										ProgressColor := "cE40000"
										}
									}
								else
									{
									ShowIconNumber := "Icon3"
									BarColor := "Background429300"
									ProgressColor := "c429300"
									}
								}
							else
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
							if FileExist(TitleIcon)
								{
								if SetAltSubmit = True
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans AltSubmit", TitleIcon)
								else
									ShowIcon := NotifyGui.Add("Picture", "x15 y5 w32 h32 BackgroundTrans", TitleIcon)
								}
							else
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
							if AppsUseLightTheme = 0
								{
								EditHwnd := ShowInfo.Hwnd
								ShowInfo.Opt("Background2C2C2C")
								DllCall("uxtheme\SetWindowTheme", "ptr", EditHwnd, "str", "DarkMode_Explorer", "ptr", 0)	
								}
							NotifyProgress := NotifyGui.Add("Progress", "x70 y100 h8 w256 0x8")
							NotifyProgress.Opt(ProgressColor)
							if AppsUseLightTheme = 0
								NotifyProgress.Opt("Background2C2C2C")
							Dummy := NotifyGui.Add("Edit", "xp yp w0 h0 ReadOnly", "")
							VirtualScreenHeight := SysGet(79)
							NotifyGui.Show("x0 y" . VirtualScreenHeight)
							RoundCorners(NotifyGui.Hwnd)
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
							if UserScript
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
							ContinueLink := NotifyGui.Add("Text", "xp+128 yp+13 w128 Right", ContinueLinkText)
							ContinueLink.SetFont(LinkOptions)
							ContinueLink.OnEvent("Click", CloseNotificationContinue)
							ContinueLink.hCursor := LoadCursor(IDC_HAND := 32649)
							SetCursorHwndList .= ContinueLink.hwnd . "|"
							OnMessage(0x20, WM_SETCURSOR)
							if (Status = "Warning" or Status = "Error")
								{
								if Status = "Warning"
									SoundPlay "*48"
								if Status = "Error"
									SoundPlay "*16"
								}
							else
								SoundPlay "*64"
							UserInteraction := 0
							Loop Wait
								{
								Sleep 1000
								if UserInteraction = 1
									break
								}
							if UserInteraction = 0
								{
								OnMessage(0x20, WM_SETCURSOR, 0)
								NotifyGui.Destroy()
								Send_NamedPipeMessage("0", UserPipeName)
								}
							}
						if Section = "*"
							{
							Send_NamedPipeMessage("0", UserPipeName)
							DllCall("shell32.dll\SHChangeNotify", "UInt", 0x08000000, "UInt", 0x0000, "Ptr", 0, "Ptr", 0)
							ExitApp 0
							}
						if Section = "?"
							{
							Send_NamedPipeMessage(A_UserName, UserPipeName)
							System_PID := NotifyCode
							SetTimer SystemProcessCheck, 1000
							}
						}
					else
						Send_NamedPipeMessage("Error", UserPipeName)
					}
				}
			else
				ExitApp 1
			}
		}
	if FileExist(Inifile)
		{
		if FileExist(LogFile)
			FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tINI file: " Inifile "`r`n" , LogFile
		SectionType := IniRead(Inifile, Section, "SectionType", "")
		SectionType := Trim(SectionType)
		if (SectionType = "Execution")
			{
			; --------------------------------------------------------------- SYSTEM Execution process ---------------------------------------------------------------
			FoundSessions := CheckForOtherSessions(0)
			if FoundSessions > 0
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ProductName " is already started for another session. This process will stop with return code: 1618 (Another Installation is already in progress)`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
				ExitApp 1618
				}
			Result := PsScriptManager.Register()
			if !Result = 0
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tPowershell COM registration error. psScript.dll failed to register on this system. Error: " Result  "`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
				ExitApp 1603
				}
			UserScript := ""
			SetTimer UserLoginPendingReboot, 250
			ActivePendingReboot := PendingReboot()
			if ActivePendingReboot
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tPending reboot has been detected.`r`n" , LogFile
				}
			if FileExist(LogFile)
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing Executing section: " Section " ]`r`n" , LogFile
			FoundProcess := 0
			CheckProcess := IniRead(Inifile, Section, "CheckProcess", "")
			if CheckProcess
				{
				CheckProcess := TransForm(CheckProcess)
				For index, Process in StrSplit(CheckProcess, ",")
					{
					Process := Trim(Process)
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheckProcess: " Process "`r`n" , LogFile
					if (PID := ProcessExist(Process))
						{
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" Process " exists with Process Id: " PID "`r`n" , LogFile
						FoundProcess++
						}
					}
				}
			else
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tContinuing without notification...`r`n" , LogFile
				}
			EnvSet "FoundProcess", FoundProcess
			if FileExist(LogFile)
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFoundProcess: " (EnvGet("FoundProcess")) "`r`n" , LogFile
			cpauPID := 0
			if FoundProcess > 0
				{
				if (PID := ProcessExist("explorer.exe"))
					{
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Activating Notification Process ]`r`n" , LogFile
					if FileExist(A_ScriptFullPath)
						SetSecurityError := SetSecurityFile(A_ScriptFullPath, "S-1-1-0", "1179817", "1")
					if FileExist(Inifile)
						SetSecurityError := SetSecurityFile(Inifile, "S-1-1-0", "1179817", "1")
					if FileExist(TitleIcon)
						SetSecurityError := SetSecurityFile(TitleIcon, "S-1-1-0", "1179817", "1")
					if FileExist(A_Temp . "\psScript.dll")
						SetSecurityError := SetSecurityFile(A_Temp . "\psScript.dll", "S-1-1-0", "1179817", "1")
					cpauPID := CreateProcessAsUser(A_ScriptName, "*:*", A_ScriptDir, "explorer.exe", 0)
					if (cpauPID = -1 or cpauPID = -2)
						{
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tCreateProcessAsUser function stopped with error: " cpauPID "`r`n" , LogFile
						}
					else
						{
						Send_NamedPipeMessage("?:" . A_PID, SystemPipeName)
						GetUserName := Receive_NamedPipeMessage(UserPipeName)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreateProcessAsUser function is started for " GetUserName " with Process Id: " cpauPID "`r`n" , LogFile
						EnvSet "UserLoginName", GetUserName
						}
					}
				if ActivePendingReboot
					{
					NotifyPendingReboot := IniRead(Inifile, Section, "NotifyPendingReboot", "")
					if NotifyPendingReboot
						{
						if (cpauPID := ProcessExist(cpauPID))
							{
							NotifyPendingReboot := Trim(NotifyPendingReboot)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing Notification section: " NotifyPendingReboot " ]`r`n" , LogFile
							Send_NamedPipeMessage(NotifyPendingReboot . ":0", SystemPipeName)
							ReturnCode := Receive_NamedPipeMessage(UserPipeName)
							if FileExist(LogFile)
								{
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ProductName " is cancelled successfully due to a pending reboot detection. (Return Code = 350)`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
								}
							Send_NamedPipeMessage("*:0", SystemPipeName)
							ReturnCode := Receive_NamedPipeMessage(UserPipeName)
							if (cpauPID := ProcessExist(cpauPID))
								ProcessClose cpauPID
							PsScriptManager.Unregister()
							ExitApp 350
							}
						}
					}
				NotifyStart := IniRead(Inifile, Section, "NotifyStart", "")
				if NotifyStart
					{
					if (cpauPID := ProcessExist(cpauPID))
						{
						NotifyStart := Trim(NotifyStart)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing Notification section: " NotifyStart " ]`r`n" , LogFile
						Send_NamedPipeMessage(NotifyStart . ":0", SystemPipeName)
						ReturnCode := Receive_NamedPipeMessage(UserPipeName)
						if (ReturnCode = 1602 or ReturnCode = 1618)
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`t" ProductName " is cancelled successfully. (Notification process returned: " ReturnCode ")`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
							Send_NamedPipeMessage("*:0", SystemPipeName)
							ReturnCode := Receive_NamedPipeMessage(UserPipeName)
							if (cpauPID := ProcessExist(cpauPID))
								ProcessClose cpauPID
							PsScriptManager.Unregister()
							ExitApp 1602
							}
						else
							{
							if ReturnCode = 0
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" ProductName " is cancelled successfully. (Return Code = 1602)`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
									}
								if !ReturnCode = "ExitApp"
									{
									Send_NamedPipeMessage("*:0", SystemPipeName)
									ReturnCode := Receive_NamedPipeMessage(UserPipeName)
									}
								if (cpauPID := ProcessExist(cpauPID))
									ProcessClose cpauPID
								PsScriptManager.Unregister()
								ExitApp 1602
								}
							}
						}
					}
				NotifyProgress := IniRead(Inifile, Section, "NotifyProgress", "")
				if NotifyProgress
					{
					if (cpauPID := ProcessExist(cpauPID))
						{
						NotifyProgress := Trim(NotifyProgress)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing Notification section: " NotifyProgress " ]`r`n" , LogFile
						Send_NamedPipeMessage(NotifyProgress . ":0", SystemPipeName)
						ReturnCode := Receive_NamedPipeMessage(UserPipeName)
						if ReturnCode = 0
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
							}
						}
					}
				}
			else
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" CheckProcess " not running. Continuing without notification...`r`n" , LogFile
				}
			PreLaunchSection := IniRead(Inifile, Section, "PreLaunchScript", "")
			if PreLaunchSection
				{
				PreLaunchSection := Trim(PreLaunchSection)
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing PreLaunch script section: " PreLaunchSection " ]`r`n" , LogFile
				ScriptSections := ""
				PreLaunchCounter := Script(PreLaunchSection)
				}
			else
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tContinuing without PreLaunch script section...`r`n" , LogFile
				}
			ExitCode := 0
			TaskKill := IniRead(Inifile, Section, "TaskKill", "")
			if TaskKill
				{
				TaskKill := TransForm(TaskKill)
				For index, StopProcess in StrSplit(TaskKill, ",")
					{
					StopProcess := Trim(StopProcess)
					Loop
						{
						if (PID := ProcessExist(StopProcess))
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStop process: " StopProcess " with Process Id: " PID "`r`n" , LogFile
							ProcessClose PID
							Sleep 1000
							}
						else
							break
						}
					}
				}
			else
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tContinuing without TaskKill...`r`n" , LogFile
				}
			WorkingDir := IniRead(Inifile, Section, "WorkingDir", "")
			if WorkingDir
				{
				WorkingDir := TransForm(WorkingDir)
				WorkingDir := A_ScriptDir "\" WorkingDir
				}
			else
				WorkingDir := A_ScriptDir
			SetSourceDir := IniRead(Inifile, Section, "SetSourceDir", "")
			if SetSourceDir
				{
				SetSourceDir := TransForm(SetSourceDir)
				DeleteSourceDir := IniRead(Inifile, Section, "DeleteSourceDir", "")
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreate SourceDir folder: " SetSourceDir " and copy content from: " WorkingDir "`r`n" , LogFile
				if !DirExist(SetSourceDir)
					DirCreate SetSourceDir
				try
					DirCopy WorkingDir, SetSourceDir, 1
				catch as err
					{
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`r`n" , LogFile
					}
				else
					{
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreate SourceDir and copy content successfully executed.`r`n" , LogFile
					}
				}
			else
				{
				DeleteSourceDir := ""
				SetSourceDir := WorkingDir
				}
			Options := IniRead(Inifile, Section, "Options", "")
			Options := TransForm(Options)
			Loop
				{
				Command := IniRead(Inifile, Section, "Command" A_Index, "")
				if Command = ""
					break
				else
					{
					Command := TransForm(Command)
					if (SubStr(Command, 1, 1) = "|" OR SubStr(Command, 1, 1) = ">")
						{
						if SubStr(Command, 1, 1) = "|"
							{
							MSIParam := StrSplit(Command, "|")
							if SubStr(SetSourceDir, StrLen(SetSourceDir), StrLen(SetSourceDir)) = "\"
								SetSourceDir := RTrim(SetSourceDir, "\")
							MSIPackagePath := SetSourceDir . "\" . MSIParam[2]
							MSICommandLine := MSIParam[3]
							MSILogFile := MSIParam[4]
							if FileExist(LogFile)
								{
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing MSI PackagePath: " MSIPackagePath "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing MSI CommandLine: " MSICommandLine "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing MSI LogFile: " MSILogFile "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStarting MsiInstallProduct...`r`n" , LogFile
								}
							try
								ExitCode := MsiInstallProduct(MSIPackagePath, MSICommandLine, MSILogFile)
							catch as err
								{
								ExitCode :=	1603
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tError: " err.Message "`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tMsiInstallProduct successfully stopped with return code: " ExitCode "`r`n" , LogFile
								}
							}
						if SubStr(Command, 1, 1) = ">"
							{
							Command := LTrim(Command, ">")
							if FileExist(LogFile)
								{
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing working directory: " SetSourceDir "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing Powershell CommandLine: " Command "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStarting PS_Script...`r`n" , LogFile
								}
							SplitPath Command,, &PSDir, &Extension
							if Extension = "ps1"
								{
								try
									Command := FileRead(Command)
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , LogFile
									}
								else
									if PSDir
										Command := "Set-Location -Path `"" PSDir "`"`r`n" Command
								}
							else
								{
								Command := StrReplace(Command, A_Space . "\n" . A_Space, "`r`n")
								Command := "Set-Location -Path `"" SetSourceDir "`"`r`n" Command
								}
							Command := Format(Command)
							try 
								psReturn := PsScriptManager.ComInstance.PS_Script(Command)
							catch as err
								{
								ExitCode :=	1603
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tPowershell command stopped with error:`r`n" err.Message "`r`n" , LogFile
								}
							else
								{
								ExitCode := 0
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tPowershell command successfully stopped (return code: 0) with return message:`r`n" psReturn "`r`n" , LogFile
								}
							}
						}
					else
						{
						if FileExist(LogFile)
							{
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing working directory: " SetSourceDir "`r`n" , LogFile
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing options: " Options "`r`n" , LogFile
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStarting command: " Command "`r`n" , LogFile
							}
						try
							ExitCode := RunWait(Command, SetSourceDir, Options, &PID)
						catch as err
							{
							ExitCode :=	1603
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tError: " err.Message "`r`n" , LogFile
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCommand with Process Id '" PID "' successfully stopped with return code: " ExitCode "`r`n" , LogFile
							}
						}
					EnvSet "ReturnCode", ExitCode
					}
				}
			PostExitSection := IniRead(Inifile, Section, "PostExitScript", "")
			if PostExitSection
				{
				PostExitSection := Trim(PostExitSection)
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing PostExit script section: " PostExitSection " ]`r`n" , LogFile
				ScriptSections := ""
				PostExitCounter := Script(PostExitSection)
				}
			else
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tContinuing without PostExit script section...`r`n" , LogFile
				}
			if DeleteSourceDir = 1
				{
				if FileExist(LogFile)
					FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDelete SourceDir folder: " SetSourceDir "`r`n" , LogFile
				try
					DirDelete SetSourceDir, 1
				catch as err
					{
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tDelete SourceDir folder failed with Error: " err.Message "`r`n" , LogFile
					}
				else
					{
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDelete SourceDir folder successfully executed.`r`n" , LogFile
					}
				}
			if (cpauPID := ProcessExist(cpauPID))
				{
				NotifyFinish := IniRead(Inifile, Section, "NotifyFinish", "")
				if NotifyFinish
					{
					NotifyFinish := Trim(NotifyFinish)
					if FileExist(LogFile)
						FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing Notification section: " NotifyFinish " ]`r`n" , LogFile
					Send_NamedPipeMessage(NotifyFinish . ":" . ExitCode, SystemPipeName)
					ReturnCode := Receive_NamedPipeMessage(UserPipeName)
					if ReturnCode = 0
						{
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
						Send_NamedPipeMessage("*:0", SystemPipeName)
						ReturnCode := Receive_NamedPipeMessage(UserPipeName)
						}
					else
						{
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
						}
					}
				}
			if FileExist(LogFile)
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" ProductName " stopped successfully. (Return Code = " ExitCode ")`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
			if (cpauPID := ProcessExist(cpauPID))
				ProcessClose cpauPID
			PsScriptManager.Unregister()
			DllCall("shell32.dll\SHChangeNotify", "UInt", 0x08000000, "UInt", 0x0000, "Ptr", 0, "Ptr", 0)
			ExitApp ExitCode
			}
		else
			{
			if FileExist(LogFile)
				FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tExecution SectionType for section '" Section "' not found. This process will stop with return code: 1 (Incorrect function.)`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
			MsgBox "Execution SectionType for section '" Section "' not found. This process will stop with return code: 1 (Incorrect function.)", ProductName " " ProductVersion, "16 T30"
			ExitApp 1
			}
		}
	else
		{
		MsgBox "No INI file found. Please make sure the following INI file exist:`n`n" Inifile, ProductName " " ProductVersion, "48 T30"
		ExitApp 1
		}
	}
else
	{
	if FileExist(Inifile)
		{
		MsgBox "No starting section name. Please make sure a starting section name is added as a parameter for this executable that will be used for the execution section in the following INI file:`n`n" Inifile, ProductName " " ProductVersion, "48 T30"
		ExitApp 1
		}
	else
		{
		; --------------------------------------------------------------- Create INI File ---------------------------------------------------------------
		SetupFilePath := FileSelect(3, A_ScriptDir, "Select a Windows Installer or Package file (exe, msi, mst, appv, msix, appx)", "(*.exe; *.msi; *.mst; *.appv; *.msix; *.appx")
		if SetupFilePath
			{
			ReservedCharacters := "<>:;`"/\|?*,|"
			SplitPath SetupFilePath, &SetupFileName, &Path, &Extension, &SetupFileNameNoExt
			WorkingDir := StrReplace(Path, A_ScriptDir, "")
			WorkingDir := LTrim(WorkingDir, "\")
			if Extension = "mst"
				TransformFileName := SetupFileName
			else
				TransformFileName := ""
			if TransformFileName
				{
				SetupFilePath := FileSelect(3, A_ScriptDir . "\" . WorkingDir, "Select the Windows Installer file (*.msi) for " TransformFileName, "(*.msi)")
				if SetupFilePath
					SplitPath SetupFilePath, &SetupFileName, &Path, &Extension
				}
			if AppsUseLightTheme = 0
				{
				BackColor := "2E2E2E"
				TextColor := "cF0F0F0"
				ProgressColor := "c4082F5 Background333333"
				}
			else
				{
				BackColor := "F9F9F9"
				TextColor := "c"
				ProgressColor := "c2169EB BackgroundE6E6E6"
				}
			ProgressGui := Gui("+DpiScale -Caption +AlwaysOnTop +Owner +E0x08000000",)
			ProgressGui.BackColor := BackColor
			ShowIcon := ProgressGui.Add("Picture", "x8 y5 w16 h16 BackgroundTrans", A_ScriptFullPath)
			ProgressGui.SetFont("S9 Bold")
			ShowText := ProgressGui.Add("Text", "xp+25 w250 R1 BackgroundTrans", ProductName " " ProductVersion)
			ShowText.Opt(TextColor)
			ProgressGui.SetFont("S7 norm")
			ShowInfo := ProgressGui.Add("Text", "xp y+3 w250 R1 BackgroundTrans", "Creating " FileName ".ini. Please wait...")
			ShowInfo.Opt(TextColor)
			NotifyProgress := ProgressGui.AddProgress("xp y+5 w250 h5 vMyProgress 0x8", 0)
			NotifyProgress.Opt(ProgressColor)
			ProgressGui.Show
			RoundCorners(ProgressGui.Hwnd)
			SetTimer RunWaitProgress, 50
			FileAppend "───────────────────────────────┤ " . ProductName . " " . ProductVersion . " Configuration file ├───────────────────────────────`r`n", Inifile, "UTF-16"
			FileAppend "`r`n", Inifile
			if Extension = "exe"
				{
				GetManufacturer := GetFileVersionInfo(SetupFilePath, "CompanyName")
				GetProductName := GetFileVersionInfo(SetupFilePath, "ProductName")
				GetProductVersion := GetFileVersionInfo(SetupFilePath, "ProductVersion")
				GetFileDescription := GetFileVersionInfo(SetupFilePath, "FileDescription")
				LogFileName := GetManufacturer . "_" . GetProductName . "_" . GetProductVersion
				loop parse ReservedCharacters
					if InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")
				IniWrite A_UserName, Inifile, "Variables", "Author"
				IniWrite FormatTime(, "ShortDate"), Inifile, "Variables", "Date"
				IniWrite GetManufacturer, Inifile, "Variables", "Manufacturer"
				IniWrite GetProductName, Inifile, "Variables", "ProductName"
				IniWrite GetProductVersion, Inifile, "Variables", "ProductVersion"
				IniWrite GetFileDescription, Inifile, "Variables", "FileDescription"
				IniWrite "%Manufacturer% %ProductName% %ProductVersion%", Inifile, "Variables", "PackageName"
				IniWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher\Packages", Inifile, "Variables", "RegistrationKey"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", Inifile, "Variables", "LogFolder"
				IniWrite LogFileName, Inifile, "Variables", "LogFileName"
				IniWrite SetupFileName, Inifile, "Variables", "SetupFileName"
				FileAppend "`r`n", Inifile
				IniWrite "Pending reboot detected. Please restart your system first to continue this installation.", Inifile, "0409", "InformationTextNotifyPendingReboot"
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", Inifile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed. Save your work and close this application when it is active.", Inifile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", Inifile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", Inifile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being installed. Do not turn off your computer during the installation. Please wait...", Inifile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", Inifile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", Inifile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", Inifile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", Inifile, "0409", "DeferLinkText"
				IniWrite "Continue", Inifile, "0409", "ContinueLinkText"
				IniWrite "Close", Inifile, "0409", "CloseLinkText"
				FileAppend "`r`n", Inifile
				IniWrite "In afwachting van opnieuw opstarten. Graag zelf uw computer opnieuw opstarten zodat deze installatie doorgang krijgt.", Inifile, "0413", "InformationTextNotifyPendingReboot"
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", Inifile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", Inifile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt geïnstalleerd. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", Inifile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", Inifile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", Inifile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", Inifile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", Inifile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", Inifile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", Inifile, "0413", "CloseLinkText"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Installation by starting: " . FileName . ".exe Install ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", Inifile, "Install", "LogFile"
				IniWrite WorkingDir, Inifile, "Install", "WorkingDir"
				IniWrite "", Inifile, "Install", "SetSourceDir"
				IniWrite "", Inifile, "Install", "DeleteSourceDir"
				IniWrite "", Inifile, "Install", "TaskKill"
				IniWrite "%SetupFileName% /install /quiet /log `"%LogFolder%\%LogFileName%_EXE_Install.log`"", Inifile, "Install", "Command1"
				IniWrite "explorer.exe", Inifile, "Install", "CheckProcess"
				IniWrite "NotifyPendingReboot", Inifile, "Install", "NotifyPendingReboot"
				IniWrite "NotifyStartInstall", Inifile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", Inifile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", Inifile, "Install", "NotifyFinish"
				IniWrite "", Inifile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", Inifile, "Install", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_Install]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_Install]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 1`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyPendingReboot", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyPendingReboot", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyPendingReboot", "ProductVersionText"
				IniWrite "%InformationTextNotifyPendingReboot%", Inifile, "NotifyPendingReboot", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyPendingReboot", "ContinueLinkText"
				IniWrite "Error", Inifile, "NotifyPendingReboot", "Status"
				IniWrite "", Inifile, "NotifyPendingReboot", "UserScript"
				IniWrite "300", Inifile, "NotifyPendingReboot", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyStart%", Inifile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", Inifile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyProgress%", Inifile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%ProductName%", Inifile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", Inifile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishInstall", "UserScript"
				IniWrite "60", Inifile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Uninstallation by starting: " . FileName . ".exe UnInstall ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", Inifile, "UnInstall", "LogFile"
				IniWrite WorkingDir, Inifile, "UnInstall", "WorkingDir"
				IniWrite "", Inifile, "UnInstall", "SetSourceDir"
				IniWrite "", Inifile, "UnInstall", "DeleteSourceDir"
				IniWrite "", Inifile, "UnInstall", "TaskKill"
				IniWrite "%SetupFileName% /uninstall /quiet /log `"%LogFolder%\%LogFileName%_EXE_UnInstall.log`"", Inifile, "UnInstall", "Command1"
				IniWrite "explorer.exe", Inifile, "UnInstall", "CheckProcess"
				IniWrite "", Inifile, "UnInstall", "NotifyPendingReboot"
				IniWrite "NotifyStartUnInstall", Inifile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", Inifile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", Inifile, "UnInstall", "NotifyFinish"
				IniWrite "", Inifile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", Inifile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_UnInstall]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_UnInstall]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 0`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyStart%", Inifile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", Inifile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyProgress%", Inifile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%ProductName%", Inifile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", Inifile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "60", Inifile, "NotifyFinishUnInstall", "Wait"
				SetTimer RunWaitProgress, 0
				ProgressGui.Destroy()
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected setup executable File:`n`n" Inifile "`n`nPlease check the install and uninstall commandline parameters.", ProductName " " ProductVersion, "64 T30"
				}
			if Extension = "msi"
				{
				GetManufacturer := WindowsInstaller(SetupFilePath, "Manufacturer")
				GetProductName := WindowsInstaller(SetupFilePath, "ProductName")
				GetProductVersion := WindowsInstaller(SetupFilePath, "ProductVersion")
				GetProductCode := WindowsInstaller(SetupFilePath, "ProductCode")
				LogFileName := GetManufacturer . "_" . GetProductName . "_" . GetProductVersion
				loop parse ReservedCharacters
					if InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")
				IniWrite A_UserName, Inifile, "Variables", "Author"
				IniWrite FormatTime(, "ShortDate"), Inifile, "Variables", "Date"
				IniWrite GetManufacturer, Inifile, "Variables", "Manufacturer"
				IniWrite GetProductName, Inifile, "Variables", "ProductName"
				IniWrite GetProductVersion, Inifile, "Variables", "ProductVersion"
				IniWrite GetProductCode, Inifile, "Variables", "ProductCode"
				IniWrite "%Manufacturer% %ProductName% %ProductVersion%", Inifile, "Variables", "PackageName"
				IniWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher\Packages", Inifile, "Variables", "RegistrationKey"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", Inifile, "Variables", "LogFolder"
				IniWrite LogFileName, Inifile, "Variables", "LogFileName"
				IniWrite SetupFileName, Inifile, "Variables", "SetupFileName"
				if TransformFileName
					IniWrite TransformFileName, Inifile, "Variables", "TransformFileName"
				FileAppend "`r`n", Inifile
				IniWrite "Pending reboot detected. Please restart your system first to continue this installation.", Inifile, "0409", "InformationTextNotifyPendingReboot"
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", Inifile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed. Save your work and close this application when it is active.", Inifile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", Inifile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", Inifile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being installed. Do not turn off your computer during the installation. Please wait...", Inifile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", Inifile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", Inifile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", Inifile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", Inifile, "0409", "DeferLinkText"
				IniWrite "Continue", Inifile, "0409", "ContinueLinkText"
				IniWrite "Close", Inifile, "0409", "CloseLinkText"
				FileAppend "`r`n", Inifile
				IniWrite "In afwachting van opnieuw opstarten. Graag zelf uw computer opnieuw opstarten zodat deze installatie doorgang krijgt.", Inifile, "0413", "InformationTextNotifyPendingReboot"
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", Inifile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", Inifile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt geïnstalleerd. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", Inifile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", Inifile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", Inifile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", Inifile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", Inifile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", Inifile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", Inifile, "0413", "CloseLinkText"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Installation by starting: " . FileName . ".exe Install ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", Inifile, "Install", "LogFile"
				IniWrite WorkingDir, Inifile, "Install", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%ProductCode%", Inifile, "Install", "SetSourceDir"
				IniWrite "0", Inifile, "Install", "DeleteSourceDir"
				IniWrite "", Inifile, "Install", "TaskKill"
				if TransformFileName
					IniWrite "|%SetupFileName%|TRANSFORMS=`"%TransformFileName%`" TRANSFORMSSECURE=`"1`" REBOOT=`"ReallySuppress`"|%LogFolder%\%LogFileName%_MSI_Install.log", Inifile, "Install", "Command1"
				else
					IniWrite "|%SetupFileName%|REBOOT=`"ReallySuppress`"|%LogFolder%\%LogFileName%_MSI_Install.log", Inifile, "Install", "Command1"
				IniWrite "explorer.exe", Inifile, "Install", "CheckProcess"
				IniWrite "NotifyPendingReboot", Inifile, "Install", "NotifyPendingReboot"
				IniWrite "NotifyStartInstall", Inifile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", Inifile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", Inifile, "Install", "NotifyFinish"
				IniWrite "", Inifile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", Inifile, "Install", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_Install]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_Install]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 1`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyPendingReboot", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyPendingReboot", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyPendingReboot", "ProductVersionText"
				IniWrite "%InformationTextNotifyPendingReboot%", Inifile, "NotifyPendingReboot", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyPendingReboot", "ContinueLinkText"
				IniWrite "Error", Inifile, "NotifyPendingReboot", "Status"
				IniWrite "", Inifile, "NotifyPendingReboot", "UserScript"
				IniWrite "300", Inifile, "NotifyPendingReboot", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyStart%", Inifile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", Inifile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%ProductName% %InstallInformationTextNotifyProgress%", Inifile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%ProductName%", Inifile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", Inifile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishInstall", "UserScript"
				IniWrite "60", Inifile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Uninstallation by starting: " . FileName . ".exe UnInstall ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", Inifile, "UnInstall", "LogFile"
				IniWrite WorkingDir, Inifile, "UnInstall", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%ProductCode%", Inifile, "UnInstall", "SetSourceDir"
				IniWrite "1", Inifile, "UnInstall", "DeleteSourceDir"
				IniWrite "", Inifile, "UnInstall", "TaskKill"
				IniWrite "|%SetupFileName%|REMOVE=`"ALL`" REBOOT=`"ReallySuppress`"|%LogFolder%\%LogFileName%_MSI_UnInstall.log", Inifile, "UnInstall", "Command1"
				IniWrite "explorer.exe", Inifile, "UnInstall", "CheckProcess"
				IniWrite "", Inifile, "UnInstall", "NotifyPendingReboot"
				IniWrite "NotifyStartUnInstall", Inifile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", Inifile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", Inifile, "UnInstall", "NotifyFinish"
				IniWrite "", Inifile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", Inifile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_UnInstall]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_UnInstall]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 0`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyStart%", Inifile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", Inifile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%ProductName%", Inifile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%ProductName% %UnInstallInformationTextNotifyProgress%", Inifile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%ProductName%", Inifile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", Inifile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "60", Inifile, "NotifyFinishUnInstall", "Wait"
				SetTimer RunWaitProgress, 0
				ProgressGui.Destroy()
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected Windows Installer File(s):`n`n" Inifile, ProductName " " ProductVersion, "64 T30"
				}
			if Extension = "appv"
				{
				GetPackageId := ""
				GetVersionId := ""
				GetProductVersion := ""
				GetDisplayName := ""
				GetPackageDescription := ""
				AppxManifestXml := Read_AppxManifest(SetupFilePath)
				Loop parse, AppxManifestXml, "`n", "`r"
					{
					if InStr(A_LoopField, "<Identity")
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
							if InStr(A_LoopField, "Version=")
								{
								GetProductVersion := StrReplace(A_LoopField , "Version=")
								GetProductVersion := StrReplace(GetProductVersion, "`"")
								}
							if InStr(A_LoopField, "appv:PackageId=")
								{
								GetPackageId := StrReplace(A_LoopField , "appv:PackageId=")
								GetPackageId := StrReplace(GetPackageId, "`"")
								}
							if InStr(A_LoopField, "appv:VersionId=")
								{
								GetVersionId := StrReplace(A_LoopField , "appv:VersionId=")
								GetVersionId := StrReplace(GetVersionId, "`"")
								}
							}
						}
					if InStr(A_LoopField, "<DisplayName>")
						{
						GetDisplayName := Trim(A_LoopField)
						GetDisplayName := StrReplace(GetDisplayName , "<DisplayName>")
						GetDisplayName := StrReplace(GetDisplayName , "</DisplayName>")
						}
					if InStr(A_LoopField, "<appv:AppVPackageDescription>")
						{
						GetPackageDescription := Trim(A_LoopField)
						GetPackageDescription := StrReplace(GetPackageDescription , "<appv:AppVPackageDescription>")
						GetPackageDescription := StrReplace(GetPackageDescription , "</appv:AppVPackageDescription>")
						}
					}
				AppxManifestXml := ""
				LogFileName := GetDisplayName
				loop parse ReservedCharacters
					if InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")
				IniWrite A_UserName, Inifile, "Variables", "Author"
				IniWrite FormatTime(, "ShortDate"), Inifile, "Variables", "Date"
				IniWrite GetDisplayName, Inifile, "Variables", "DisplayName"
				IniWrite GetProductVersion, Inifile, "Variables", "ProductVersion"
				IniWrite GetPackageDescription, Inifile, "Variables", "PackageDescription"
				IniWrite GetPackageId, Inifile, "Variables", "PackageId"
				IniWrite GetVersionId, Inifile, "Variables", "VersionId"
				IniWrite "%DisplayName%", Inifile, "Variables", "PackageName"
				IniWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher\Packages", Inifile, "Variables", "RegistrationKey"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", Inifile, "Variables", "LogFolder"
				IniWrite LogFileName, Inifile, "Variables", "LogFileName"
				IniWrite SetupFileName, Inifile, "Variables", "SetupFileName"
				FileAppend "`r`n", Inifile
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", Inifile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed from your system. Save your work and close this application when it is active.", Inifile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", Inifile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", Inifile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being added to your system. Do not turn off your computer during the installation. Please wait...", Inifile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", Inifile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", Inifile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", Inifile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", Inifile, "0409", "DeferLinkText"
				IniWrite "Continue", Inifile, "0409", "ContinueLinkText"
				IniWrite "Close", Inifile, "0409", "CloseLinkText"
				FileAppend "`r`n", Inifile
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", Inifile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", Inifile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt toegevoegd aan het systeem. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", Inifile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", Inifile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", Inifile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", Inifile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", Inifile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", Inifile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", Inifile, "0413", "CloseLinkText"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Installation by starting: " . FileName . ".exe Install ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", Inifile, "Install", "LogFile"
				IniWrite WorkingDir, Inifile, "Install", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageId%", Inifile, "Install", "SetSourceDir"
				IniWrite "0", Inifile, "Install", "DeleteSourceDir"
				IniWrite "", Inifile, "Install", "TaskKill"
				IniWrite ">Add-AppVClientPackage -path '%SetupFileName%' \n Publish-AppVClientPackage -Global -PackageId %PackageId% -VersionId %VersionId%", Inifile, "Install", "Command1"
				IniWrite "explorer.exe", Inifile, "Install", "CheckProcess"
				IniWrite "", Inifile, "Install", "NotifyPendingReboot"
				IniWrite "", Inifile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", Inifile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", Inifile, "Install", "NotifyFinish"
				IniWrite "Check_AppVStatus", Inifile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", Inifile, "Install", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_AppVStatus]`r`n", Inifile
				FileAppend "RegRead, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppV\Client, Enabled, 0`r`n", Inifile
				FileAppend "ifEqual, Enabled, 0, EnableAppV`r`n", Inifile
				FileAppend "RegRead, HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\AppV\Client\Scripting, EnablePackageScripts, 0`r`n", Inifile
				FileAppend "ifEqual, EnablePackageScripts, 0, EnablePackageScripts`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[EnableAppV]`r`n", Inifile
				FileAppend "Powershell, Enable-Appv`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[EnablePackageScripts]`r`n", Inifile
				FileAppend "Powershell, Set-AppvClientConfiguration -EnablePackageScripts 1`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Check_Install]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_Install]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 1`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyStart%", Inifile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", Inifile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyProgress%", Inifile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%DisplayName%", Inifile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", Inifile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishInstall", "UserScript"
				IniWrite "5", Inifile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Uninstallation by starting: " . FileName . ".exe UnInstall ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", Inifile, "UnInstall", "LogFile"
				IniWrite WorkingDir, Inifile, "UnInstall", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageId%", Inifile, "UnInstall", "SetSourceDir"
				IniWrite "1", Inifile, "UnInstall", "DeleteSourceDir"
				IniWrite "", Inifile, "UnInstall", "TaskKill"
				IniWrite ">Stop-AppVClientPackage -Name '%DisplayName%' -Global \n Remove-AppVClientPackage -PackageId %PackageId% -VersionId %VersionId%", Inifile, "UnInstall", "Command1"
				IniWrite "explorer.exe", Inifile, "UnInstall", "CheckProcess"
				IniWrite "", Inifile, "UnInstall", "NotifyPendingReboot"
				IniWrite "", Inifile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", Inifile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", Inifile, "UnInstall", "NotifyFinish"
				IniWrite "", Inifile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", Inifile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_UnInstall]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_UnInstall]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 0`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyStart%", Inifile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", Inifile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyProgress%", Inifile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%DisplayName%", Inifile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", Inifile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "5", Inifile, "NotifyFinishUnInstall", "Wait"
				SetTimer RunWaitProgress, 0
				ProgressGui.Destroy()
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected APPV File:`n`n" Inifile, ProductName " " ProductVersion, "64 T30"
				}
			if (Extension = "msix" OR Extension = "appx")
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
					if InStr(A_LoopField, "<Identity")
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
							if InStr(A_LoopField, "Name=")
								{
								GetProductName := StrReplace(A_LoopField, "Name=")
								GetProductName := StrReplace(GetProductName, "`"")
								}
							if InStr(A_LoopField, "Version=")
								{
								GetProductVersion := StrReplace(A_LoopField , "Version=")
								GetProductVersion := StrReplace(GetProductVersion, "`"")
								}
							if InStr(A_LoopField, "ProcessorArchitecture=")
								{
								GetProcessorArchitecture := StrReplace(A_LoopField , "ProcessorArchitecture=")
								GetProcessorArchitecture := StrReplace(GetProcessorArchitecture, "`"")
								}
							}
						}
					if InStr(A_LoopField, "<DisplayName>")
						{
						GetDisplayName := Trim(A_LoopField)
						GetDisplayName := StrReplace(GetDisplayName , "<DisplayName>")
						GetDisplayName := StrReplace(GetDisplayName , "</DisplayName>")
						}
					if InStr(A_LoopField, "<PublisherDisplayName>")
						{
						GetPublisher := Trim(A_LoopField)
						GetPublisher := StrReplace(GetPublisher, "<PublisherDisplayName>")
						GetPublisher := StrReplace(GetPublisher, "</PublisherDisplayName>")
						}
					if InStr(A_LoopField, "<Description>")
						{
						GetDescription := Trim(A_LoopField)
						GetDescription := StrReplace(GetDescription, "<Description>")
						GetDescription := StrReplace(GetDescription, "</Description>")
						}
					}
				AppxManifestXml := ""
				LogFileName := GetDisplayName
				loop parse ReservedCharacters
					if InStr(LogFileName, A_LoopField)
						LogFileName := StrReplace(LogFileName, A_LoopField)
				LogFileName := StrReplace(LogFileName, A_Space, "_")
				LogFileName := StrReplace(LogFileName, ".", "_")
				IniWrite A_UserName, Inifile, "Variables", "Author"
				IniWrite FormatTime(, "ShortDate"), Inifile, "Variables", "Date"
				IniWrite GetProductName, Inifile, "Variables", "ProductName"
				IniWrite GetProductVersion, Inifile, "Variables", "ProductVersion"
				IniWrite GetProcessorArchitecture, Inifile, "Variables", "ProcessorArchitecture"
				IniWrite GetDisplayName, Inifile, "Variables", "DisplayName"
				IniWrite GetPublisher, Inifile, "Variables", "Publisher"
				IniWrite GetDescription, Inifile, "Variables", "PackageDescription"
				IniWrite SetupFileNameNoExt, Inifile, "Variables", "PackageName"
				IniWrite "%DisplayName%", Inifile, "Variables", "PackageName"
				IniWrite "HKEY_LOCAL_MACHINE\SOFTWARE\Intune Win32 Launcher\Packages", Inifile, "Variables", "RegistrationKey"
				IniWrite "%ProgramData%\Microsoft\IntuneManagementExtension\Logs", Inifile, "Variables", "LogFolder"
				IniWrite LogFileName, Inifile, "Variables", "LogFileName"
				IniWrite SetupFileName, Inifile, "Variables", "SetupFileName"
				FileAppend "`r`n", Inifile
				IniWrite "is about to be installed on this system. Save your work and close this application when it is active.", Inifile, "0409", "InstallInformationTextNotifyStart"
				IniWrite "is about to be removed from your system. Save your work and close this application when it is active.", Inifile, "0409", "UnInstallInformationTextNotifyStart"
				IniWrite "The installation will start in (mm:ss):", Inifile, "0409", "InstallTimerTextNotifyStart"
				IniWrite "The removal will start in (mm:ss):", Inifile, "0409", "UnInstallTimerTextNotifyStart"
				IniWrite "is being added to your system. Do not turn off your computer during the installation. Please wait...", Inifile, "0409", "InstallInformationTextNotifyProgress"
				IniWrite "is being removed from this system. Do not turn off your computer during the removal. Please wait...", Inifile, "0409", "UnInstallInformationTextNotifyProgress"
				IniWrite "The installation of the software is successfully completed.", Inifile, "0409", "InstallInformationTextNotifyFinish"
				IniWrite "The removal of the software is successfully finished.", Inifile, "0409", "UnInstallInformationTextNotifyFinish"
				IniWrite "Defer", Inifile, "0409", "DeferLinkText"
				IniWrite "Continue", Inifile, "0409", "ContinueLinkText"
				IniWrite "Close", Inifile, "0409", "CloseLinkText"
				FileAppend "`r`n", Inifile
				IniWrite "wordt zo meteen geïnstalleerd op dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "InstallInformationTextNotifyStart"
				IniWrite "wordt zo meteen verwijderd van dit systeem. Sluit de betreffende applicatie af indien deze actief is.", Inifile, "0413", "UnInstallInformationTextNotifyStart"
				IniWrite "De installatie begint na (mm:ss):", Inifile, "0413", "InstallTimerTextNotifyStart"
				IniWrite "Het verwijderen begint na (mm:ss):", Inifile, "0413", "UnInstallTimerTextNotifyStart"
				IniWrite "wordt toegevoegd aan het systeem. Zorg ervoor dat dit systeem actief blijft gedurende de installatie. Even geduld a.u.b...", Inifile, "0413", "InstallInformationTextNotifyProgress"
				IniWrite "wordt verwijderd. Zorg ervoor dat dit systeem actief blijft gedurende het verwijderen. Even geduld a.u.b...", Inifile, "0413", "UnInstallInformationTextNotifyProgress"
				IniWrite "De installatie is succesvol afgerond.", Inifile, "0413", "InstallInformationTextNotifyFinish"
				IniWrite "Het verwijderen is succesvol afgerond.", Inifile, "0413", "UnInstallInformationTextNotifyFinish"
				IniWrite "Uitstellen", Inifile, "0413", "DeferLinkText"
				IniWrite "Doorgaan", Inifile, "0413", "ContinueLinkText"
				IniWrite "Sluiten", Inifile, "0413", "CloseLinkText"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Installation by starting: " . FileName . ".exe Install ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "Install", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_Install.log", Inifile, "Install", "LogFile"
				IniWrite WorkingDir, Inifile, "Install", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageName%", Inifile, "Install", "SetSourceDir"
				IniWrite "0", Inifile, "Install", "DeleteSourceDir"
				IniWrite "", Inifile, "Install", "TaskKill"
				IniWrite ">Add-AppProvisionedPackage -online -packagepath '%SetupFileName%' -skiplicense", Inifile, "Install", "Command1"
				IniWrite "explorer.exe", Inifile, "Install", "CheckProcess"
				IniWrite "", Inifile, "Install", "NotifyPendingReboot"
				IniWrite "", Inifile, "Install", "NotifyStart"
				IniWrite "NotifyProgressInstall", Inifile, "Install", "NotifyProgress"
				IniWrite "NotifyFinishInstall", Inifile, "Install", "NotifyFinish"
				IniWrite "", Inifile, "Install", "PreLaunchScript"
				IniWrite "Check_Install", Inifile, "Install", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_Install]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_Install`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_Install]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 1`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyStartInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyStart%", Inifile, "NotifyStartInstall", "InformationText"
				IniWrite "%InstallTimerTextNotifyStart%", Inifile, "NotifyStartInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyProgressInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressInstall", "ProductVersionText"
				IniWrite "%DisplayName% %InstallInformationTextNotifyProgress%", Inifile, "NotifyProgressInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishInstall", "ReturnCodes"
				IniWrite "%DisplayName%", Inifile, "NotifyFinishInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishInstall", "ProductVersionText"
				IniWrite "%InstallInformationTextNotifyFinish%", Inifile, "NotifyFinishInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishInstall", "UserScript"
				IniWrite "5", Inifile, "NotifyFinishInstall", "Wait"
				FileAppend "`r`n", Inifile
				FileAppend "───────────────────────────────┤ Uninstallation by starting: " . FileName . ".exe UnInstall ├───────────────────────────────`r`n", Inifile
				IniWrite "Execution", Inifile, "UnInstall", "SectionType"
				IniWrite "%LogFolder%\%LogFileName%_UnInstall.log", Inifile, "UnInstall", "LogFile"
				IniWrite WorkingDir, Inifile, "UnInstall", "WorkingDir"
				IniWrite "%ProgramData%\Package Cache\%PackageName%", Inifile, "UnInstall", "SetSourceDir"
				IniWrite "1", Inifile, "UnInstall", "DeleteSourceDir"				
				IniWrite "", Inifile, "UnInstall", "TaskKill"
				IniWrite ">Remove-AppPackage -AllUsers -package '%PackageName%'", Inifile, "UnInstall", "Command1"
				IniWrite "explorer.exe", Inifile, "UnInstall", "CheckProcess"
				IniWrite "", Inifile, "UnInstall", "NotifyPendingReboot"
				IniWrite "", Inifile, "UnInstall", "NotifyStart"
				IniWrite "NotifyProgressUnInstall", Inifile, "UnInstall", "NotifyProgress"
				IniWrite "NotifyFinishUnInstall", Inifile, "UnInstall", "NotifyFinish"
				IniWrite "", Inifile, "UnInstall", "PreLaunchScript"
				IniWrite "Check_UnInstall", Inifile, "UnInstall", "PostExitScript"
				FileAppend "`r`n", Inifile
				FileAppend "[Check_UnInstall]`r`n", Inifile
				FileAppend "ifEqual, ReturnCode, 0 OR 1707 OR 1641 OR 3010, Register_UnInstall`r`n", Inifile
				FileAppend "`r`n", Inifile
				FileAppend "[Register_UnInstall]`r`n", Inifile
				FileAppend "RegWrite, REG_DWORD, %RegistrationKey%, %PackageName%, 0`r`n", Inifile
				FileAppend "`r`n", Inifile
				IniWrite "Notification", Inifile, "NotifyStartUnInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyStartUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyStartUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyStart%", Inifile, "NotifyStartUnInstall", "InformationText"
				IniWrite "%UnInstallTimerTextNotifyStart%", Inifile, "NotifyStartUnInstall", "TimerText"
				IniWrite "%DeferLinkText%", Inifile, "NotifyStartUnInstall", "DeferLinkText"
				IniWrite "%ContinueLinkText%", Inifile, "NotifyStartUnInstall", "ContinueLinkText"
				IniWrite "3", Inifile, "NotifyStartUnInstall", "DeferTimes"
				IniWrite "300", Inifile, "NotifyStartUnInstall", "Wait"
				FileAppend "`r`n", Inifile
				IniWrite "Progress", Inifile, "NotifyProgressUnInstall", "SectionType"
				IniWrite "%DisplayName%", Inifile, "NotifyProgressUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyProgressUnInstall", "ProductVersionText"
				IniWrite "%DisplayName% %UnInstallInformationTextNotifyProgress%", Inifile, "NotifyProgressUnInstall", "InformationText"
				FileAppend "`r`n", Inifile
				IniWrite "Message", Inifile, "NotifyFinishUnInstall", "SectionType"
				IniWrite "0,1707", Inifile, "NotifyFinishUnInstall", "ReturnCodes"
				IniWrite "%DisplayName%", Inifile, "NotifyFinishUnInstall", "ProductNameText"
				IniWrite "%ProductVersion%", Inifile, "NotifyFinishUnInstall", "ProductVersionText"
				IniWrite "%UnInstallInformationTextNotifyFinish%", Inifile, "NotifyFinishUnInstall", "InformationText"
				IniWrite "%CloseLinkText%", Inifile, "NotifyFinishUnInstall", "ContinueLinkText"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "Status"
				IniWrite "", Inifile, "NotifyFinishUnInstall", "UserScript"
				IniWrite "5", Inifile, "NotifyFinishUnInstall", "Wait"
				SetTimer RunWaitProgress, 0
				ProgressGui.Destroy()
				MsgBox "The following INI file is created using the startup parameters `"Install`" and `"UnInstall`" for the selected MSIX/APPX File:`n`n" Inifile, ProductName " " ProductVersion, "64 T30"
				}
			}
		ExitApp 0
		}
	}

; --------------------------------------------------------------- GUI function to round GUI corners ---------------------------------------------------------------
RoundCorners(Hwnd) 
	{
    WinGetClientPos(&gX, &gY, &gWidth, &gHeight, Hwnd)
    WinSetRegion(Format("0-0 w{1} h{2} r10-10", gWidth, gHeight), Hwnd)
	}

; --------------------------------------------------------------- GUI function for clicking continue ---------------------------------------------------------------
CloseNotificationContinue(*)
	{
	Global UserInteraction
	UserInteraction := 1
	NotifyGui.Destroy()
	Send_NamedPipeMessage("0", UserPipeName)
	}

; --------------------------------------------------------------- GUI function for clicking defer (cancel) ---------------------------------------------------------------
CloseNotificationRetry(*)
	{
	Global UserInteraction
	UserInteraction := 1
	ProductNameText := ShowProductName.Value
	ProductVersionText := ShowProductVersion.Value
	Section := NotifyGui.Title
	if ProductVersionText = ""
		ProductVersionText := "1"
	DeferredCount := RegRead("HKEY_CURRENT_USER\SOFTWARE\" . ProductName . "\DeferTimes", Section . " " . ProductNameText . " " . ProductVersionText, 0)
	DeferredCount++
	RegWrite DeferredCount, "REG_DWORD", "HKEY_CURRENT_USER\SOFTWARE\" . ProductName . "\DeferTimes", Section . " " . ProductNameText . " " . ProductVersionText
	NotifyGui.Destroy()
	Send_NamedPipeMessage("1602", UserPipeName)
	}

; --------------------------------------------------------------- GUI function to change mouse pointer when hovering over a GUI control Hwnd listed in the variable: SetCursorHwndList ---------------------------------------------------------------
WM_SETCURSOR(wp, *)
	{
	if InStr(SetCursorHwndList, wp . "|")
		return DllCall('SetCursor', 'Ptr', GuiCtrlFromHwnd(wp).hCursor)
	}

LoadCursor(cursorId)
	{
	static IMAGE_CURSOR := 2, flags := (LR_DEFAULTSIZE := 0x40) | (LR_SHARED := 0x8000)
	return DllCall('LoadImage', 'Ptr', 0, 'UInt', cursorId, 'UInt', IMAGE_CURSOR, 'Int', 0, 'Int', 0, 'UInt', flags, 'Ptr')
	}

; --------------------------------------------------------------- Timer (USER) to check if the SYSTEM process (PID) is up and running ---------------------------------------------------------------
SystemProcessCheck()
	{
	Global System_PID
	if System_PID
		{
		if !(ProcessExist(System_PID))
			{
			NotifyGui.Destroy()
			MsgBox "The communication is lost with the main (SYSTEM) process with PID: " System_PID, ProductName " " ProductVersion, "16 T30"
			ExitApp 1
			}
		}
	}

; --------------------------------------------------------------- GUI timer for Run and wait notification progress bar ---------------------------------------------------------------
RunWaitProgress(ProgressBar := 0)
	{
	NotifyProgress.Value := ProgressBar++
	if ProgressBar = 100
		ProgressBar := 0
	Sleep 50
	}

; --------------------------------------------------------------- Function for detecting system shutdown/logoff and allows the user to abort it ---------------------------------------------------------------
On_WM_QUERYENDSESSION(wParam, lParam, *)
	{
	Global ProductName
	ENDSESSION_LOGOFF := 0x80000000
	if (lParam & ENDSESSION_LOGOFF)
		EventType := "Logoff"
	else
		EventType := "Shutdown"
	Try
		{
		BlockShutdown(ProductName . " attempting to prevent " . EventType ".")
		return False
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
	if (hPipe = -1)
		return hPipe
	DllCall("ConnectNamedPipe", "ptr", hPipe, "ptr", 0)
	f := FileOpen(hPipe, "h")
	f.Write(Message . "`n")
	f.ReadLine()
	DllCall("CloseHandle", "ptr", hPipe)
	return "delivered"
	}

; --------------------------------------------------------------- Functions for receiving a Named Pipe message ---------------------------------------------------------------
Receive_NamedPipeMessage(PipeName := "PipeName")
	{
	While !DllCall("WaitNamedPipe", "Str", "\\.\pipe\" . PipeName, "UInt", 0xffffffff)
		Sleep 250
	f := FileOpen("\\.\pipe\" . PipeName, "r")
	PipeMessage := f.ReadLine()
	f.Close()
	return PipeMessage
	}

; --------------------------------------------------------------- Timer for setting the "UserLogin" and "RebootRequired" session environment variables to value 1 or 0 during runtime ---------------------------------------------------------------
UserLoginPendingReboot()
	{
	if (PID := ProcessExist("explorer.exe"))
		EnvSet "UserLogin", 1
	else
		EnvSet "UserLogin", 0
	EnvSet "RebootRequired", PendingReboot()
	}

; --------------------------------------------------------------- Function to check if a reboot is pending on the system. Returns 1 (true) or 0 (false) ---------------------------------------------------------------
PendingReboot()
	{
	RebootRequired := RegKeyExists("HKLM", "SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
	RebootPending := RegKeyExists("HKLM", "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
	PendingFileRenameOperations := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager", "PendingFileRenameOperations", 0)
	if (RebootRequired = 1 or RebootPending = 1 or PendingFileRenameOperations != 0)
		return 1
	else
		return 0
	}

; --------------------------------------------------------------- Function to check if a registry key exists. Returns 1 (true) or 0 (false) ---------------------------------------------------------------
RegKeyExists(RootKey := 0x80000002, SubKey := "SOFTWARE")
	{
	if (RootKey = "HKEY_CLASSES_ROOT" or RootKey = "HKCR")
		RootKey := 0x80000000
	if (RootKey = "HKEY_CURRENT_USER" or RootKey = "HKCU")
		RootKey := 0x80000001
	if (RootKey = "HKEY_LOCAL_MACHINE" or RootKey = "HKLM")
		RootKey := 0x80000002
	if (RootKey = "HKEY_USERS" or RootKey = "HKU")
		RootKey := 0x80000003
	if (RootKey = "HKEY_CURRENT_CONFIG" or RootKey = "HKCC")
		RootKey := 0x80000005
	hKey := "69.420.parrot"
	ERROR_SUCCESS := 0
	KEY_READ := 131097
	RegKeyExists := ERROR_SUCCESS == DllCall("RegOpenKeyExW", "Ptr", RootKey, "Wstr", SubKey, "Uint", 0, "Uint", KEY_READ, "Str", hKey)
	DllCall("RegCloseKey", "Str", hKey)
	return RegKeyExists
	}

; --------------------------------------------------------------- Function for getting the system error message text using the error number ---------------------------------------------------------------
GetSysErrorText(errNr){
	bufferSize := 1024
	VarSetStrCapacity(&RetBuffer, bufferSize)
	DllCall("FormatMessage", "UInt", FORMAT_MESSAGE_FROM_SYSTEM := 0x1000, "UInt", 0, "UInt", errNr, "UInt", 0, "Str", Retbuffer, "UInt", bufferSize, "UInt", 0)
	return Retbuffer
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
	return A_LastError
	}

; --------------------------------------------------------------- Function for retrieving property values using the EXE file name and the property name ("Comments", "CompanyName", "FileDescription", "FileVersion", "InternalName", "LegalCopyright", "LegalTrademarks", "OriginalFilename", "PrivateBuild", "ProductName", "ProductVersion", "SpecialBuild") ---------------------------------------------------------------
GetFileVersionInfo(ExeFile := "", Property := "ProductName")
    {
	if FileExist(ExeFile)
		{
		if (Size := DllCall("version\GetFileVersionInfoSizeW", "Str", ExeFile, "Ptr", 0, "UInt"))
			{
			Data := Buffer(Size)
			if (DllCall("version\GetFileVersionInfoW", "Str", ExeFile, "UInt", 0, "UInt", Data.Size, "Ptr", Data))
				{
				if (DllCall("version\VerQueryValueW", "Ptr", Data, "Str", "\VarFileInfo\Translation", "Ptr*", &Buf := 0, "UInt*", &Len := 0))
					{
					LangCP := Format("{:04X}{:04X}", NumGet(Buf + 0, "UShort"), NumGet(Buf + 2, "UShort"))
					FileInfo := Map()
					if (DllCall("version\VerQueryValueW", "Ptr", Data, "Str", "\StringFileInfo\" . LangCP . "\" . Property, "Ptr*", &Buf, "UInt*", &Len))
						{
						FileInfo[Property] := StrGet(Buf, Len, "UTF-16")
						return Trim(FileInfo[Property])
						}
					}
				}
			}
		}
    }

; --------------------------------------------------------------- Function for retrieving the AppxManifest.xml file content from an App-V (*.appv) or MSIX (*.msix) package file ---------------------------------------------------------------
Read_AppxManifest(AppxFile := "")
	{
	if FileExist(AppxFile)
		{
		FileMove AppxFile, AppxFile . ".zip"
		psh := ComObject("Shell.Application")
		psh.Namespace(A_Temp).CopyHere(psh.Namespace(AppxFile . ".zip").items.Item("AppxManifest.xml"), 4|16)
		FileMove AppxFile  . ".zip", AppxFile
		While !FileExist(A_Temp . "\AppxManifest.xml")
			Sleep 50
		AppxManifestXml := FileRead(A_Temp . "\AppxManifest.xml")
		FileDelete A_Temp . "\AppxManifest.xml"
		return AppxManifestXml
		}
	}

; --------------------------------------------------------------- Function for retrieving property values using the MSI file name and the property name ---------------------------------------------------------------
WindowsInstaller(Msifile := "", Property := "")
	{
	if FileExist(Msifile)
		{
		installer := ComObject("WindowsInstaller.Installer")
		if !database := installer.OpenDatabase(Msifile, 0)
			return A_LastError
		view := database.OpenView("SELECT `Value` FROM `Property` WHERE `Property` = '" Property "'")
		view.Execute
		record := view.Fetch
		return record.StringData(1)
		}
  	}

; --------------------------------------------------------------- Function for directly installing or removing a Windows Installer MSI file ---------------------------------------------------------------
MsiInstallProduct(PackagePath := "", CommandLine := "", LogFile := "")
	{
	Installer := ComObject("WindowsInstaller.Installer")
	if LogFile
		Installer.EnableLog "voicewarmup", LogFile
	Installer.UILevel := 2
	return DllCall("Msi.dll\MsiInstallProductW", "Wstr", PackagePath, "Wstr", CommandLine)
	}

; --------------------------------------------------------------- Function for creating a scheduled task at next Logon (default) or Boot ---------------------------------------------------------------
CreateTask(TaskName := "", Description := "", Path := "", Arguments := "", ExecutionTimeLimit := "12", TriggerType := "Logon", Battery := "")
	{
	Local ActionTypeExec := 0
	Local LogonType := 5
	Local TaskCreateOrUpdate := 6
	If TriggerType = "Logon"
		TriggerType := 9
	If TriggerType = "Boot"
		TriggerType := 8
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
	if Battery
		settings.DisallowStartifOnBatteries := False
	else
		settings.DisallowStartifOnBatteries := True
	settings.StopifGoingOnBatteries := False
	settings.AllowHardTerminate := True
	settings.StartWhenAvailable := False
	settings.RunOnlyifNetworkAvailable := False
	settings.Enabled := True
	settings.Hidden := False
	settings.RunOnlyifIdle := False
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
	try rootFolder.RegisterTaskDefinition(TaskName, taskDefinition, TaskCreateOrUpdate , "", "", 4)
	catch as err
		return err.Message
	else
		return 0
  	}

; --------------------------------------------------------------- Function for removing a scheduled task using the task name ---------------------------------------------------------------
DeleteTask(TaskName := "")
	{
	service := ComObject("Schedule.Service")
	service.Connect()
	rootFolder := service.GetFolder("\")
	try rootFolder.DeleteTask(TaskName, 0)
	catch as err
		return err.Message
	else
		return 0
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
		if (m[1])
			out .= (env := EnvGet(m[2])) ? env : m[1]
		else
 			out .= m[3]
		}
	return out SubStr(Str, spo)
	}

; --------------------------------------------------------------- Function for comparing two (version) numbers. if the CompareValue is higher the result will be greater. will return equal, greater or smaller depending on the (version) number values ---------------------------------------------------------------
CompareVersionString(CheckValue := 0, CompareValue := 0)
	{
	Result := VerCompare(CompareValue, CheckValue)
	if Result = 0
		return "Equal"
	if Result = 1
		return "Greater"
	if Result = -1
		return "Smaller"
	}

; --------------------------------------------------------------- Function for creating a new (user) process using the token of an existing running process name. (if wait is 1 the process exit code will return or else, the PID will return) ---------------------------------------------------------------
CreateProcessAsUser(cmd := "", args := "", workingDir := "", processName := "", wait := 0) 
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
	if !(ProcessId := ProcessExist(processName))
		{
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t[ CreateProcessAsUser ] Process not found: " processName "`r`n" , LogFile
		return -1
		}
	DllCall("ProcessIdToSessionId", "Uint", ProcessId, "Ptr*", &SessionId := 0)
	if !SessionId
		{
		FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t[ CreateProcessAsUser ] ProcessIdToSessionId function failed with error: " A_LastError "`r`n" , LogFile
		return -1
		}
	DllCall("wtsapi32\WTSQueryUserToken", "Int", SessionId, "Ptr*", &hToken := 0)
	DllCall("CloseHandle", "Ptr", SessionId)
	DllCall("Userenv.dll\CreateEnvironmentBlock", "Ptr*", &lpEnvironment := 0, "Ptr", hToken, "Int", 0)
	if args
		cmd := cmd " " args
	if !DllCall("Advapi32.dll\CreateProcessAsUserW", "Ptr", hToken, "Ptr", 0, "WStr", cmd, "Int", 0, "Int", 0, "Int", 0, "Int", 0x00000400, "Ptr", lpEnvironment, "WStr", workingDir, "Ptr", startupInfo, "Ptr", processInfo)
		{
		if FileExist(LogFile)
			FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t[ CreateProcessAsUser ] CreateProcessAsUserW function failed with error: " A_LastError "`r`n" , LogFile
		DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
		DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
		return -2
		}
	if wait = 1
		{
		loop
			{
			DllCall("GetExitCodeProcess", "Ptr", NumGet(processInfo, 0, "Ptr"), "UIntP" , &exitCode := 0)
			if exitCode = 259
				Sleep 500
			else
				break
			}
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, A_PtrSize, "Ptr"))
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, 0, "Ptr"))
		DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
		DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
		return exitCode
		}
	else
		{
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, A_PtrSize, "Ptr"))
		DllCall("kernel32.dll\CloseHandle", "Ptr", NumGet(processInfo, 0, "Ptr"))
		DllCall("Userenv.dll\DestroyEnvironmentBlock", "Ptr", lpEnvironment)
		DllCall("kernel32.dll\CloseHandle", "Ptr", hToken)
		return NumGet(processInfo, 2 * A_PtrSize, "Int")
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
				if id != A_PID
					{
					FoundSessions++
					if Stop = 1
						{
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStop notification process with Process Id: " id "`r`n" , LogFile
						ProcessClose id
						}
					}
				}
			}
		}
	DllCall("FreeLibrary", "Ptr", hModule)
	return FoundSessions
	}

; --------------------------------------------------------------- Function for converting the number of seconds to the mm:ss time format ---------------------------------------------------------------
FormatSeconds(NumberOfSeconds) 
	{
	time := 19990101
	time := DateAdd(time, NumberOfSeconds, "Seconds")
	return FormatTime(time, "mm:ss")
	}

; --------------------------------------------------------------- Function for setting session environment variables ---------------------------------------------------------------
SetEnv(Section)
	{
	Global Inifile
	Found := 0
	Counter := 0
	Loop read, Inifile
		{
		if Found = 1
			{
			if A_LoopReadLine
				{
				if InStr(A_LoopReadLine, "=")
					{
					EnvSetString := StrSplit(A_LoopReadLine, "=")
					EnvSetName := EnvSetString[1]
					EnvSetValue := EnvSetString[2]
					EnvSetValue := TransForm(EnvSetValue)
					EnvSet EnvSetName, EnvSetValue
					Counter++
					}
				else
					break
				}
			}
		if A_LoopReadLine = "[" Section "]"
			Found := 1
		}
	return Counter
	}

; --------------------------------------------------------------- psScript Manager Class for directly executing Powershell code ---------------------------------------------------------------
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
			try FileDelete SetInstallDir . "\psScript.dll"
		return 0
		}
	}

; --------------------------------------------------------------- WinService Class for controling Windows Services ---------------------------------------------------------------
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
	static State(ServiceName, textResult := false)
		{
		scm := DllCall("advapi32\OpenSCManagerW", "Ptr", 0, "Ptr", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32\OpenServiceW", "Ptr", scm, "Str", ServiceName, "UInt", this.SERVICE_QUERY_STATUS)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
			return err
			}
		buf := Buffer(28, 0)
		ok := DllCall("advapi32\QueryServiceStatus", "Ptr", hSvc, "Ptr", buf.ptr)
		DllCall("advapi32\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
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
	static Start(ServiceName)
		{
		scm := DllCall("advapi32\OpenSCManagerW", "Ptr", 0, "Ptr", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_START)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
			return err
			}
		ok := DllCall("advapi32\StartServiceW", "Ptr", hSvc, "UInt", 0, "Ptr", 0)
		err := ok ? ok : A_LastError
		DllCall("advapi32\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
		return err
		}
	static Stop(ServiceName)
		{
		scm := DllCall("advapi32\OpenSCManagerW", "Ptr", 0, "Ptr", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_STOP)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
			return err
			}
		status := Buffer((A_PtrSize = 4) ? 28 : 32, 0)
		ok := DllCall("advapi32\ControlService", "Ptr", hSvc, "UInt", 1, "Ptr", status.ptr)
		err := ok ? ok : A_LastError
		DllCall("advapi32\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
		return err
		}
	static Add(ServiceName, BinaryPath, StartType := "", DisplayName := "", ServiceStartName := "", Password := "")
		{
		StartType := (StartType = "Auto" || StartType = "Automatic") ? 0x2
				: (StartType = "Demand" || StartType = "OnDemand") ? 0x3
				: 0x4
		scm := DllCall("advapi32\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CREATE_SERVICE)
		if ServiceStartName
			svc := DllCall("advapi32\CreateServiceW","Ptr", scm, "Ptr", StrPtr(ServiceName), "Ptr", StrPtr(DisplayName ? DisplayName : ServiceName), "UInt", this.SERVICE_ALL_ACCESS, "UInt", this.SERVICE_WIN32_OWN_PROCESS, "UInt", StartType, "UInt", this.SERVICE_ERROR_NORMAL, "Ptr", StrPtr(BinaryPath), "Ptr", 0, "UInt", 0, "Ptr", 0, "Ptr", StrPtr(ServiceStartName), "Ptr", StrPtr(Password))
		else
			svc := DllCall("advapi32\CreateServiceW","Ptr", scm, "Ptr", StrPtr(ServiceName), "Ptr", StrPtr(DisplayName ? DisplayName : ServiceName), "UInt", this.SERVICE_ALL_ACCESS, "UInt", this.SERVICE_WIN32_OWN_PROCESS, "UInt", StartType, "UInt", this.SERVICE_ERROR_NORMAL, "Ptr", StrPtr(BinaryPath), "Ptr", 0, "UInt", 0, "Ptr", 0, "Ptr", 0, "Ptr", 0)
		result := A_LastError ? svc "," A_LastError : 1
		DllCall("advapi32\CloseServiceHandle", "Ptr", svc)
		DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
		return result
		}
	static Delete(ServiceName)
		{
		scm := DllCall("advapi32\OpenSCManagerW", "Int", 0, "Int", 0, "UInt", this.SC_MANAGER_CONNECT)
		hSvc := DllCall("advapi32\OpenServiceW", "Ptr", scm, "Ptr", StrPtr(ServiceName), "UInt", this.SERVICE_ALL_ACCESS)
		if !hSvc
			{
			err := A_LastError
			DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
			return err
			}
		ok := DllCall("advapi32\DeleteService", "Ptr", hSvc)
		err := ok ? ok : A_LastError
		DllCall("advapi32\CloseServiceHandle", "Ptr", hSvc)
		DllCall("advapi32\CloseServiceHandle", "Ptr", scm)
		return err
		}
	}

; --------------------------------------------------------------- Function for importing registry information content from a REG file. Set "UseTransform" to 1 for resolving all environment variables used inside the register information content to their actual values ---------------------------------------------------------------
SetReg(RegContent := "", UseTransform := 0)
	{
	Global LogFile
	Global UserScript
	if RegContent
		{
		StartTime := A_TickCount
		RegKey_created_success := 0
		RegKey_created_failed := 0
		RegKey_deleted_success := 0
		RegKey_deleted_failed := 0
		RegValue_created_success := 0
		RegValue_created_failed := 0
		RegValue_deleted_success := 0
		RegValue_deleted_failed := 0
		if InStr(RegContent, "\`r`n" A_Space A_Space)
			RegContent := StrReplace(RegContent, "\`r`n" A_Space A_Space, "")
		KeyName := ""
		loop parse, RegContent, "`n", "`r"
			{
			if UseTransform = 1
				LoopReadLine := Transform(A_LoopField)
			else
				LoopReadLine := Trim(A_LoopField)				
			if SubStr(LoopReadLine, 1, 1) = "["
				{
				KeyName := SubStr(LoopReadLine, 2, StrLen(LoopReadLine) - 2)
				if SubStr(KeyName, 1, 1) = "-"
					{
					KeyName := SubStr(KeyName, 2, StrLen(KeyName) - 1)
					try 
						RegDeleteKey KeyName
					catch as err
						RegKey_deleted_failed++
					else
						RegKey_deleted_success++
					}
				else
					{
					try
						RegCreateKey KeyName
					catch as err
						{
						RegKey_created_failed++
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate registry key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
						else
							if UserScript
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate registry key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
						}
					else
						RegKey_created_success++
					}
				}
			else
				{
				if InStr(LoopReadLine, "\`"")
					LoopReadLine := StrReplace(LoopReadLine, "\`"", "`t")
				if InStr(LoopReadLine, "\\")
					LoopReadLine := StrReplace(LoopReadLine, "\\", "\")
				if (InStr(LoopReadLine, "`"=-",, 1) OR InStr(LoopReadLine, "`"=dword:",, 1) OR InStr(LoopReadLine, "`"=hex:",, 1) OR InStr(LoopReadLine, "`"=hex(2):",, 1) OR InStr(LoopReadLine, "`"=hex(7):",, 1) OR InStr(LoopReadLine, "`"=hex(b):",, 1))
					{
					if InStr(LoopReadLine, "`"=-",, 1)
						{
						ValueType := ""
						temp_array := StrSplit(LoopReadLine, "`"=-")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						temp_array := ""
						try 
							RegDelete KeyName, ValueName
						catch as err
							RegValue_deleted_failed++
						else
							RegValue_deleted_success++
						}
					if InStr(LoopReadLine, "`"=hex:",, 1)
						{
						ValueType := "REG_BINARY"
						temp_array := StrSplit(LoopReadLine, "`"=hex:")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						Value := temp_array[2]
						temp_array := ""
						Value := StrReplace(Value := StrUpper(Value), ",")
						try
							RegWrite Value, ValueType, KeyName, ValueName
						catch as err
							{
							RegValue_created_failed++
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
							}
						else
							RegValue_created_success++
						}
					if InStr(LoopReadLine, "`"=dword:",, 1)
						{
						ValueType := "REG_DWORD"
						temp_array := StrSplit(LoopReadLine, "`"=dword:")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						Value := temp_array[2]
						temp_array := ""
						Value := "0x" Value
						try
							RegWrite Value, ValueType, KeyName, ValueName
						catch as err
							{
							RegValue_created_failed++
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
							}
						else
							RegValue_created_success++
						}
					if InStr(LoopReadLine, "`"=hex(b):",, 1)
						{
						ValueType := "REG_QWORD"
						temp_array := StrSplit(LoopReadLine, "`"=hex(b):")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						Value := temp_array[2]
						temp_array := ""
						ConvertedValue := ""
						Loop parse, Value, ","
							ConvertedValue := A_LoopField . ConvertedValue
						ConvertedValue := "0x" . ConvertedValue
						Value := ConvertedValue + 0
						ConvertedValue := ""
						loop parse, KeyName, "\"
							{
							if (A_LoopField = "HKCR" or A_LoopField = "HKEY_CLASSES_ROOT" or A_LoopField = "HKCU" or A_LoopField = "HKEY_CURRENT_USER" or A_LoopField = "HKLM" or A_LoopField = "HKEY_LOCAL_MACHINE" or A_LoopField = "HKU" or A_LoopField = "HKEY_USERS" or A_LoopField = "HKCC" or A_LoopField = "HKEY_CURRENT_CONFIG")
								{
								RootKey := A_LoopField
								SubKey := StrReplace(KeyName, A_LoopField . "\", "")
								Break
								}
							}
						try
							Reg_QWORD.Write(RootKey, SubKey, ValueName, Value)
						catch as err
							{
							RegValue_created_failed++
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (REG_QWORD) registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (REG_QWORD) registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
							}
						else
							RegValue_created_success++
						}
					if InStr(LoopReadLine, "`"=hex(7):",, 1)
						{
						ValueType := "REG_MULTI_SZ"
						temp_array := StrSplit(LoopReadLine, "`"=hex(7):")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						Value := temp_array[2]
						temp_array := ""
						ConvertedValue := ""
						ZeroCounter := 0
						loop parse, Value, ","
							{
							if ZeroCounter = 2
								ConvertedValue := ConvertedValue . "`n"
							if A_LoopField = "00"
								ZeroCounter++
							else
								{
								ZeroCounter := 0
								ConvertedValue := ConvertedValue . Chr("0x" A_LoopField)
								}
							}
						Value := ConvertedValue
						ConvertedValue := ""
						ZeroCounter := ""
						try
							RegWrite Value, ValueType, KeyName, ValueName
						catch as err
							{
							RegValue_created_failed++
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
							}
						else
							RegValue_created_success++
						}
					if InStr(LoopReadLine, "`"=hex(2):",, 1)
						{
						ValueType := "REG_EXPAND_SZ"
						temp_array := StrSplit(LoopReadLine, "`"=hex(2):")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						Value := temp_array[2]
						temp_array := ""
						ConvertedValue := ""
						Loop parse, Value, ","
							{
							If A_LoopField != "00"
								ConvertedValue := ConvertedValue . Chr("0x" A_LoopField)
							}
						Value := ConvertedValue
						ConvertedValue := ""
						try
							RegWrite Value, ValueType, KeyName, ValueName
						catch as err
							{
							RegValue_created_failed++
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
							}
						else
							RegValue_created_success++
						}
					}
				else
					{
					if InStr(LoopReadLine, "`"=`"",, 1)
						{
						ValueType := "REG_SZ"
						temp_array := StrSplit(LoopReadLine, "`"=`"")
						ValueName := SubStr(temp_array[1], 2, StrLen(temp_array[1]) - 1)
						if InStr(ValueName, "`t")
							ValueName := StrReplace(ValueName, "`t", "`"")
						Value := SubStr(temp_array[2], 1, StrLen(temp_array[2]) - 1)
						if InStr(Value, "`t")
							Value := StrReplace(Value, "`t", "`"")
						temp_array := ""
						try
							RegWrite Value, ValueType, KeyName, ValueName
						catch as err
							{
							RegValue_created_failed++
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate (" ValueType ") registry value name `"" ValueName "`" in key `"" KeyName "`" failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
							}
						else
							RegValue_created_success++
						}
					if SubStr(LoopReadLine, 1, 2) = "@="
						{
						ValueType := "REG_SZ"
						ValueName := ""
						if LoopReadLine = "@=-"
							{
							try 
								RegDelete KeyName, ValueName
							catch as err
								RegValue_deleted_failed++
							else
								RegValue_deleted_success++
							}
						else
							{
							temp_array := StrSplit(LoopReadLine, "@=")
							Value := SubStr(temp_array[2], 2, StrLen(temp_array[2]) - 2)
							if InStr(Value, "`t")
								Value := StrReplace(Value, "`t", "`"")
							temp_array := ""
							try
								RegWrite Value, ValueType, KeyName, ValueName
							catch as err
								{
								RegValue_created_failed++
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate default (" ValueType ") registry value `"" Value "`" in key `"" KeyName "`" failed with Error: " . err.Message "`r`n" , LogFile
								else
									if UserScript
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tCreate default (" ValueType ") registry value `"" Value "`" in key `"" KeyName "`" failed with Error: " . err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
								}
							else
								RegValue_created_success++
							}
						}
					}
				}
			}
		ElapsedTime := A_TickCount - StartTime
		return "Processed registry content in " ElapsedTime " milliseconds. Registry keys created: " RegKey_created_success " (failed: " RegKey_created_failed "). Registry keys removed: " RegKey_deleted_success " (failed: " RegKey_deleted_failed "). Registry values created: " RegValue_created_success " (failed: " RegValue_created_failed "). Registry values removed: " RegValue_deleted_success " (failed: " RegValue_deleted_failed ")."
		}
	else
		return "No registry import content found!"
	}

; --------------------------------------------------------------- Class for Reading and writing QWORD type registry values  ---------------------------------------------------------------
Class Reg_QWORD
	{
	Static Call(*) => 0
	Static CheckRoot(&Root) 
		{
    	Static RootKey := {HKCR: 0x80000000, HKEY_CLASSES_ROOT:   0x80000000,
        	               HKCU: 0x80000001, HKEY_CURRENT_USER:   0x80000001,
            	           HKLM: 0x80000002, HKEY_LOCAL_MACHINE:  0x80000002,
                	       HKU:  0x80000003, HKEY_USERS:          0x80000003,
                    	   HKCC: 0x80000005, HKEY_CURRENT_CONFIG: 0x80000005}
		if !(Type(Root) = "String")
			Throw TypeError("Expected a string but got a " Type(Root), -1, "Root")
		if !RootKey.HasProp(Root)
			Throw Error("Invalid parameter Root!", -1, IsObject(Root) ? "*OBJECT*" : String(Root))
		Root := RootKey.%Root%
		return True
		}
	Static SetView() 
		{
		; KEY_WOW64_32KEY = 0x0200, KEY_WOW64_64KEY = 0x0100
		Switch A_RegView 
			{
			Case 32: return 0x0200
			Case 64: return 0x0100
			Default: return 0
			}
		}
	Static Read(Root, SubKey, Value := "") 
		{
		; RRF_RT_QWORD = 0x00000048
		if !This.CheckRoot(&Root)
			return 0
		Local QWORD := 0, Size := 8
		Local View := This.SetView()
		Local Ret := DllCall("Advapi32.dll\RegGetValueW",
							"Ptr",     Root,
							"Str",     SubKey,
							"Str",     Value,
							"UInt",    0x00000048 | View,
							"Ptr",     0,
							"UInt64*", &QWORD,
							"UInt*",   &Size,
							"UInt")
		if !(Ret = 0)
			Throw Error("RegGetValueW failed with error " . Ret, -1)
		return QWORD
		}
	Static Write(Root, SubKey, ValueName, Value) 
		{
		; KEY_WRITE = 0x20006, REG_QWORD = 11
		if !This.CheckRoot(&Root)
			return 0
		Local Disp := 0, HKEY := 0, Ret := 0
		Local View := This.SetView()
		Ret := DllCall("Advapi32.dll\RegCreateKeyExW", "Ptr",    Root,
							"Str",    &SubKey,
							"UInt",   0,
							"Ptr",    0,
							"UInt",   0,
							"UInt",   0x20006 | View,
							"Ptr",    0,
							"PtrP",   &HKEY,
							"UInt*",  &Disp)
		if !(Ret = 0)
			Throw Error("RegCreateKeyExW failed with error " . Ret, -1)
		Ret := DllCall("Advapi32.dll\RegSetValueExW", "Ptr",    HKEY,
							"Str",    ValueName,
							"UInt",   0,
							"UInt",   11,
							"Int64*", Value = "" ? 0 : Value,
							"UInt",   8)
		DllCall("Advapi32.dll\RegCloseKey", "Ptr", HKEY)
		if !(Ret = 0)
			Throw Error("RegSetValueExW failed with error " . Ret, -1)
		return Disp
		}
	}

; --------------------------------------------------------------- Function for executing PreLaunch, PostExit and User Script items ---------------------------------------------------------------
Script(Section)
	{
	Global Inifile
	Global LogFile
	Global ScriptSections
	Global cpauPID
	Global SystemPipeName
	Global UserPipeName
	Global UserScript
	if UserScript
		ProgressBar := 0
	Found := 0
	Counter := 0
	if Instr(ScriptSections, Section "`r`n")
		{
		if FileExist(LogFile)
			FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tScript section " Section " is already executed. Each script section can only be executed once for each PreLaunch, PostExit and UserScript script section. Please check the if/else statements.`r`n" , LogFile
		return Counter
		}
	ScriptSections := ScriptSections . Section "`r`n"
	Loop read, Inifile
		{
		if Found = 1
			{
			if A_LoopReadLine
				{
				if InStr(A_LoopReadLine, ",")
					{
					if UserScript
						{
						NotifyProgress.Value := ProgressBar++
						if ProgressBar = 100
							ProgressBar := 0
						}
					ScriptItemString := StrSplit(A_LoopReadLine, ",")
					ScriptItemName := ScriptItemString[1]
					ScriptItemName := Trim(ScriptItemName)
					if (ScriptItemName = "ifEqual" or ScriptItemName = "ifNotEqual" or ScriptItemName = "ifGreater" or ScriptItemName = "ifGreaterOrEqual" or ScriptItemName = "ifNotGreater" or ScriptItemName = "ifNotGreaterOrEqual" or ScriptItemName = "ifSmaller" or ScriptItemName = "ifSmallerOrEqual" or ScriptItemName = "ifNotSmaller" or ScriptItemName = "ifNotSmallerOrEqual")
						{
						CheckVariable := ""
						CheckValue := ""
						GotoifSection := ""
						GotoelseSection := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								CheckVariable := ScriptItemString[2]
							if A_Index = 3
								CheckValue := ScriptItemString[3]
							if A_Index = 4
								GotoifSection := ScriptItemString[4]
							if A_Index = 5
								GotoelseSection := ScriptItemString[5]
							if A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						CheckVariable := Trim(CheckVariable)
						CheckValue := Transform(CheckValue)
						GotoifSection := Trim(GotoifSection)
						GotoelseSection := Trim(GotoelseSection)
						UserDescription := TransForm(UserDescription)
						if ScriptItemName = "ifEqual"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							CheckValue := StrReplace(CheckValue, A_Space . "OR" . A_Space, "`n|`t")
							CheckValue := "`n|`t" . CheckValue
							Result := False
							Loop parse, CheckValue, "`n"
								{
								if A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									if Check_array[1] = "|"
										{
										if GetValue = Check_array[2]
											Result := True
										}
									}
								}
							if Result = True
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotEqual"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							CheckValue := StrReplace(CheckValue, A_Space . "OR" . A_Space, "`n|`t")
							CheckValue := "`n|`t" . CheckValue
							Result := False
							Loop parse, CheckValue, "`n"
								{
								if A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									if Check_array[1] = "|"
										{
										if GetValue != Check_array[2]
											Result := True
										}
									}
								}
							if Result = True
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifGreater"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if Result = "Greater"
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifGreaterOrEqual"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if (Result = "Greater" or Result = "Equal")
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotGreater"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotGreater statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if (Result = "Smaller" or Result = "Equal")
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotGreaterOrEqual"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotGreaterOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if Result = "Smaller"
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifSmaller"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if Result = "Smaller"
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifSmallerOrEqual"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if (Result = "Smaller" or Result = "Equal")
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotSmaller"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotSmaller statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if (Result = "Greater" or Result = "Equal")
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotSmallerOrEqual"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotSmallerOrEqual statement for variable name '" CheckVariable "' with value '" CheckValue "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							GetValue := EnvGet(CheckVariable)
							Result := CompareVersionString(CheckValue, GetValue)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCheck value '" CheckValue "' with value '" GetValue "'. The last value will have the following result: " Result "`r`n" , LogFile
							if Result = "Greater"
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						}
					if (ScriptItemName = "ifExist" or ScriptItemName = "ifNotExist" or ScriptItemName = "ifProcessExist" or ScriptItemName = "ifNotProcessExist")
						{
						CheckFileFolderRegKey := ""
						GotoifSection := ""
						GotoelseSection := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								CheckFileFolderRegKey := ScriptItemString[2]
							if A_Index = 3
								GotoifSection := ScriptItemString[3]
							if A_Index = 4
								GotoelseSection := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						CheckFileFolderRegKey := TransForm(CheckFileFolderRegKey)
						GotoifSection := Trim(GotoifSection)
						GotoelseSection := Trim(GotoelseSection)
						UserDescription := TransForm(UserDescription)
						if ScriptItemName = "ifExist"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "OR" . A_Space, "`n|`t")
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "AND" . A_Space, "`n&`t")
							CheckFileFolderRegKey := "`n|`t" . CheckFileFolderRegKey
							Result := False
							Loop parse, CheckFileFolderRegKey, "`n"
								{
								if A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									if Check_array[1] = "|"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										if (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											if RegKeyExists(Path_Array[1], Check_array[2]) = 1
												Result := True
											}
										else
											{
											SplitPath Check_array[2],,, &Extension
											if Extension
												{
												if FileExist(Check_array[2])
													Result := True
												}
											else
												{
												if DirExist(Check_array[2])
													Result := True
												}
											}
										}
									if Check_array[1] = "&"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										if (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											if RegKeyExists(Path_Array[1], Check_array[2]) = 0
												Result := false
											}
										else
											{
											SplitPath Check_array[2],,, &Extension
											if Extension
												{
												if Not FileExist(Check_array[2])
													Result := False
												}
											else
												{
												if Not DirExist(Check_array[2])
													Result := False
												}
											}
										}
									}
								}
							if Result = true
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotExist"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotExist statement for file, folder (path) or registry key name '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "OR" . A_Space, "`n|`t")
							CheckFileFolderRegKey := StrReplace(CheckFileFolderRegKey, A_Space . "AND" . A_Space, "`n&`t")
							CheckFileFolderRegKey := "`n|`t" . CheckFileFolderRegKey
							Result := False
							Loop parse, CheckFileFolderRegKey, "`n"
								{
								if A_LoopField
									{
									Check_array := StrSplit(A_LoopField, "`t")
									if Check_array[1] = "|"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										if (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											if RegKeyExists(Path_Array[1], Check_array[2]) = 0
												Result := True
											}
										else
											{
											SplitPath Check_array[2],,, &Extension
											if Extension
												{
												if Not FileExist(Check_array[2])
													Result := True
												}
											else
												{
												if Not DirExist(Check_array[2])
													Result := True
												}
											}
										}
									if Check_array[1] = "&"
										{
										Path_Array := StrSplit(Check_array[2], "\",)
										if (Path_Array[1] = "HKEY_CLASSES_ROOT" or Path_Array[1] = "HKCR" or Path_Array[1] = "HKEY_CURENT_USER" or Path_Array[1] = "HKCU" or Path_Array[1] = "HKEY_LOCAL_MACHINE" or Path_Array[1] = "HKLM" or Path_Array[1] = "HKEY_USERS" or Path_Array[1] = "HKU" or Path_Array[1] = "HKEY_CURRENT_CONFIG" or Path_Array[1] = "HKCC")
											{
											Check_array[2] := StrReplace(Check_array[2], Path_Array[1] . "\", "")
											if RegKeyExists(Path_Array[1], Check_array[2]) = 1
												Result := false
											}
										else
											{
											SplitPath Check_array[2],,, &Extension
											if Extension
												{
												if FileExist(Check_array[2])
													Result := False
												}
											else
												{
												if DirExist(Check_array[2])
													Result := False
												}
											}
										}
									}
								}
							if Result = true
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifProcessExist"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifProcessExist statement for process (ID) name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							if (PID := ProcessExist(CheckFileFolderRegKey))
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						if ScriptItemName = "ifNotProcessExist"
							{
							if GotoelseSection
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'.`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotProcessExist statement for process (ID) name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else goto script section '" GotoelseSection "'."
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tUsing ifNotProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue...`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Using ifNotProcessExist statement for process name or ID '" CheckFileFolderRegKey "'. When true, goto script section '" GotoifSection "'. else, continue..."
								}
							if Not (PID := ProcessExist(CheckFileFolderRegKey))
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoifSection " ]`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
									   		ShowInfo.Value := "Performing script section: " GotoifSection
								StatementCounter := Script(GotoifSection)
								}
							else
								{
								if GotoelseSection
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing script section: " GotoelseSection " ]`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Performing script section: " GotoelseSection
									StatementCounter := Script(GotoelseSection)
									}
								}
							}
						}
					if ScriptItemName = "FileCopy"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Source := ScriptItemString[2]
							if A_Index = 3
								Destination := ScriptItemString[3]
							if A_Index = 4
								OverWrite := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FileCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						if FileExist(Source)
							{
							try
								FileCopy Source, Destination, OverWrite
							catch as Err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileCopy failed for " err.Extra " files.`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileCopy failed for " err.Extra " files.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "FileCopy failed for " err.Extra " files."
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileCopy successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FileCopy successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "FileMove"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Source := ScriptItemString[2]
							if A_Index = 3
								Destination := ScriptItemString[3]
							if A_Index = 4
								OverWrite := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileMove (File rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FileMove (File rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						if FileExist(Source)
							{
							try
								FileMove Source, Destination, OverWrite
							catch as Err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileMove failed for " err.Extra " files.`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileMove failed for " err.Extra " files.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "FileMove failed for " err.Extra " files."
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileMove successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FileMove successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "FolderCopy"
						{
						Source := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Source := ScriptItemString[2]
							if A_Index = 3
								Destination := ScriptItemString[3]
							if A_Index = 4
								OverWrite := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FolderCopy from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						if (DirExist(Source) or FileExist(Source))
							{
							try
								DirCopy Source, Destination, OverWrite
							catch as err
								{
								if FileExist(LogFile)
									{
									if OverWrite = 1
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`r`n" , LogFile
									else
										if !UserDescription
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderCopy successfully executed.`r`n" , LogFile
									}
								else
									if UserScript
										{
										if OverWrite = 1
											{
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderCopy failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
											ShowInfo.Value := "FolderCopy failed with Error: " err.Message
											Sleep 10000
											}
										else
											if !UserDescription
												ShowInfo.Value := "FolderCopy successfully executed."
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderCopy successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FolderCopy successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "FolderMove"
						{
						Source := ""
						Destination := ""
						OverWrite := "R"
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Source := ScriptItemString[2]
							if A_Index = 3
								Destination := ScriptItemString[3]
							if A_Index = 4
								OverWrite := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Source := TransForm(Source)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderMove (Folder rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FolderMove (Folder rename) from source '" Source "' to destination '" Destination "' (OverWrite=" OverWrite ")"
						if DirExist(Source)
							{
							try
								DirMove Source, Destination, OverWrite
							catch as err
								{
								if FileExist(LogFile)
									{
									if OverWrite = 1
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderMove failed with Error: " err.Message "`r`n" , LogFile
									else
										if !UserDescription
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderMove successfully executed.`r`n" , LogFile
									}
								else
									if UserScript
										{
										if OverWrite = 1
											{
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderMove failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
											ShowInfo.Value := "FolderMove failed with Error: " err.Message
											Sleep 10000
											}
										else
											if !UserDescription
												ShowInfo.Value := "FolderMove successfully executed."
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderMove successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FolderMove successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "FolderCreate"
						{
						Destination := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Destination := ScriptItemString[2]
							if A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						Destination := TransForm(Destination)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderCreate: " Destination "`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FolderCreate: " Destination
						if Not FileExist(Destination)
							{
							try
								DirCreate Destination
							catch as Err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderCreate failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderCreate failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "FolderCreate failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderCreate successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FolderCreate successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" Destination "' already exists."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "FileDelete"
						{
						FilePattern := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								FilePattern := ScriptItemString[2]
							if A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						FilePattern := TransForm(FilePattern)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileDelete: " FilePattern "`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FileDelete: " FilePattern
						if FileExist(FilePattern)
							{
							try
								FileDelete FilePattern
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileDelete failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "FileDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileDelete successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FileDelete successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" FilePattern "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" FilePattern "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" FilePattern "' not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "FolderDelete"
						{
						DirName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								DirName := ScriptItemString[2]
							if A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						DirName := TransForm(DirName)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderDelete: " DirName "`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FolderDelete: " DirName
						if DirExist(DirName)
							{
							try
								DirDelete DirName, 1
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderDelete failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFolderDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "FolderDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFolderDelete successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "FolderDelete successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" DirName "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" DirName "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" DirName "' not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "Set"
						{
						SetName := ""
						SetValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								SetName := ScriptItemString[2]
							if A_Index = 3
								SetValue := ScriptItemString[3]
							if A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						SetName := Trim(SetName)
						SetValue := TransForm(SetValue)
						UserDescription := TransForm(UserDescription)
						if SetName
							{
							if SetValue
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSet session environment variable name '" SetName "' with value: " SetValue "`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Set session environment variable name '" SetName "' with value: " SetValue
								try
									EnvSet SetName, SetValue
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSession environment variable failed to set for '" SetName "' with value '" SetValue "' Error: " err.Message "`r`n" , LogFile
									else
										if UserScript
											{
											ShowInfo.Value := "Session environment variable failed to set for '" SetName "' with value '" SetValue "' Error: " err.Message
											Sleep 10000
											}
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSession environment variable successfully set for '" SetName "' with value: " SetValue "`r`n" , LogFile
									else
										if UserScript
											{
											if !UserDescription
												ShowInfo.Value := "Session environment variable successfully set for '" SetName "' with value: " SetValue
											}
									}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSet remove session environment variable name: " SetName "`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "Set remove session environment variable name: " SetName
								try
									EnvSet SetName
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSession environment variable failed to remove for '" SetName "' Error: " err.Message "`r`n" , LogFile
									else
										if UserScript
											{
											ShowInfo.Value := "Session environment variable failed to remove for '" SetName "' Error: " err.Message
											Sleep 10000
											}
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSession environment variable successfully removed for variable name: " SetName "`r`n" , LogFile
									else
										if UserScript
											{
											if !UserDescription
												ShowInfo.Value := "Session environment variable successfully removed for variable name: " SetName
											}
									}
								}
							}
						}
					if ScriptItemName = "EnvSet"
						{
						EnvSetName := ""
						EnvSetValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								EnvSetName := ScriptItemString[2]
							if A_Index = 3
								EnvSetValue := ScriptItemString[3]
							if A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						EnvSetName := Trim(EnvSetName)
						EnvSetValue := TransForm(EnvSetValue)
						UserDescription := TransForm(UserDescription)
						if EnvSetName
							{
							if EnvSetValue
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvSet setting system environment variable name '" EnvSetName "' with value: " EnvSetValue "`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "EnvSet setting user environment variable name '" EnvSetName "' with value: " EnvSetValue
								if UserScript
									{
									try
										RegWrite EnvSetValue, "REG_EXPAND_SZ", "HKEY_CURRENT_USER\Environment", EnvSetName
									catch as err
										{
										ShowInfo.Value := "User environment variable registry (HKEY_CURRENT_USER\Environment) failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message
										Sleep 10000
										}
									else
										{
										try
											EnvSet EnvSetName, EnvSetValue
										catch as err
											{
											ShowInfo.Value := "User environment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message
											Sleep 10000
											}
										else
											{
											SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
											if !UserDescription
												ShowInfo.Value := "User environment variable successfully set for '" EnvSetName "' with value: " EnvSetValue
											}
										}
									}
								else
									{
									try
										RegWrite EnvSetValue, "REG_EXPAND_SZ", "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", EnvSetName
									catch as err
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSystem environment variable registry (HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment) failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
										}
									else
										{
										try
											EnvSet EnvSetName, EnvSetValue
										catch as err
											{
											if FileExist(LogFile)
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSystem environment variable failed to set for '" EnvSetName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
											}
										else
											{
											SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
											if FileExist(LogFile)
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSystem environment variable successfully set for '" EnvSetName "' with value: " EnvSetValue "`r`n" , LogFile
											}
										}
									}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvSet remove system environment variable name: " EnvSetName "`r`n" , LogFile
								else
									if UserScript
										if UserDescription
											ShowInfo.Value := UserDescription
										else
											ShowInfo.Value := "EnvSet remove user environment variable name: " EnvSetName
								if UserScript
									{
									EnvGetValue := RegRead("HKEY_CURRENT_USER\Environment", EnvSetName, 0)
									if EnvGetValue != 0
										{
										try
											RegDelete "HKEY_CURRENT_USER\Environment", EnvSetName
										catch as err
											{
											ShowInfo.Value := "User environment variable registry (HKEY_CURRENT_USER\Environment) failed to remove for '" EnvSetName "' Error: " err.Message
											Sleep 10000
											}
										else
											{
											try
												EnvSet EnvSetName
											catch as err
												{
												ShowInfo.Value := "User environment variable failed to remove for '" EnvSetName "' Error: " err.Message
												Sleep 10000
												}
											else
												{
												SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
												if !UserDescription
													ShowInfo.Value := "User environment variable successfully removed for variable name: " EnvSetName
												}
											}
										}
									else
										{
										ShowInfo.Value := "User environment variable '" EnvSetName "' not found."
										Sleep 10000										
										}
									}
								else
									{
									EnvGetValue := RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", EnvSetName, 0)
									if EnvGetValue != 0
										{
										try
											RegDelete "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment", EnvSetName
										catch as err
											{
											if FileExist(LogFile)
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSystem environment variable registry (HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Environment) failed to remove for '" EnvSetName "' Error: " err.Message "`r`n" , LogFile
											}
										else
											{
											try
												EnvSet EnvSetName
											catch as err
												{
												if FileExist(LogFile)
													FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSystem environment variable failed to remove for '" EnvSetName "' Error: " err.Message "`r`n" , LogFile
												}
											else
												{
												SendMessage(0x1A, 0, StrPtr("Environment"), 0xFFFF)
												if FileExist(LogFile)
													FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tSystem environment variable successfully removed for variable name: " EnvSetName "`r`n" , LogFile
												}
											}
										}
									else
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tSystem environment variable '" EnvSetName "' not found.`r`n" , LogFile
										}
									}
								}
							}
						}
					if ScriptItemName = "RegImport"
						{
						ReturnMsg := ""
						RegContent := ""
						RegFile := ""
						UseTransForm := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								RegFile := ScriptItemString[2]
							if A_Index = 3
								UseTransform := ScriptItemString[3]
							if A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						RegFile := TransForm(RegFile)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tImporting registry file '" RegFile "'. Use transform is set to: " UseTransForm "`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "Importing registry file: " RegFile
						if FileExist(RegFile)
							{
							ReturnMsg := SetReg(FileRead(RegFile), UseTransform)
							RegContent := ""
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" ReturnMsg "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegistry file '" RegFile "' imported. " ReturnMsg "`r`n" , A_AppData . "\Intune Win32 Launcher Register Import info.log", "UTF-16"
									if !UserDescription
										ShowInfo.Value := ReturnMsg
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegistry file not found. Skipping import.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegistry file '" RegFile "' not found. Skipping import.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "Registry file '" RegFile "' not found. Skipping import."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "RegRead"
						{
						KeyName := ""
						ValueName := ""
						DefaultValue := ""
						EnvSetValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								KeyName := ScriptItemString[2]
							if A_Index = 3
								ValueName := ScriptItemString[3]
							if A_Index = 4
								DefaultValue := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						KeyName := TransForm(KeyName)
						ValueName := Trim(ValueName)
						DefaultValue := TransForm(DefaultValue)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegRead registry item '" ValueName "' in '" KeyName "' with default value '" DefaultValue "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "RegRead registry item '" ValueName "' in '" KeyName "' with default value '" DefaultValue "'"
						try
							EnvSetValue := RegRead(KeyName, ValueName, DefaultValue)
						catch as err
							{
							if err.number = 1630
								{
								RootKey := ""
								SubKey := ""
								loop parse, KeyName, "\"
									{
									if (A_LoopField = "HKCR" or A_LoopField = "HKEY_CLASSES_ROOT" or A_LoopField = "HKCU" or A_LoopField = "HKEY_CURRENT_USER" or A_LoopField = "HKLM" or A_LoopField = "HKEY_LOCAL_MACHINE" or A_LoopField = "HKU" or A_LoopField = "HKEY_USERS" or A_LoopField = "HKCC" or A_LoopField = "HKEY_CURRENT_CONFIG")
										{
										RootKey := A_LoopField
										SubKey := StrReplace(KeyName, RootKey . "\", "")
										break
										}
									}
								try
									EnvSetValue := Reg_QWORD.Read(RootKey, SubKey, ValueName)
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegRead (REG_QWORD) failed with Error: " err.Message "`r`n" , LogFile
									else
										if UserScript
											{
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegRead (REG_QWORD) failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
											ShowInfo.Value := "RegRead failed with Error: " err.Message
											Sleep 10000
											}
									EnvSetValue := DefaultValue
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegRead successfully executed with value: " EnvSetValue "`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "RegRead successfully executed with value: " EnvSetValue
									try
										EnvSet ValueName, EnvSetValue
									catch as err
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
										else
											if UserScript
												{
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
												ShowInfo.Value := "Environment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message
												Sleep 10000
												}
										}
									else
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for '" ValueName "' with value: " EnvSetValue "`r`n" , LogFile
										else
											if UserScript
												if !UserDescription
													ShowInfo.Value := "Environment variable successfully set for '" ValueName "' with value: " EnvSetValue
										}
									}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegRead failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegRead failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "RegRead failed with Error: " err.Message
										Sleep 10000
										}
								EnvSetValue := DefaultValue
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegRead successfully executed with value: " EnvSetValue "`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "RegRead successfully executed with value: " EnvSetValue
							}
						try
							EnvSet ValueName, EnvSetValue
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "Environment variable failed to set for '" ValueName "' with value '" EnvSetValue "' Error: " err.Message
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for '" ValueName "' with value: " EnvSetValue "`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "Environment variable successfully set for '" ValueName "' with value: " EnvSetValue
							}
						}
					if ScriptItemName = "RegWrite"
						{
						ValueType := ""
						KeyName := ""
						ValueName := ""
						Value := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								ValueType := ScriptItemString[2]
							if A_Index = 3
								KeyName := ScriptItemString[3]
							if A_Index = 4
								ValueName := ScriptItemString[4]
							if A_Index = 5
								Value := ScriptItemString[5]
							if A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						ValueType := TransForm(ValueType)
						KeyName := TransForm(KeyName)
						ValueName := TransForm(ValueName)
						Value := TransForm(Value)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegWrite create '" ValueType "' registry item '" ValueName "' in '" KeyName "' with value '" Value "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "RegWrite create '" ValueType "' registry item '" ValueName "' in '" KeyName "' with value '" Value "'"
						if (ValueType = "REG_SZ" OR ValueType = "REG_EXPAND_SZ" OR ValueType = "REG_MULTI_SZ" OR ValueType = "REG_DWORD" OR ValueType = "REG_BINARY")
							{
							try
								RegWrite Value, ValueType, KeyName, ValueName
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegWrite failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegWrite failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "RegWrite failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegWrite successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "RegWrite successfully executed."
								}
							}
						if ValueType = "REG_QWORD"
							{
							RootKey := ""
							SubKey := ""
							loop parse, KeyName, "\"
								{
								if (A_LoopField = "HKCR" or A_LoopField = "HKEY_CLASSES_ROOT" or A_LoopField = "HKCU" or A_LoopField = "HKEY_CURRENT_USER" or A_LoopField = "HKLM" or A_LoopField = "HKEY_LOCAL_MACHINE" or A_LoopField = "HKU" or A_LoopField = "HKEY_USERS" or A_LoopField = "HKCC" or A_LoopField = "HKEY_CURRENT_CONFIG")
									{
									RootKey := A_LoopField
									SubKey := StrReplace(KeyName, RootKey . "\", "")
									break
									}
								}
							try
								Reg_QWORD.Write(RootKey, SubKey, ValueName, Value)
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegWrite (REG_QWORD) failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegWrite (REG_QWORD) failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "RegWrite failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegWrite (REG_QWORD) successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "RegWrite successfully executed."
								}
							}
						}
					if ScriptItemName = "RegDelete"
						{
						KeyName := ""
						ValueName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								KeyName := ScriptItemString[2]
							if A_Index = 3
								ValueName := ScriptItemString[3]
							if A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						KeyName := TransForm(KeyName)
						ValueName := TransForm(ValueName)
						UserDescription := TransForm(UserDescription)
						if ValueName
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegDelete registry value '" ValueName "' from '" KeyName "'`r`n" , LogFile
							else
								if UserScript
									if UserDescription
										ShowInfo.Value := UserDescription
									else
										ShowInfo.Value := "RegDelete registry value '" ValueName "' from '" KeyName "'"
							try
								RegDelete KeyName, ValueName
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "RegDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegDelete successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "RegDelete successfully executed."
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegDelete registry key: " KeyName "`r`n" , LogFile
							else
								if UserScript
									if UserDescription
										ShowInfo.Value := UserDescription
									else
										ShowInfo.Value := "RegDelete registry key: " KeyName
							try
								RegDeleteKey KeyName
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRegDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "RegDelete failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRegDelete successfully executed.`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "RegDelete successfully executed."
								}
							}
						}
					if ScriptItemName = "FileVersion"
						{
						Filename := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Filename := ScriptItemString[2]
							if A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						Filename := TransForm(Filename)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRetrieve the version for file name: " Filename "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "Retrieve the version for file name: " Filename
						if FileExist(Filename)
							{
							FileVersion := FileGetVersion(Filename)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileVersion successfully executed with value: " FileVersion "`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "FileVersion successfully executed with value: " FileVersion
							try
								EnvSet "FileVersion", FileVersion
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for 'FileVersion' with value '" FileVersion "' Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for 'FileVersion' with value '" FileVersion "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "Environment variable failed to set for 'FileVersion' with value '" FileVersion "' Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for 'FileVersion' with value: " FileVersion "`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Environment variable successfully set for 'FileVersion' with value: " FileVersion
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFile not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFile not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "File not found."
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "IniRead"
						{
						Filename := ""
						Section := ""
						KeyName := ""
						DefaultValue := ""
						IniValue := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Filename := ScriptItemString[2]
							if A_Index = 3
								Section := ScriptItemString[3]
							if A_Index = 4
								KeyName := ScriptItemString[4]
							if A_Index = 5
								DefaultValue := ScriptItemString[5]
							if A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						Filename := TransForm(Filename)
						Section := Trim(Section)
						KeyName := Trim(KeyName)
						DefaultValue := TransForm(DefaultValue)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tIniRead from file name '" Filename "' using section name '" Section "' for key name '" KeyName "' with default value '" DefaultValue "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "IniRead from file name '" Filename "' using section name '" Section "' for key name '" KeyName "' with default value '" DefaultValue "'"
						try
							IniValue := IniRead(Filename, Section, KeyName, DefaultValue)
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tIniRead failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tIniRead failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "IniRead failed with Error: " err.Message
									Sleep 10000
									}
							IniValue := DefaultValue
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tIniRead successfully executed with value: " IniValue "`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "IniRead successfully executed with value: " IniValue
							}
						try
							EnvSet KeyName, IniValue
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "Environment variable failed to set for '" KeyName "' with value '" IniValue "' Error: " err.Message
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for '" KeyName "' with value: " IniValue "`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "Environment variable successfully set for '" KeyName "' with value: " IniValue
							}
						}
					if ScriptItemName = "IniWrite"
						{
						Value := ""
						Filename := ""
						Section := ""
						KeyName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Value := ScriptItemString[2]
							if A_Index = 3
								Filename := ScriptItemString[3]
							if A_Index = 4
								Section := ScriptItemString[4]
							if A_Index = 5
								KeyName := ScriptItemString[5]
							if A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						Value := TransForm(Value)
						Filename := TransForm(Filename)
						Section := Trim(Section)
						KeyName := Trim(KeyName)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tIniWrite value '" Value "' in file name '" Filename "' for section name '" Section "' with key name '" KeyName "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "IniWrite value '" Value "' in file name '" Filename "' for section name '" Section "' with key name '" KeyName "'"
						try
							IniWrite Value, Filename, Section, KeyName
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tIniWrite failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tIniWrite failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "IniWrite failed with Error: " err.Message
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tIniWrite successfully executed.`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "IniWrite successfully executed."
							}
						}
					if ScriptItemName = "IniDelete"
						{
						Filename := ""
						Section := ""
						KeyName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Filename := ScriptItemString[2]
							if A_Index = 3
								Section := ScriptItemString[3]
							if A_Index = 4
								KeyName := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Filename := TransForm(Filename)
						Section := Trim(Section)
						KeyName := Trim(KeyName)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tIniDelete using file name '" Filename "' for section name '" Section "' (with key name '" KeyName "')`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "IniDelete using file name '" Filename "' for section name '" Section "' (with key name '" KeyName "')"
						try
							IniDelete Filename, Section, KeyName
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tIniDelete failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tIniDelete failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "IniDelete failed with Error: " err.Message
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tIniDelete successfully executed.`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "IniDelete successfully executed."
							}
						}
					if ScriptItemName = "FileCreateShortcut"
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
							if A_Index = 2
								ShortcutTarget := ScriptItemString[2]
							if A_Index = 3
								ShortcutLinkFile := ScriptItemString[3]
							if A_Index = 4
								ShortcutWorkingDir := ScriptItemString[4]
							if A_Index = 5
								ShortcutArgs := ScriptItemString[5]
							if A_Index = 6
								ShortcutDescription := ScriptItemString[6]
							if A_Index = 7
								ShortcutIconFile := ScriptItemString[7]
							if A_Index = 8
								ShortcutKey := ScriptItemString[8]
							if A_Index = 9
								ShortcutIconNumber := ScriptItemString[9]
							if A_Index = 10
								ShortcutRunState := ScriptItemString[10]
							if A_Index = 11
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
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileCreateShortcut using the following properties: Target = '" ShortcutTarget "' LinkFile = '" ShortcutLinkFile "' WorkingDir = '" ShortcutWorkingDir "' Args = '" ShortcutArgs "' Description = '" ShortcutDescription "' IconFile = '" ShortcutIconFile "' ShortcutKey = '" ShortcutKey "' IconNumber = '" ShortcutIconNumber "' RunState = '" ShortcutRunState "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FileCreateShortcut using the following properties: Target = '" ShortcutTarget "' LinkFile = '" ShortcutLinkFile "' WorkingDir = '" ShortcutWorkingDir "' Args = '" ShortcutArgs "' Description = '" ShortcutDescription "' IconFile = '" ShortcutIconFile "' ShortcutKey = '" ShortcutKey "' IconNumber = '" ShortcutIconNumber "' RunState = '" ShortcutRunState "'"
						try
							FileCreateShortcut ShortcutTarget, ShortcutLinkFile, ShortcutWorkingDir, ShortcutArgs, ShortcutDescription, ShortcutIconFile, ShortcutKey, ShortcutIconNumber, ShortcutRunState
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileCreateShortcut failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileCreateShortcut failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "FileCreateShortcut failed with Error: " err.Message
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileCreateShortcut successfully executed.`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "FileCreateShortcut successfully executed."
							}
						}
					if ScriptItemName = "Run"
						{
						Target := ""
						WorkingDir := ""
						Options := ""
						NoWait := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Target := ScriptItemString[2]
							if A_Index = 3
								WorkingDir := ScriptItemString[3]
							if A_Index = 4
								Options := ScriptItemString[4]
							if A_Index = 5
								NoWait := ScriptItemString[5]
							if A_Index = 6
								UserDescription := ScriptItemString[6]
							}
						Target := TransForm(Target)
						if WorkingDir
							WorkingDir := TransForm(WorkingDir)
						else
							{
							if !UserScript
								WorkingDir := A_ScriptDir
							}
						Options := TransForm(Options)
						NoWait := TransForm(NoWait)
						UserDescription := TransForm(UserDescription)
						if NoWait = 0
							{
							if FileExist(LogFile)
								{
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRun and wait for Target: " Target "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
								}
							else
								if UserScript
									if UserDescription
										ShowInfo.Value := UserDescription
									else
										ShowInfo.Value := "Run and wait for Target: " Target "`r`nWorking directory: " WorkingDir
							if UserScript
								SetTimer RunWaitProgress, 50
							try
								ReturnCode := RunWait(Target, WorkingDir, Options, &PID)
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										SetTimer RunWaitProgress, 0
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "Run failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully stopped with return code: " ReturnCode "`r`n" , LogFile
								else
									if UserScript
										{
										SetTimer RunWaitProgress, 0
										if !UserDescription
											ShowInfo.Value := "Run target with Process Id '" PID "' successfully stopped with return code: " ReturnCode
										}
								try
									EnvSet "ReturnCode", ReturnCode
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , LogFile
									else
										if UserScript
											{
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
											ShowInfo.Value := "Environment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message
											Sleep 10000
											}
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for 'ReturnCode' with value: " ReturnCode "`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Environment variable successfully set for 'ReturnCode' with value: " ReturnCode
									}
								}
							}
						else
							{
							if FileExist(LogFile)
								{
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRun Target: " Target "`r`n" , LogFile
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
								}
							else
								if UserScript
									if UserDescription
										ShowInfo.Value := UserDescription
									else
										ShowInfo.Value := "Run Target: " Target "`r`nWorking directory: " WorkingDir
							try
								ReturnCode := Run(Target, WorkingDir, Options, &PID)
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tRun failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "Run failed with Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRun target with Process Id '" PID "' successfully started with return code: " ReturnCode "`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Run target with Process Id '" PID "' successfully started with return code: " ReturnCode
								try
									EnvSet "ReturnCode", ReturnCode
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , LogFile
									else
										if UserScript
											{
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
											ShowInfo.Value := "Environment variable failed to set for 'ReturnCode' with value '" ReturnCode "' Error: " err.Message
											Sleep 10000
											}
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for 'ReturnCode' with value: " ReturnCode "`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "Environment variable successfully set for 'ReturnCode' with value: " ReturnCode
									}
								}
							}
						}
					if ScriptItemName = "StringReplace"
						{
						Source := ""
						FindText := ""
						ReplaceText := ""
						Destination := ""
						OverWrite := 0
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Source := ScriptItemString[2]
							if A_Index = 3
								FindText := ScriptItemString[3]
							if A_Index = 4
								ReplaceText := ScriptItemString[4]
							if A_Index = 5
								Destination := ScriptItemString[5]
							if A_Index = 6
								OverWrite := ScriptItemString[6]
							if A_Index = 7
								UserDescription := ScriptItemString[7]
							}
						Source := TransForm(Source)
						FindText := TransForm(FindText)
						ReplaceText := TransForm(ReplaceText)
						Destination := TransForm(Destination)
						OverWrite := TransForm(OverWrite)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStringReplace text '" FindText "' with '" ReplaceText "' using source file '" Source "' for destination file: '" Destination "' (OverWrite=" OverWrite ")`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "StringReplace text '" FindText "' with '" ReplaceText "' using source file '" Source "' for destination file: '" Destination "' (OverWrite=" OverWrite ")"
						if FileExist(Source)
							{
							if Source = Destination
								{
								Contents := FileRead(Source)
								Contents := StrReplace(Contents, FindText, ReplaceText)
								FileDelete Source
								try
									FileAppend Contents, Source
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , LogFile
									else
										if UserScript
											{
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
											ShowInfo.Value := "StringReplace failed with Error: " err.Message
											Sleep 10000
											}
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile
									else
										if UserScript
											if !UserDescription
												ShowInfo.Value := "StringReplace successfully executed."
									}
								Contents := ""
								}
							else
								{
								if OverWrite = 1
									{
									if FileExist(Destination)
										FileDelete Destination
									Contents := FileRead(Source)
									Contents := StrReplace(Contents, FindText, ReplaceText)
									try
										FileAppend Contents, Destination
									catch as err
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , LogFile
										else
											if UserScript
												{
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
												ShowInfo.Value := "StringReplace failed with Error: " err.Message
												Sleep 10000
												}
										}
									else
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile
										else
											if UserScript
												if !UserDescription
													ShowInfo.Value := "StringReplace successfully executed."
										}
									Contents := ""
									}
								else
									{
									if Not FileExist(Destination)
										{
										Contents := FileRead(Source)
										Contents := StrReplace(Contents, FindText, ReplaceText)
										try
											FileAppend Contents, Destination	
										catch as err
											{
											if FileExist(LogFile)
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , LogFile
											else
												if UserScript
													{
													FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStringReplace failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
													ShowInfo.Value := "StringReplace failed with Error: " err.Message
													Sleep 10000
													}
											}
										else
											{
											if FileExist(LogFile)
												FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStringReplace successfully executed.`r`n" , LogFile
											else
												if UserScript
													if !UserDescription
														ShowInfo.Value := "StringReplace successfully executed."
											}
										Contents := ""
										}
									else
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t'" Destination "' already exists.`r`n" , LogFile
										else
											if UserScript
												ShowInfo.Value := "'" Destination "' already exists."
										}
									}
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" Source "' not found.`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "'" Source "' not found."
									Sleep 10000
									}
							}
						}	
					if ScriptItemName = "FileAppend"
						{
						Text := ""
						FileName := ""
						Encoding := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								Text := ScriptItemString[2]
							if A_Index = 3
								FileName := ScriptItemString[3]
							if A_Index = 4
								Encoding := ScriptItemString[4]
							if A_Index = 5
								UserDescription := ScriptItemString[5]
							}
						Text := TransForm(Text)
						FileName := TransForm(FileName)
						Encoding := TransForm(Encoding)
						UserDescription := TransForm(UserDescription)
						if Encoding = ""
							Encoding := A_FileEncoding
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileAppend write text '" Text "' to file: " FileName " (" Encoding ")`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "FileAppend write text '" Text "' to file: " FileName " (" Encoding ")"
						Text := Text "`r`n"
						try	
							FileAppend Text, FileName, "`r`n " Encoding
						catch as err
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tFileAppend failed with Error: " err.Message "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tFileAppend failed with Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "FileAppend failed with Error: " err.Message
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileAppend successfully executed.`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "FileAppend successfully executed."
							}
						}
					if ScriptItemName = "Download"
						{
						URL := ""
						FileName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								URL := ScriptItemString[2]
							if A_Index = 3
								FileName := ScriptItemString[3]
							if A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						URL := Trim(URL)
						FileName := TransForm(FileName)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDownload '" URL "' for file: " FileName "`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "Download '" URL "' for file: " FileName
						Download URL, FileName
						if Not FileExist(FileName)
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tDownload failed with Error: " A_LastError "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tDownload failed with Error: " A_LastError "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "Download failed with Error: " A_LastError
									Sleep 10000
									}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDownload successfully executed.`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "Download successfully executed."
							}
						}
					if ScriptItemName = "Powershell"
						{
						psCommand := ""
						WorkingDir := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								psCommand := ScriptItemString[2]
							if A_Index = 3
								WorkingDir := ScriptItemString[3]
							if A_Index = 4
								UserDescription := ScriptItemString[4]
							}
						psCommand := TransForm(psCommand)
						if WorkingDir
							WorkingDir := TransForm(WorkingDir)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							{
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tExecute Powershell command: " psCommand "`r`n" , LogFile
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tWorking directory: " WorkingDir "`r`n" , LogFile
							}
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "Execute Powershell command: " psCommand
						SplitPath psCommand,,, &Extension
						if Extension = "ps1"
							{
							try
								psCommand := FileRead(psCommand)
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tReading Powershell script failed with error:`r`n" err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := err.Message
										Sleep 10000
										}
								}
							}
						else
							psCommand := StrReplace(psCommand, A_Space . "\n" . A_Space, "`r`n")
						if WorkingDir
							psCommand := "Set-Location -Path `"" WorkingDir "`"`r`n" psCommand
						psCommand := Format(psCommand)
						if UserScript
							{
							try 
								ps := ComObject("psScript")
							catch as err
								{
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tpsScript error:`r`n" err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
								ShowInfo.Value := err.Message
								Sleep 10000
								}
							else
								{
								try
									psReturn := ps.PS_Script(psCommand)
								catch as err
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tPowershell command stopped with error:`r`n" err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := err.Message
									Sleep 10000
									}
								else
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tPowershell command successfully stopped with return message:`r`n" psReturn "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									if !UserDescription
										ShowInfo.Value := "Powershell command successfully executed."
									}
								}
							}
						else
							{
							try 
								psReturn := PsScriptManager.ComInstance.PS_Script(psCommand)
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tPowershell command stopped with error:`r`n" err.Message "`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tPowershell command successfully stopped with return message:`r`n" psReturn "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "ServiceState"
						{
						Status := ""
						ServiceName := ""
						UserDescription := ""
						Loop ScriptItemString.Length
							{
							if A_Index = 2
								ServiceName := ScriptItemString[2]
							if A_Index = 3
								UserDescription := ScriptItemString[3]
							}
						ServiceName := Transform(ServiceName)
						UserDescription := TransForm(UserDescription)
						if FileExist(LogFile)
							FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRetrieve running status for the Windows Service name: " ServiceName "'`r`n" , LogFile
						else
							if UserScript
								if UserDescription
									ShowInfo.Value := UserDescription
								else
									ShowInfo.Value := "Retrieve running status for the Windows Service name: " ServiceName
						Status := WinService.State(ServiceName, true)
						if (Status = "Stopped" OR Status = "Start Pending" OR Status = "Stop Pending" OR Status = "Running" OR Status = "Continue Pending" OR Status = "Pause Pending" OR Status = "Paused" OR Status = "Unknown")
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tService status successfully retrieved with value: " Status "`r`n" , LogFile
							else
								if UserScript
									if !UserDescription
										ShowInfo.Value := "Service status successfully retrieved with value: " Status
							try
								EnvSet ServiceName, Status
							catch as err
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" ServiceName "' with value '" Status "' Error: " err.Message "`r`n" , LogFile
								else
									if UserScript
										{
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tEnvironment variable failed to set for '" ServiceName "' with value '" Status "' Error: " err.Message "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
										ShowInfo.Value := "Environment variable failed to set for '" ServiceName "' with value '" Status "' Error: " err.Message
										Sleep 10000
										}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tEnvironment variable successfully set for '" ServiceName "' with value: " Status "`r`n" , LogFile
								else
									if UserScript
										if !UserDescription
											ShowInfo.Value := "Environment variable successfully set for '" ServiceName "' with value: " Status
								}
							}
						else
							{
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRetrieving service status failed with error: " Status "`r`n" , LogFile
							else
								if UserScript
									{
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tRetrieving service status for service '" ServiceName "' failed with error: " Status "`r`n" , A_AppData . "\Intune Win32 Launcher error.log", "UTF-16"
									ShowInfo.Value := "Retrieving service status for service '" ServiceName "' failed with error: " Status
									Sleep 10000
									}
							}
						}
					if ScriptItemName = "ServiceCreate"
						{
						if !UserScript
							{
							ReturnMsg := ""
							Status := ""
							ServiceName := ""
							BinaryPath := ""
							StartType := ""
							DisplayName := ""
							ServiceStartName := ""
							Password := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									ServiceName := ScriptItemString[2]
								if A_Index = 3
									BinaryPath := ScriptItemString[3]
								if A_Index = 4
									StartType := ScriptItemString[4]
								if A_Index = 5
									DisplayName := ScriptItemString[5]
								if A_Index = 6
									ServiceStartName := ScriptItemString[6]
								if A_Index = 7
									Password := ScriptItemString[7]
								}
							ServiceName := Transform(ServiceName)
							BinaryPath := Transform(BinaryPath)
							StartType := Trim(StartType)
							DisplayName := Transform(DisplayName)
							ServiceStartName := Trim(ServiceStartName)
							Password := Trim(Password)
							if ServiceStartName
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreate Windows Service ('" ServiceStartName "' account) with service name '" ServiceName "' using executable path '" BinaryPath "' with start type '" StartType "' and display name '" DisplayName "'.`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreate Windows Service (LocalSystem account) with service name '" ServiceName "' using executable path '" BinaryPath "' with start type '" StartType "' and display name '" DisplayName "'.`r`n" , LogFile
								}
							ReturnMsg := WinService.Add(ServiceName, BinaryPath, StartType, DisplayName, ServiceStartName, Password)
							if ReturnMsg = 1
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreate Windows Service successfully executed.`r`n" , LogFile
								Loop 10
									{
									Status := WinService.State(ServiceName, false)
									if (Status = 1 OR  Status = 4)
										break
									else
										Sleep 1000
									}
								Status := WinService.State(ServiceName, true)
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tService status successfully retrieved with value: " Status "`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tCreate Windows Service failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "ServiceDelete"
						{
						if !UserScript
							{
							ReturnMsg := ""
							Status := ""
							ServiceName := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									ServiceName := ScriptItemString[2]
								}
							ServiceName := TransForm(ServiceName)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDelete Windows Service (system) with service name '" ServiceName "'.`r`n" , LogFile
							ReturnMsg := WinService.Delete(ServiceName)
							if ReturnMsg = 1
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDelete Windows Service successfully executed.`r`n" , LogFile
								Loop 10
									{
									Status := WinService.State(ServiceName, false)
									if (Status = 1060)
										break
									else
										Sleep 1000
									}
								Status := WinService.State(ServiceName, true)
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tService status successfully retrieved with value: " Status "`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tDelete Windows Service failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "ServiceStart"
						{
						if !UserScript
							{
							ReturnMsg := ""
							Status := ""
							ServiceName := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									ServiceName := ScriptItemString[2]
								}
							ServiceName := TransForm(ServiceName)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart Windows Service with service name '" ServiceName "'.`r`n" , LogFile
							ReturnMsg := WinService.Start(ServiceName)
							if (ReturnMsg = 1 OR ReturnMsg = 1056)
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStart Windows Service successfully executed.`r`n" , LogFile
								Loop 10
									{
									Status := WinService.State(ServiceName, false)
									if Status = 4
										break
									else
										Sleep 1000
									}
								Status := WinService.State(ServiceName, true)
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tService status successfully retrieved with value: " Status "`r`n" , LogFile
								}
							else
								{
								if ReturnMsg = 1060
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tThe specified service does not exist as an installed service.`r`n" , LogFile
								else
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStart Windows Service failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "ServiceStop"
						{
						if !UserScript
							{
							ReturnMsg := ""
							Status := ""
							ServiceName := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									ServiceName := ScriptItemString[2]
								}
							ServiceName := TransForm(ServiceName)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStop Windows Service with service name '" ServiceName "'.`r`n" , LogFile
							ReturnMsg := WinService.Stop(ServiceName)
							if (ReturnMsg = 1 OR ReturnMsg = 1062)
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStop Windows Service successfully executed.`r`n" , LogFile
								Loop 10
									{
									Status := WinService.State(ServiceName, false)
									if Status = 1
										break
									else
										Sleep 1000
									}
								Status := WinService.State(ServiceName, true)
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tService status successfully retrieved with value: " Status "`r`n" , LogFile
								}
							else
								{
								if ReturnMsg = 1060
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tThe specified service does not exist as an installed service.`r`n" , LogFile
								else
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tStop Windows Service failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "CreateTask"
						{
						if !UserScript
							{
							ReturnMsg := ""
							TaskName := ""
							Description := ""
							Path := ""
							Arguments := ""
							ExecutionTimeLimit := 12
							TriggerType := "Logon"
							Battery := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									TaskName := ScriptItemString[2]
								if A_Index = 3
									Description := ScriptItemString[3]
								if A_Index = 4
									Path := ScriptItemString[4]
								if A_Index = 5
									Arguments := ScriptItemString[5]
								if A_Index = 6
									ExecutionTimeLimit := ScriptItemString[6]
								if A_Index = 7
									TriggerType := ScriptItemString[7]
								if A_Index = 8
									Battery := ScriptItemString[8]
								}
							TaskName := TransForm(TaskName)
							Description := TransForm(Description)
							Path := TransForm(Path)
							Arguments := TransForm(Arguments)
							ExecutionTimeLimit := Trim(ExecutionTimeLimit)
							TriggerType := Trim(TriggerType)
							Battery := Trim(Battery)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreateTask (system) with task name '" TaskName "' and description '" Description "' using executable path '" Path "' with the following arguments '" Arguments "'. Set limited runtime execution for " ExecutionTimeLimit " hours. Set trigger type for next " TriggerType ". Run on battery is set with value '" Battery "'. When empty the task will only run when a power supply is connected.`r`n" , LogFile
							ReturnMsg := CreateTask(TaskName, Description, Path, Arguments, ExecutionTimeLimit, TriggerType, Battery)
							if ReturnMsg = 0
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tCreateTask successfully executed.`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tCreateTask failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "DeleteTask"
						{
						if !UserScript
							{
							ReturnMsg := ""
							TaskName := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									TaskName := ScriptItemString[2]
								}
							TaskName := TransForm(TaskName)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDeleteTask (system) with task name '" TaskName "'.`r`n" , LogFile
							ReturnMsg := DeleteTask(TaskName)
							if ReturnMsg = 0
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tDeleteTask successfully executed.`r`n" , LogFile
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tDeleteTask failed with Error: " ReturnMsg "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "Notify"
						{
						if !UserScript
							{
							MessageSection := ""
							ExitCode := 0
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									MessageSection := ScriptItemString[2]
								}
							MessageSection := TransForm(MessageSection)
							SectionType := IniRead(Inifile, MessageSection, "SectionType", "")
							SectionType := Trim(SectionType)
							if (SectionType = "Message")
								{
								if (cpauPID := ProcessExist(cpauPID))
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t[ Performing Message notification section: " MessageSection " ]`r`n" , LogFile
									Send_NamedPipeMessage(MessageSection . ":" . ExitCode, SystemPipeName)
									ReturnCode := Receive_NamedPipeMessage(UserPipeName)
									if ReturnCode = 0
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
										}
									else
										{
										if FileExist(LogFile)
											FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`tNotification process returned: " ReturnCode "`r`n" , LogFile
										}
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tNotification process not running. Skipping message notification section: " MessageSection "`r`n" , LogFile
									}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tMessage notification section not found: " MessageSection "`r`n" , LogFile
								}
							}
						}
					if ScriptItemName = "ExitApp"
						{
						if !UserScript
							{
							ExitCode := 0
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									ExitCode := ScriptItemString[2]
								}
							ExitCode := Trim(ExitCode)
							if (cpauPID := ProcessExist(cpauPID))
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStopping notification process...`r`n" , LogFile
								Send_NamedPipeMessage("*:0", SystemPipeName)
								ReturnCode := Receive_NamedPipeMessage(UserPipeName)
								}
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" ProductName " stopped successfully. (Return Code = " ExitCode ")`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
							ExitApp ExitCode
							}
						}
					if ScriptItemName = "Reboot"
						{
						if !UserScript
							{
							ExitCode := 350
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									ExitCode := ScriptItemString[2]
								}
							ExitCode := Trim(ExitCode)
							if (cpauPID := ProcessExist(cpauPID))
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tStopping notification process...`r`n" , LogFile
								Send_NamedPipeMessage("*:0", SystemPipeName)
								ReturnCode := Receive_NamedPipeMessage(UserPipeName)
								}
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tRun Target: " A_Comspec " /c Shutdown /r`r`n" , LogFile
							ReturnCode := Run(A_Comspec " /c Shutdown /r", "", "", &PID)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`t" ProductName " stopped successfully. (Return Code = " ExitCode ")`r`n---------- --------`t---------`t----------------------------------------------------------------------------------------------------------------------------`r`n" , LogFile
							ExitApp ExitCode
							}
						}
					if ScriptItemName = "FileAccess"
						{
						if !UserScript
							{
							FileName := ""
							Loop ScriptItemString.Length
								{
								if A_Index = 2
									FileName := ScriptItemString[2]
								}
							FileName := TransForm(FileName)
							if FileExist(LogFile)
								FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileAccess (everyone read and execute permission) for filename: " FileName "`r`n" , LogFile
							if FileExist(FileName)
								{
								try
									ReturnValue := SetSecurityFile(FileName, "S-1-1-0", "1179817", "1")
								catch as err
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ WARN  ]`tFileAccess failed with Error: " err.Message "`r`n" , LogFile
									}
								else
									{
									if FileExist(LogFile)
										FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ INFO  ]`tFileAccess successfully executed. return value: " ReturnValue "`r`n" , LogFile
									}
								}
							else
								{
								if FileExist(LogFile)
									FileAppend FormatTime(A_Now, "dd-MM-yyyy HH:mm:ss") "`t[ ERROR ]`t'" FileName "' not found.`r`n" , LogFile
								}
							}
						}
					Counter++
					}
				else
					break
				}
			}
		if A_LoopReadLine = "[" Section "]"
			Found := 1
		}
	Return Counter
	}