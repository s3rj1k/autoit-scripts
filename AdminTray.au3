#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.2.0
 Author:		 Serj

 Script Function:
	AdminTray

#ce ----------------------------------------------------------------------------

#NoTrayIcon
; #RequireAdmin
#include <Array.au3>
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <SliderConstants.au3>
#include <Security.au3>
#include <SliderConstants.au3>
#include <StaticConstants.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>
Opt("GUIOnEventMode", 1)
Opt("MustDeclareVars", 1)
Opt("TrayMenuMode", 1 + 2 + 4)
Opt("TrayOnEventMode", 1)
Global $ExplorerPID, $SplashTimerInit, $SplashTimerDiff, $DLL;, $TrayState = True
Global CONST $DropBoxSizeX = 48, $DropBoxSizeY = 48
#OnAutoItStartRegister "OnAutoItStart"
OnAutoItExitRegister("_Exit")

If IsAdmin() Then
	Init()
Else
	MsgBox(0, "Ошибка", "Требуются права Администратора!") ; Admin rights are required!
	_Exit()
EndIf

Func OnAutoItStart()
	TraySetClick(0)
	TraySetState(2)
EndFunc ; OnAutoItStart

Func Init()
	Global $sProfilesDirectory = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory")
	$sProfilesDirectory = ExpandEnvironmentVariable($sProfilesDirectory)
	Local CONST $AdminGroupSID = "S-1-5-32-544"
		; Administrators [S-1-5-32-544], Users [S-1-5-32-545], Guests [S-1-5-32-546], Power Users [S-1-5-32-547]
	Local $aLocalAdminGroupName = _Security__LookupAccountSid($AdminGroupSID)
	If @error Or $aLocalAdminGroupName[2] <> 4 Then
		MsgBox(0, "Ошибка, Local Admin Group Name", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aCompSID = _Security__LookupAccountName(@ComputerName)
	If @error Or $aCompSID[2] <> 3 Then
		MsgBox(0, "Ошибка, Computer SID", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aLocalAdminName = _Security__LookupAccountSid($aCompSID[0] & "-500")
	If @error Or $aLocalAdminName[2] <> 1 Then
		MsgBox(0, "Ошибка, Local Admin Name", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aUserEnum = _NetUserEnum()
	If @error Then
		MsgBox(0, "Ошибка, User Enum", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Global CONST $LocalAdminName = $aLocalAdminName[0]
	Global CONST $LocalAdminGroupName = $aLocalAdminGroupName[0]
	If $CmdLine[0] == 1 And $CmdLine[1] == "-Menu" Then
		Main()
	ElseIf $CmdLine[0] == 3 And $CmdLine[1] == "-AccountCreds" Then
		Local $TestCredsRet = TestCreds($CmdLine[2], $CmdLine[3])
		If @error Then
			MsgBox(0, "Ошибка", $TestCredsRet)
			_Exit()
		Else
			AdminMenuInit($CmdLine[2], $CmdLine[3])
		EndIf
	Else
		GetAccountCreds($aUserEnum)
	EndIf
EndFunc ; Init

Func GetAccountCreds($aUserEnum)
	Local $sUserEnum = _ArrayToString($aUserEnum, "|", 1, $aUserEnum[0] + 1)
	Local $sUserEnumDefault = $aUserEnum[1]
	For $i = 1 to $aUserEnum[0] Step 1
		Local $IsAdmin = _IsAdmin($aUserEnum[$i])
		If $IsAdmin == False Then
			Local $sUserEnumDefault = $aUserEnum[$i]
			ExitLoop
		EndIf
	Next
	Local $HWND = GUICreate("Настройки AdminTray:", 272, 96, -1, -1, BitOr($WS_CAPTION, $WS_SYSMENU), BitOr($WS_EX_APPWINDOW, $WS_EX_TOPMOST)) ; Setup
		GUICtrlCreateGroup("", 6, 0, 260, 64)
			GUICtrlCreateLabel("Логин:", 12, 15, 60, 20) ; Login
			Local $Login_Combo = GUICtrlCreateCombo("", 80, 12, 180, 60, BitOr($CBS_DROPDOWN, $CBS_AUTOHSCROLL, $WS_VSCROLL, $CBS_SORT, $CBS_DISABLENOSCROLL))
				GUICtrlSetData($Login_Combo, $sUserEnum, $sUserEnumDefault)
				GUICtrlSetTip($Login_Combo, "Логин пользователя") ; Account login
				GUICtrlCreateLabel("Пароль:", 12, 40, 60, 20) ; Password
			Local $Password_Input = GUICtrlCreateInput("", 80, 38, 180, 20, BitOr($ES_PASSWORD, $ES_AUTOHSCROLL))
				GUICtrlSetTip($Password_Input, "Пароль пользователя") ; Account password
				GUICtrlSetState($Password_Input, $GUI_FOCUS)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		Local $Save_Button = GUICtrlCreateButton("Сохранить", 5, 68, 156, 23) ; Save
			GUICtrlSetOnEvent($Save_Button, "Save_Button")
			GUICtrlSetState($Save_Button, $GUI_DEFBUTTON)
		Local $Exit_Button = GUICtrlCreateButton("Выйти", 167, 68, 100, 23) ; Exit
			GUICtrlSetOnEvent($Exit_Button, "_Exit")
		GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit", $HWND)
	GUISetState(@SW_SHOW)
	While 1
		Sleep(1000)
	WEnd
EndFunc ; GetAccountCreds

Func Save_Button()
	Local $UserName = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:1]")
	Local $Password = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:2]")
	Local $TestCredsRet = TestCreds($UserName, $Password)
	If @error Then
		ControlSetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:2]", "")
		MsgBox(0, "Ошибка", $TestCredsRet, 0, @GUI_WINHANDLE)
	Else
		GUIDelete(@GUI_WINHANDLE)
		AdminMenuInit($UserName, $Password)
	EndIf
EndFunc	;Save_Button

Func TestCreds($UserName, $Password)
	Local $sRet = ""
	Local $ErrExtendedMacro
	Local $IsAdmin = _IsAdmin($UserName)
	If @error Then
		$sRet &= "_IsAdmin: " & @CRLF & @TAB & $IsAdmin & @CRLF
		$ErrExtendedMacro = 0x1
	EndIf
	If $IsAdmin == True Then $sRet &= "_IsAdmin: " & @CRLF & @TAB & "Указанный пользователь находится в локальной группе администраторов!" & @CRLF
		; Specified account is member of local administrator group!
	Local $LogonUser = _LogonUser($UserName, $Password)
	If @error Then
		$sRet &= "_LogonUser: " & @CRLF & @TAB & $LogonUser & @CRLF
		$ErrExtendedMacro = 0x1
	EndIf
	If $LogonUser == False Then $sRet &= "_LogonUser: " & @CRLF & @TAB & "Не верный логин/пароль." _
		& @CRLF & @TAB & "Пустые пароли не поддерживаются!" & @CRLF ; Incorrect Login/Password. Blank passwords are forbidden!
	If @OSVersion == "WIN_2000" Then $LogonUser = True
	If $LogonUser == True And $IsAdmin == False Then
		Return SetError(0)
	Else
		Return SetError(1, $ErrExtendedMacro, $sRet)
	EndIf
EndFunc	;TestCreds

Func AdminMenuInit($UserName, $Password)
	If @compiled == 1 Then
		Local $ExecPath = @AutoItExe
	Else
		Local $ExecPath = '"' & @AutoItExe & '" "' & @ScriptFullPath & '"'
	EndIf
	Local $WorkingDir = @ScriptDir
	Local $sRet = _MakeMeAdmin_RunAs($UserName, $Password, $ExecPath & ' "-Menu"', $WorkingDir)
	If @error Then
		MsgBox(0, "Ошибка, _MakeMeAdmin_RunAs", $sRet & "@extended = " & @extended)
	Else
		If @OSVersion <> "WIN_2000" Then
			Local $NoDefaultAdminOwnerFlag = RegRead("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa", "NoDefaultAdminOwner")
			If @error Or ($NoDefaultAdminOwnerFlag == 0 And @EXTENDED == $REG_DWORD) Then
				RegWrite("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Lsa", "NoDefaultAdminOwner", "REG_DWORD", "00000001")
				Msgbox(0, "Внимание", "Требуется перезагрузка системы для правильной работы программы!")
					; Restart is required for proper application functionality
			EndIf
		EndIf
	EndIf
	_Exit()
EndFunc	;AdminMenuInit

Func Main()
	If _Singleton("AdminTray", 1) = 0 Then
		Msgbox(0, "Ошибка", "Программа уже запущена!") ; Script is already running
		_Exit()
	EndIf
	If @OSVersion == "WIN_2000" Then
		Local $SeparateProcessFlag = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SeparateProcess")
		If @error Or ($SeparateProcessFlag == 0 And @EXTENDED == $REG_DWORD) Then
			RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "SeparateProcess", "REG_DWORD", "00000001")
			Msgbox(0, "Внимание", "Требуется перезагрузка системы для правильной работы программы!")
				; Restart is required for proper application functionality
		EndIf
	EndIf
	Tray()
	GUI()
	$DLL = DllOpen("user32.dll")
	While 1
		HotKeySet("#{1}", "MyComputer")
		HotKeySet("#{2}", "cpl")
		HotKeySet("#{3}", "NetworkConnections")
		HotKeySet("#{0}", "RestartShell")
		HotKeySet("#{-}", "cmd")
		HotKeySet("#{=}", "open")
		HotKeySet("#{SPACE}", "taskmgr")
		HotKeySet("^!{BREAK}", "_Exit")
		$SplashTimerDiff = TimerDiff($SplashTimerInit)
		If $SplashTimerDiff > 4000 Then SplashOff()
		Sleep(500)
	WEnd
	_Exit()
EndFunc ; Main

Func Tray()
	TraySetState(1)
	TraySetClick(8)
	If @Compiled == 1 Then
		TraySetIcon(@AutoItExe, 99)
	Else
		TraySetIcon(@ScriptDir & "\Res\AdminTray.ico")
	EndIf
	TraySetToolTip("Быстрый запуск с правами администратора") ; Quick launch with admin rights
	Local $open = TrayCreateItem("Открыть") ; Open with admin rights
		TrayItemSetOnEvent($open, "open")
		TrayItemSetState($open, $TRAY_DEFAULT)
	TrayCreateItem("")
	Local $explorer_menu = TrayCreateMenu("Проводник") ; Explorer
		Local $MyComputer = TrayCreateItem("Мой компьютер", $explorer_menu) ; My Computer
			TrayItemSetState($MyComputer, $TRAY_DEFAULT)
			TrayItemSetOnEvent($MyComputer, "MyComputer")
			TraySetOnEvent($TRAY_EVENT_PRIMARYDOWN, "MyComputer")
		Local $cpl = TrayCreateItem("Панель управления", $explorer_menu) ; Control Panel
			TrayItemSetOnEvent($cpl, "cpl")
		Local $UserProfileDir = TrayCreateItem("Папка профилей", $explorer_menu) ; User profiles folder
			TrayItemSetOnEvent($UserProfileDir, "UserProfileDir")
		Local $EmptyRecycleBin = TrayCreateItem("Отчистить корзину", $explorer_menu) ; Empty Recycle Bin
			TrayItemSetOnEvent($EmptyRecycleBin, "EmptyRecycleBin")
		Local $RestartShell = TrayCreateItem("Перезагрузить", $explorer_menu) ; Restart shell (explorer)
			TrayItemSetOnEvent($RestartShell, "RestartShell")
	Local $cpl_menu = TrayCreateMenu("Панель управления") ; Control Panel
		Local $firewall = TrayCreateItem("Брандмауэр Windows", $cpl_menu) ; Firewall
			TrayItemSetOnEvent($firewall, "firewall")
			If @OSVersion == "WIN_2000" Then TrayItemSetState($firewall, $TRAY_DISABLE)
		Local $timedate = TrayCreateItem("Дата и время", $cpl_menu) ; Date and Time
			TrayItemSetOnEvent($timedate, "timedate")
		Local $inetcpl = TrayCreateItem("Свойства обозревателя", $cpl_menu) ; IE Options
			TrayItemSetOnEvent($inetcpl, "inetcpl")
		Local $sysdm = TrayCreateItem("Свойства системы", $cpl_menu) ; System Properties
			TrayItemSetOnEvent($sysdm, "sysdm")
		Local $NetworkConnections = TrayCreateItem("Сетевые подключения", $cpl_menu) ; NetworkConnections
			TrayItemSetOnEvent($NetworkConnections, "NetworkConnections")
			TrayItemSetState($NetworkConnections, $TRAY_DEFAULT)
		Local $appwiz = TrayCreateItem("Установка и удаление программ", $cpl_menu) ; Add/Remove programs
			TrayItemSetOnEvent($appwiz, "appwiz")
		Local $hdwwiz = TrayCreateItem("Установка оборудования", $cpl_menu) ; Add hardware wizard
			TrayItemSetOnEvent($hdwwiz, "hdwwiz")
		Local $nusrmgr = TrayCreateItem("Учетные записи пользователей", $cpl_menu) ; User Accounts (nusrmgr)
			TrayItemSetOnEvent($nusrmgr, "nusrmgr")
	Local $util_menu = TrayCreateMenu("Утилиты") ; Utilities
		Local $regedit = TrayCreateItem("Редактор реестра", $util_menu) ; Regedit
			TrayItemSetOnEvent($regedit, "regedit")
		Local $gpedit = TrayCreateItem("Групповая политика", $util_menu) ; Group policy editor
			TrayItemSetOnEvent($gpedit, "gpedit")
		Local $taskmgr = TrayCreateItem("Диспетчер задач", $util_menu) ; Task manager
			TrayItemSetOnEvent($taskmgr, "taskmgr")
		Local $cmd = TrayCreateItem("Командная строка", $util_menu) ; cmd.exe
			TrayItemSetOnEvent($cmd, "cmd")
		Local $msconfig = TrayCreateItem("Настройка системы", $util_menu) ; MsConfig
			TrayItemSetOnEvent($msconfig, "msconfig")
			If @OSVersion == "WIN_2000" Then TrayItemSetState($msconfig, $TRAY_DISABLE)
		Local $msinfo32 = TrayCreateItem("Сведения о системе", $util_menu) ; MsInfo32
			TrayItemSetOnEvent($msinfo32, "msinfo32")
		Local $compmgmt = TrayCreateItem("Управление компьютером", $util_menu) ; Computer management
			TrayItemSetState($compmgmt, $TRAY_DEFAULT)
			TrayItemSetOnEvent($compmgmt, "compmgmt")
	TrayCreateItem("")
		Local $DropBoxMenu = TrayCreateMenu("Корзина запуска") ; Drop-Box
			Local $DropBoxPlacementMenu = TrayCreateMenu("Поместить", $DropBoxMenu) ; Move to
				Local $MoveBottomRight = TrayCreateItem("Внизу-Справа", $DropBoxPlacementMenu) ; Bottom-Right
					TrayItemSetOnEvent($MoveBottomRight, "MoveBottomRight")
					TrayItemSetState($MoveBottomRight, $TRAY_DEFAULT)
				Local $MoveBottomCenter = TrayCreateItem("Внизу по центру", $DropBoxPlacementMenu) ; Bottom-Center
					TrayItemSetOnEvent($MoveBottomCenter, "MoveBottomCenter")
				Local $MoveBottomLeft = TrayCreateItem("Внизу-Слева", $DropBoxPlacementMenu) ; Bottom-Left
					TrayItemSetOnEvent($MoveBottomLeft, "MoveBottomLeft")
			Local $ShowHideDropBox = TrayCreateItem("Отобразить/Скрыть", $DropBoxMenu) ; Show/Hide
				TrayItemSetOnEvent($ShowHideDropBox, "ShowHideDropBox")
				TrayItemSetState($ShowHideDropBox, $TRAY_DEFAULT)
			Local $DropBoxTransparency = TrayCreateItem("Прозрачность", $DropBoxMenu) ; Transparency
				TrayItemSetOnEvent($DropBoxTransparency, "ShowHideDropBoxTrans")
		; Local $TrayIconMenu = TrayCreateMenu("Область уведомлений") ; Tray Icon
			; Local $ShowHideTrayIcon = TrayCreateItem("Отобразить/Скрыть", $TrayIconMenu) ; Show/Hide
				; TrayItemSetOnEvent($ShowHideTrayIcon, "ShowHideTrayIcon")
				; TrayItemSetState($ShowHideTrayIcon, $TRAY_DEFAULT)
		Local $info = TrayCreateItem("Информация") ; Information
			TrayItemSetOnEvent($info, "info")
	TrayCreateItem("")
		Local $exit = TrayCreateItem("Выход") ; Exit
			TrayItemSetOnEvent($exit, "_Exit")
EndFunc	;Tray

Func GUI()
	Local $aDesktopPos = ControlGetPos("[CLASS:Progman]", "", "[CLASS:SysListView32; INSTANCE:1]")
	Local $aGetSystemMetrics = GetSystemMetrics()
		Local $PositionX = $aDesktopPos[2] + $aDesktopPos[0] - $DropBoxSizeX - ($aGetSystemMetrics[0] + $aGetSystemMetrics[1]) * 2
		Local $PositionY = $aDesktopPos[3] + $aDesktopPos[1] - $DropBoxSizeY - ($aGetSystemMetrics[2] + $aGetSystemMetrics[3]) * 2
	Global $DropBoxHWND = GUICreate("", $DropBoxSizeX, $DropBoxSizeY, $PositionX, $PositionY, BitOR($WS_POPUP, $WS_VISIBLE, $WS_CLIPSIBLINGS), BitOR($WS_EX_TOOLWINDOW, $WS_EX_ACCEPTFILES, $WS_EX_TOPMOST))
		If @Compiled == 1 Then
			Local $DropBoxIcon = GUICtrlCreateIcon(@AutoItExe, 99, 0, 0, 48, 48, 0, $GUI_WS_EX_PARENTDRAG)
		Else
			Local $DropBoxIcon = GUICtrlCreateIcon(@ScriptDir & "\Res\AdminTray.ico", -1, 0, 0, 48, 48, 0, $GUI_WS_EX_PARENTDRAG)
		EndIf
			GUICtrlSetState($DropBoxIcon, $GUI_DROPACCEPTED)
		GUIRegisterMsg($WM_NCHITTEST, "WM_NCHITTEST")
		GUISetOnEvent($GUI_EVENT_DROPPED, "GUI_EVENT_DROPPED", $DropBoxHWND)
		GUISetOnEvent($GUI_EVENT_SECONDARYDOWN, "PopupMenu", $DropBoxHWND)
		Local $contextmenu = GUICtrlCreateContextMenu($DropBoxIcon)
		Global $hMenu = GUICtrlGetHandle($contextmenu)
			Local $open = GUICtrlCreateMenuItem("Открыть", $contextmenu) ; Open with admin rights
				GUICtrlSetOnEvent($open, "open")
				GUICtrlSetState($open, $GUI_DEFBUTTON)
			GUICtrlCreateMenuItem("", $contextmenu)
			Local $explorer_menu = GUICtrlCreateMenu("Проводник", $contextmenu) ; Explorer
				Local $MyComputer = GUICtrlCreateMenuItem("Мой компьютер", $explorer_menu) ; My Computer
					GUICtrlSetState($MyComputer, $GUI_DEFBUTTON)
					GUICtrlSetOnEvent($MyComputer, "MyComputer")
				Local $cpl = GUICtrlCreateMenuItem("Панель управления", $explorer_menu) ; Control Panel
					GUICtrlSetOnEvent($cpl, "cpl")
				Local $UserProfileDir = GUICtrlCreateMenuItem("Папка профилей", $explorer_menu) ; User profiles folder
					GUICtrlSetOnEvent($UserProfileDir, "UserProfileDir")
				Local $EmptyRecycleBin = GUICtrlCreateMenuItem("Отчистить корзину", $explorer_menu) ; Empty Recycle Bin
					GUICtrlSetOnEvent($EmptyRecycleBin, "EmptyRecycleBin")
				Local $RestartShell = GUICtrlCreateMenuItem("Перезагрузить", $explorer_menu) ; Restart shell (explorer)
					GUICtrlSetOnEvent($RestartShell, "RestartShell")
			Local $cpl_menu = GUICtrlCreateMenu("Панель управления", $contextmenu) ; Control Panel
				Local $firewall = GUICtrlCreateMenuItem("Брандмауэр Windows", $cpl_menu) ; Firewall
					GUICtrlSetOnEvent($firewall, "firewall")
					If @OSVersion == "WIN_2000" Then GUICtrlSetState($firewall, $GUI_DISABLE)
				Local $timedate = GUICtrlCreateMenuItem("Дата и время", $cpl_menu) ; Date and Time
					GUICtrlSetOnEvent($timedate, "timedate")
				Local $inetcpl = GUICtrlCreateMenuItem("Свойства обозревателя", $cpl_menu) ; IE Options
					GUICtrlSetOnEvent($inetcpl, "inetcpl")
				Local $sysdm = GUICtrlCreateMenuItem("Свойства системы", $cpl_menu) ; System Properties
					GUICtrlSetOnEvent($sysdm, "sysdm")
				Local $NetworkConnections = GUICtrlCreateMenuItem("Сетевые подключения", $cpl_menu) ; NetworkConnections
					GUICtrlSetOnEvent($NetworkConnections, "NetworkConnections")
					GUICtrlSetState($NetworkConnections, $GUI_DEFBUTTON)
				Local $appwiz = GUICtrlCreateMenuItem("Установка и удаление программ", $cpl_menu) ; Add/Remove programs
					GUICtrlSetOnEvent($appwiz, "appwiz")
				Local $hdwwiz = GUICtrlCreateMenuItem("Установка оборудования", $cpl_menu) ; Add hardware wizard
					GUICtrlSetOnEvent($hdwwiz, "hdwwiz")
				Local $nusrmgr = GUICtrlCreateMenuItem("Учетные записи пользователей", $cpl_menu) ; User Accounts (nusrmgr)
					GUICtrlSetOnEvent($nusrmgr, "nusrmgr")
			Local $util_menu = GUICtrlCreateMenu("Утилиты", $contextmenu) ; Utilities
				Local $regedit = GUICtrlCreateMenuItem("Редактор реестра", $util_menu) ; Regedit
					GUICtrlSetOnEvent($regedit, "regedit")
				Local $gpedit = GUICtrlCreateMenuItem("Групповая политика", $util_menu) ; Group policy editor
					GUICtrlSetOnEvent($gpedit, "gpedit")
				Local $taskmgr = GUICtrlCreateMenuItem("Диспетчер задач", $util_menu) ; Task manager
					GUICtrlSetOnEvent($taskmgr, "taskmgr")
				Local $cmd = GUICtrlCreateMenuItem("Командная строка", $util_menu) ; cmd.exe
					GUICtrlSetOnEvent($cmd, "cmd")
				Local $msconfig = GUICtrlCreateMenuItem("Настройка системы", $util_menu) ; MsConfig
					GUICtrlSetOnEvent($msconfig, "msconfig")
					If @OSVersion == "WIN_2000" Then GUICtrlSetState($msconfig, $GUI_DISABLE)
				Local $msinfo32 = GUICtrlCreateMenuItem("Сведения о системе", $util_menu) ; MsInfo32
					GUICtrlSetOnEvent($msinfo32, "msinfo32")
				Local $compmgmt = GUICtrlCreateMenuItem("Управление компьютером", $util_menu) ; Computer management
					GUICtrlSetState($compmgmt, $GUI_DEFBUTTON)
					GUICtrlSetOnEvent($compmgmt, "compmgmt")
			GUICtrlCreateMenuItem("", $contextmenu)
				Local $DropBoxMenu = GUICtrlCreateMenu("Корзина запуска", $contextmenu) ; Drop-Box
					Local $DropBoxPlacementMenu = GUICtrlCreateMenu("Поместить", $DropBoxMenu) ; Move to
						Local $MoveBottomRight = GUICtrlCreateMenuItem("Внизу-Справа", $DropBoxPlacementMenu) ; Bottom-Right
							GUICtrlSetOnEvent($MoveBottomRight, "MoveBottomRight")
							GUICtrlSetState($MoveBottomRight, $GUI_DEFBUTTON)
						Local $MoveBottomCenter = GUICtrlCreateMenuItem("Внизу по центру", $DropBoxPlacementMenu) ; Bottom-Center
							GUICtrlSetOnEvent($MoveBottomCenter, "MoveBottomCenter")
						Local $MoveBottomLeft = GUICtrlCreateMenuItem("Внизу-Слева", $DropBoxPlacementMenu) ; Bottom-Left
							GUICtrlSetOnEvent($MoveBottomLeft, "MoveBottomLeft")
					Local $ShowHideDropBox = GUICtrlCreateMenuItem("Отобразить/Скрыть", $DropBoxMenu) ; Show/Hide
						GUICtrlSetOnEvent($ShowHideDropBox, "ShowHideDropBox")
						GUICtrlSetState($ShowHideDropBox, $GUI_DEFBUTTON)
					Local $DropBoxTransparency = GUICtrlCreateMenuItem("Прозрачность", $DropBoxMenu) ; Transparency
						GUICtrlSetOnEvent($DropBoxTransparency, "ShowHideDropBoxTrans")
				; Local $TrayIconMenu = GUICtrlCreateMenu("Область уведомлений", $contextmenu) ; Tray Icon
					; Local $ShowHideTrayIcon = GUICtrlCreateMenuItem("Отобразить/Скрыть", $TrayIconMenu) ; Show/Hide
						; GUICtrlSetOnEvent($ShowHideTrayIcon, "ShowHideTrayIcon")
						; GUICtrlSetState($ShowHideTrayIcon, $GUI_DEFBUTTON)
				Local $info = GUICtrlCreateMenuItem("Информация", $contextmenu) ; Information
					GUICtrlSetOnEvent($info, "info")
			GUICtrlCreateMenuItem("", $contextmenu)
				Local $exit = GUICtrlCreateMenuItem("Выход", $contextmenu) ; Exit
					GUICtrlSetOnEvent($exit, "_Exit")
		WinSetTrans($DropBoxHWND, "", 200)
	GUISetState(@SW_SHOW, $DropBoxHWND)
	Global $DropBoxTransHWND = GUICreate("", 276, 50, -1, -1, $WS_POPUP + $WS_BORDER, BitOR($WS_EX_TOOLWINDOW, $WS_EX_TOPMOST))
			Local $Frame = GUICtrlCreateGraphic(0, 0, 276, 22)
				GUICtrlSetGraphic($Frame, $GUI_GR_COLOR, 0xc0c0ff)
				GUICtrlSetGraphic($Frame, $GUI_GR_RECT, 0, 0, 276, 22)
				GUICtrlSetGraphic($Frame, $GUI_GR_MOVE, 0, 22)
				GUICtrlSetGraphic($Frame, $GUI_GR_COLOR, 0x0)
				GUICtrlSetGraphic($Frame, $GUI_GR_LINE, 276, 22)
				GUICtrlSetBkColor($Frame, 0xc0c0ff)
			Local $DropBoxTransLabelStatic = GUICtrlCreateLabel("Уровень прозрачности корзины запуска", 6, 4, 240, 18, $SS_CENTER, $GUI_WS_EX_PARENTDRAG)
				GUICtrlSetState($DropBoxTransLabelStatic, $GUI_ONTOP)
				GUICtrlSetBkColor($DropBoxTransLabelStatic, 0xc0c0ff)
			Local $DropBoxTransHideButton = GUICtrlCreateButton("X", 256, 2, 18, 18, BitOr($BS_CENTER, $BS_FLAT, $BS_VCENTER, $BS_NOTIFY))
				GUICtrlSetState($DropBoxTransHideButton, $GUI_ONTOP)
				GUICtrlSetOnEvent($DropBoxTransHideButton, "ShowHideDropBoxTrans")
			Global $DropBoxTransSlider = GUICtrlCreateSlider(2, 24, 272, 24, $TBS_NOTICKS)
				GUICtrlSetState($DropBoxTransSlider, $GUI_ONTOP + $GUI_FOCUS + $GUI_DEFBUTTON)
				GUICtrlSetTip($DropBoxTransSlider, "200")
				GUICtrlSetLimit($DropBoxTransSlider, 255, 0)
				GUICtrlSetData($DropBoxTransSlider, 200)
		GUIRegisterMsg($WM_HSCROLL, "WM_HSCROLL")
		GUIRegisterMsg($WM_NCHITTEST, "WM_NCHITTEST")
		GUISetOnEvent($GUI_EVENT_CLOSE, "ShowHideDropBoxTrans", $DropBoxTransHWND)
		; GUISetOnEvent($GUI_EVENT_SECONDARYDOWN, "ShowHideDropBoxTrans", $DropBoxTransHWND)
	GUISetState(@SW_HIDE, $DropBoxTransHWND)
EndFunc	;GUI

Func PopupMenu()
	Local $aMousePosRet = MouseGetPos()
	TrackPopupMenu($DropBoxHWND, $hMenu, $aMousePosRet[0], $aMousePosRet[1])
EndFunc ;PopupMenu

Func open()
	TraySetClick(0)
	TraySetOnEvent($TRAY_EVENT_PRIMARYDOWN, "")
	Local $FilePath = FileOpenDialog("Открыть с правами администратора...", @DesktopDir & "\", "Все файлы (*.*)", 1 + 2 + 8) ; Open with admin rights
		If Not @error Then
			ShellExecute($FilePath)
		EndIf
	TraySetOnEvent($TRAY_EVENT_PRIMARYDOWN, "MyComputer")
	TraySetClick(8)
EndFunc	;open

Func MyComputer()
	_Explorer("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", False)
EndFunc	;MyComputer

Func UserProfileDir()
	_Explorer($sProfilesDirectory, False)
EndFunc	;MyComputer

Func EmptyRecycleBin()
	FileRecycleEmpty()
EndFunc	;EmptyRecycleBin

Func RestartShell()
	If _ProcessGetName($ExplorerPID) == "explorer.exe" Then ProcessClose($ExplorerPID)
	ProcessClose("explorer.exe")
EndFunc	;RestartShell

Func cpl()
	If @OSVersion == "WIN_2000" Or @OSVersion == "WIN_XP" Or @OSVersion == "WIN_2003" Then
		_Explorer("::{20D04FE0-3AEA-1069-A2D8-08002B30309D}\::{21EC2020-3AEA-1069-A2DD-08002B30309D}", False)
	Else
		_Explorer("::{21EC2020-3AEA-1069-A2DD-08002B30309D}", False)
	EndIf
EndFunc	;cpl

Func NetworkConnections()
	_Explorer("::{7007ACC7-3202-11D1-AAD2-00805FC1270E}", False)
EndFunc	;NetworkConnections

Func firewall()
	Run("rundll32.exe shell32.dll, Control_RunDLL firewall.cpl")
EndFunc	;firewall

Func timedate()
	Run("rundll32.exe shell32.dll, Control_RunDLL timedate.cpl")
EndFunc	;timedate

Func inetcpl()
	Run("rundll32.exe shell32.dll, Control_RunDLL inetcpl.cpl")
EndFunc	;inetcpl

Func sysdm()
	Run("rundll32.exe shell32.dll, Control_RunDLL sysdm.cpl")
EndFunc	;sysdm

Func appwiz()
	Run("rundll32.exe shell32.dll, Control_RunDLL appwiz.cpl")
EndFunc	;appwiz

Func hdwwiz()
	Run("rundll32.exe shell32.dll, Control_RunDLL hdwwiz.cpl")
EndFunc	;hdwwiz

Func nusrmgr()
	If @OSVersion <> "WIN_2000" Then
		Run("rundll32.exe shell32.dll, Control_RunDLL nusrmgr.cpl")
	ElseIf @OSVersion == "WIN_2000" Then
		Run("rundll32.exe netplwiz.dll,UsersRunDll")
	EndIf
EndFunc	;nusrmgr

Func regedit()
	Run(@WindowsDir & "\regedit.exe")
EndFunc	;regedit

Func gpedit()
	Run(@SystemDir & "\mmc.exe " & @SystemDir & "\gpedit.msc")
EndFunc	;gpedit

Func taskmgr()
	Run(@SystemDir & "\taskmgr.exe")
EndFunc	;taskmgr

Func cmd()
	Run(@SystemDir & "\cmd.exe")
EndFunc	;cmd

Func msconfig()
	Run(@WindowsDir & "\pchealth\helpctr\binaries\msconfig.exe")
EndFunc	;msconfig

Func msinfo32()
	Run(@CommonFilesDir & "\Microsoft Shared\MSInfo\msinfo32.exe")
EndFunc	;msinfo32

Func compmgmt()
	Run(@SystemDir & "\mmc.exe " & @SystemDir & "\compmgmt.msc /s")
EndFunc	;compmgmt

Func MoveBottomRight()
	Local $aDesktopPos = ControlGetPos("[CLASS:Progman]", "", "[CLASS:SysListView32; INSTANCE:1]")
	Local $aGetSystemMetrics = GetSystemMetrics()
	Local $PositionX = $aDesktopPos[2] + $aDesktopPos[0] - $DropBoxSizeX - ($aGetSystemMetrics[0] + $aGetSystemMetrics[1]) * 2
	Local $PositionY = $aDesktopPos[3] + $aDesktopPos[1] - $DropBoxSizeY - ($aGetSystemMetrics[2] + $aGetSystemMetrics[3]) * 2
	WinMove($DropBoxHWND, "", $PositionX, $PositionY)
EndFunc	;MoveBottomRight

Func MoveBottomCenter()
	Local $aDesktopPos = ControlGetPos("[CLASS:Progman]", "", "[CLASS:SysListView32; INSTANCE:1]")
	Local $aGetSystemMetrics = GetSystemMetrics()
	Local $PositionX = (($aDesktopPos[2] + $aDesktopPos[0] - $DropBoxSizeX) / 2)
	Local $PositionY = $aDesktopPos[3] + $aDesktopPos[1] - $DropBoxSizeY - ($aGetSystemMetrics[2] + $aGetSystemMetrics[3]) * 2
	WinMove($DropBoxHWND, "", $PositionX, $PositionY)
EndFunc	;MoveBottomCenter

Func MoveBottomLeft()
	Local $aDesktopPos = ControlGetPos("[CLASS:Progman]", "", "[CLASS:SysListView32; INSTANCE:1]")
	Local $aGetSystemMetrics = GetSystemMetrics()
	Local $PositionX = $aDesktopPos[0] + ($aGetSystemMetrics[0] + $aGetSystemMetrics[1]) * 2
	Local $PositionY = $aDesktopPos[3] + $aDesktopPos[1] - $DropBoxSizeY - ($aGetSystemMetrics[2] + $aGetSystemMetrics[3]) * 2
	WinMove($DropBoxHWND, "", $PositionX, $PositionY)
EndFunc	;MoveBottomLeft

Func ShowHideDropBox()
	Local $WinState = WinGetState($DropBoxHWND, "")
	If BitAnd($WinState, 2) Then
		GUISetState(@SW_HIDE, $DropBoxHWND)
	Else
		GUISetState(@SW_SHOW, $DropBoxHWND)
	EndIf
EndFunc	;ShowHideDropBox

Func ShowHideDropBoxTrans()
	Local $WinState = WinGetState($DropBoxTransHWND, "")
	If BitAnd($WinState, 2) Then
		GUISetState(@SW_HIDE, $DropBoxTransHWND)
	Else
		GUISetState(@SW_SHOW, $DropBoxTransHWND)
		GUISetState(@SW_SHOW, $DropBoxHWND)
		ControlFocus($DropBoxTransHWND, "", "[CLASS:msctls_trackbar32; INSTANCE:1]")
	EndIf
EndFunc	;ShowHideDropBoxTrans

; Func ShowHideTrayIcon()
	; If $TrayState == True Then
		; $TrayState = False
		; TraySetClick(0)
		; TraySetState(2)
	; ElseIf $TrayState == False Then
		; $TrayState = True
		; TraySetClick(8)
		; TraySetState(1)
	; EndIf
; EndFunc	;ShowHideDropBox

Func info()
	TraySetClick(0)
	TraySetOnEvent($TRAY_EVENT_PRIMARYDOWN, '')
	Msgbox(0, 'Информация', _
		'Кнопки быстрого запуска:' _
		& @CRLF & @TAB & 'Мой компьютер - "Win"+"1"' & @CRLF & @TAB & 'Панель управления - "Win"+"2"' _
		& @CRLF & @TAB & 'Сетевые подключения - "Win"+"3"' & @CRLF & @TAB & 'Перезагрузить оболочку - "Win"+"0"' _
		& @CRLF & @TAB & 'Открыть - "Win"+"="' & @CRLF & @TAB & 'Командная строка - "Win"+"-"' & @CRLF & @TAB & 'Диспетчер задач - "Win"+"Space"' _
		& @CRLF & @TAB & 'Выход - "Ctrl"+"Alt"+"Break"' _
		& @CRLF & 'Командная строка (параметры):' _
		& @CRLF & @TAB & '"image.exe" "-AccountCreds" "UserName" "Password"', 60)
		; HotKey list: 'MyComputer - "Win"+"1"', 'Control Panel - "Win"+"2"', 'NetworkConnections - "Win"+"3"', 'Restart Shell - "Win"+"0"',
					; 'Open - "Win"+"="', 'CMD - "Win"+"-"', 'Task manager - "Win"+"Space"', 'Exit - "Ctrl"+"Alt"+"Break"'
		; Application parameters: '"image.exe" "-AccountCreds" "UserName" "Password"'
	TraySetOnEvent($TRAY_EVENT_PRIMARYDOWN, 'MyComputer')
	TraySetClick(8)
EndFunc	;info

Func _Exit()
	DllClose($DLL)
	SplashOff()
	If _ProcessGetName($ExplorerPID) == "explorer.exe" Then ProcessClose($ExplorerPID)
	Exit
EndFunc	;_Exit

Func WM_NCHITTEST($hWnd, $iMsg, $iwParam, $ilParam)
	If (($hWnd = $DropBoxHWND) Or ($hWnd = $DropBoxTransHWND)) And ($iMsg = $WM_NCHITTEST) Then
		Return $HTCAPTION
	Else
		Return $GUI_RUNDEFMSG
	EndIf
EndFunc

Func WM_HSCROLL($hWnd, $iMsg, $iwParam, $ilParam)
	If ($hWnd = $DropBoxTransHWND) And ($iMsg = $WM_HSCROLL) Then
		WinSetTrans($DropBoxHWND, "", GUICtrlRead($DropBoxTransSlider))
		GUICtrlSetTip($DropBoxTransSlider, GUICtrlRead($DropBoxTransSlider))
		Return $GUI_RUNDEFMSG
	EndIf
EndFunc

Func GUI_EVENT_DROPPED()
	;If @GUI_DragID == -1 And @GUI_DropID == 3 Then
	If @GUI_DragID == -1 Then
		Dim $szDrive, $szDir, $szFName, $szExt
		Local $aPath = _PathSplit(@GUI_DragFile, $szDrive, $szDir, $szFName, $szExt)
		If StringLen($szExt) == 0 Then
			_Explorer($szDrive & $szDir & $szFName, False)
		Else
			ShellExecute(@GUI_DragFile, "", $szDrive & $szDir)
		EndIf
	EndIf
EndFunc	;GUI_EVENT_DROPPED

Func TrackPopupMenu($hWnd, $hMenu, $x, $y)
	DllCall("user32.dll", "int", "TrackPopupMenuEx", "hwnd", $hMenu, "int", 0, "int", $x, "int", $y, "hwnd", $hWnd, "ptr", 0)
EndFunc ;TrackPopupMenu

Func GetSystemMetrics()
	Local $aRet[4]
	$aRet[0] = _WinAPI_GetSystemMetrics(7) ; SM_CXFIXEDFRAME
	$aRet[1] = _WinAPI_GetSystemMetrics(5) ; SM_CXBORDER
	$aRet[2] = _WinAPI_GetSystemMetrics(8) ; SM_CYFIXEDFRAME
	$aRet[3] = _WinAPI_GetSystemMetrics(6) ; SM_CYBORDER
	Return $aRet
EndFunc	;GetSystemMetrics

Func _Explorer($_ExplorerPath, $ShowWarningSplash)
	If @OSVersion == "WIN_2000" Or $ShowWarningSplash == True Then
		$SplashTimerDiff = TimerDiff($SplashTimerInit)
		If $SplashTimerDiff > 4000 Then
			$SplashTimerInit = TimerInit()
			Local $iIndex = $GWL_EXSTYLE
			Local $iValue = $WS_EX_WINDOWEDGE + $WS_EX_TOPMOST + $WS_EX_TRANSPARENT
			Local $HWND = SplashTextOn("", "Не забываем обновлять окно проводника!", 350, 24, -1, -1, 33)
			; Don't forget to manually refresh explorer window
			_WinAPI_SetWindowLong($HWND, $iIndex, $iValue)
			WinSetTrans($HWND, "", 200)
		EndIf
	EndIf
	If _ProcessGetName($ExplorerPID) == "explorer.exe" Or @OSVersion == "WIN_2000" Then
		Run(@WindowsDir & "\explorer.exe /n," & $_ExplorerPath)
	Else
		$ExplorerPID = Run(@WindowsDir & "\explorer.exe /n,/separate," & $_ExplorerPath)
	EndIf
EndFunc	;_Explorer

Func _ProcessGetName($PID)
	Local $aProcesses = ProcessList()
	If Not @error And $aProcesses[0][0] <> 0 Then
		For $i = 1 To $aProcesses[0][0]
			If $aProcesses[$i][1] == $PID Then Return $aProcesses[$i][0]
		Next
	Else
		Return SetError(1, 1, 0)
	EndIf
EndFunc	;_ProcessGetName

Func _SleepANDTrayNotify($Delay)
	If @Compiled == 1 Then
		TraySetIcon(@AutoItExe, 162)
		TraySetClick(0)
		Sleep($Delay)
		TraySetClick(8)
		TraySetIcon(@AutoItExe, 99)
	Else
		TraySetIcon(@ScriptDir & "\Res\blank.ico")
		TraySetClick(0)
		Sleep($Delay)
		TraySetClick(8)
		TraySetIcon(@ScriptDir & "\Res\AdminTray.ico")
	EndIf
EndFunc	;_SleepANDTrayNotify

Func _MakeMeAdmin_RunAs($UserName, $Password, $ExecPath, $WorkingDir)
	Local $ErrMsg, $ErrExtendedMacro = Int(0)
	_NetLocalGroupAddMember($UserName, $LocalAdminGroupName)
	If Not @error Then
		_SleepANDTrayNotify(100)
		Local $PID = RunAs($UserName, @ComputerName, $Password, 0, $ExecPath, $WorkingDir)
		If @error And $PID == 0 Then
			$ErrMsg &= _WinAPI_GetLastErrorMessage() & @CRLF
			$ErrExtendedMacro = Int(2)
		EndIf
		_SleepANDTrayNotify(50)
		_NetLocalGroupDelMembers($UserName, $LocalAdminGroupName)
		If @error Then
			$ErrMsg &= _WinAPI_GetLastErrorMessage() & @CRLF
			$ErrExtendedMacro = Int(3)
		EndIf
	Else
		$ErrMsg &= _WinAPI_GetLastErrorMessage() & @CRLF
		$ErrExtendedMacro = Int(1)
	EndIf
	If $ErrExtendedMacro == 0 Then
		Return $PID
	Else
		Return SetError(1, $ErrExtendedMacro, $ErrMsg)
	EndIf
EndFunc	;_MakeMeAdmin_RunAs

Func _ExpandEnvironmentVariable($sString)
	Local $aRet = StringRegExp($sString, "(?i)%[[:alnum:]]*%", 3)
	For $i = 0 to (UBound($aRet) - 1)
		Local $EnvironmentVariable = StringMid($aRet[$i], 2, (StringLen($aRet[$i]) - 2))
		Local $RetrievedEnvironmentVariable = EnvGet($EnvironmentVariable)
		$sString = StringReplace($sString, $aRet[$i], $RetrievedEnvironmentVariable)
	Next
	$sString = StringReplace($sString, "/", "\")
	Return $sString
EndFunc	;_ExpandEnvironmentVariable

Func _LogonUser($sUsername, $sPassword, $sServer = '.') ; Returns True if user exists
	Local $stToken = DllStructCreate("int")
	Local $aRet = DllCall("advapi32.dll", "int", "LogonUser", _
			"str", $sUsername, "str", $sServer, "str", $sPassword, "dword", 3, "dword", 0, "ptr", DllStructGetPtr($stToken))
	; Local $hToken = DllStructGetData($stToken, 1)
	If @error Then
		Local $ErrorMsg = _WinAPI_GetLastErrorMessage()
		Return SetError(1, @error, $ErrorMsg)
	ElseIf $aRet[0] <> 0 Then
		Return True
	ElseIf $aRet[0] == 0 Then
		Return False
	EndIf
EndFunc	;_LogonUser

Func _IsAdmin($sUserName)
	Local $aLocalGroupNames = _NetUserGetLocalGroups($sUserName)
	If @error Then
		MsgBox(0, "Ошибка, _NetUserGetLocalGroups", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aArraySearch = _ArraySearch($aLocalGroupNames, $sLocalAdminGroupName, 1)
	If $aArraySearch == -1 Then
		Return False
	Else
		Return True
	EndIf
EndFunc	;_IsAdmin

Func _NetLocalGroupAddMember($sUsername, $sGroup, $sServer = '')
	Local $twUser = DllStructCreate("wchar["& StringLen($sUsername)+1 &"]")
	Local $tpUser = DllStructCreate("ptr")
	DllStructSetData($twUser, 1, $sUsername)
	DllStructSetData($tpUser, 1, DllStructGetPtr($twUser))
	Local $aRet = DllCall("netapi32.dll", "int", "NetLocalGroupAddMembers", "wstr", $sServer, "wstr", _
		$sGroup, "int", 3, "ptr", DllStructGetPtr($tpUser), "int", 1 )
	If $aRet[0] Then Return SetError(1, $aRet[0], False)
	Return True
EndFunc	;_NetLocalGroupAddMember

Func _NetLocalGroupDelMembers($sUsername, $sGroup, $sServer = '')
	Local $twUser = DllStructCreate("wchar["& StringLen($sUsername)+1 &"]")
	Local $tpUser = DllStructCreate("ptr")
	DllStructSetData($twUser, 1, $sUsername)
	DllStructSetData($tpUser, 1, DllStructGetPtr($twUser))
	Local $aRet = DllCall("netapi32.dll", "int", "NetLocalGroupDelMembers", "wstr", $sServer, "wstr", _
		$sGroup, "int", 3, "ptr", DllStructGetPtr($tpUser), "int", 1 )
	If $aRet[0] Then Return SetError(1, $aRet[0], False)
	Return True
EndFunc	;_NetLocalGroupDelMembers

Func _NetUserEnum($sServer = "")
	Local $tBufPtr = DllStructCreate("ptr")
	Local $tEntriesRead = DllStructCreate("dword")
	Local $tTotalEntries = DllStructCreate("dword")
	Local $aRet = DllCall("Netapi32.dll", "int", "NetUserEnum", "wstr", $sServer, "dword", 1, "dword", 2, "ptr", DllStructGetPtr($tBufPtr), _
		"dword", -1, "ptr", DllStructGetPtr($tEntriesRead), "ptr", DllStructGetPtr($tTotalEntries), "ptr", 0 )
	If $aRet[0] Then Return SetError(1, $aRet[0], False)
	Local Const $UF_ACCOUNTDISABLE = 0x2
	Local $iEntriesRead = DllStructGetData($tEntriesRead, 1)
	Local $pBuf = DllStructGetData($tBufPtr, 1)
	Local $aUserEnum[1] = [0]
	Local $sUserInfo1 = "ptr;ptr;dword;dword;ptr;ptr;dword;ptr"
	Local $tUserInfo1 = DllStructCreate ($sUserInfo1)
	Local $zUserInfo1 = DllStructGetSize($tUserInfo1)
	For $i = 1 To $iEntriesRead
		$tUserInfo1 = DllStructCreate($sUserInfo1, $pBuf+($i-1)*$zUserInfo1)
		Local $tName = DllStructCreate("wchar[256]", DllStructGetData($tUserInfo1, 1))
		Local $tFlag = DllStructGetData($tUserInfo1, 7)
		If BitAnd($tFlag, $UF_ACCOUNTDISABLE) = 0 Then
			$aUserEnum[0] += 1
			ReDim $aUserEnum[$aUserEnum[0]+1]
			$aUserEnum[$aUserEnum[0]] = DllStructGetData($tName, 1)
		EndIf
	Next
	DllCall("Netapi32.dll", "int", "NetApiBufferFree", "ptr", $pBuf)
	Return $aUserEnum ; array[0] contains number of elements
EndFunc	;_NetUserEnum

Func _NetUserGetLocalGroups($sUsername, $sServer = "")
	Local CONST $LG_INCLUDE_INDIRECT = 0x1
	Local $tBufPtr = DllStructCreate("ptr")
	Local $ptBufPtr = DllStructGetPtr($tBufPtr)
	Local $tEntriesRead = DllStructCreate("dword")
	Local $ptEntriesRead = DllStructGetPtr($tEntriesRead)
	Local $tTotalEntries = DllStructCreate("dword")
	Local $ptTotalEntries = DllStructGetPtr($tTotalEntries)
	Local $aRet = DllCall("Netapi32.dll", "int", "NetUserGetLocalGroups", "wstr", $sServer, "wstr", $sUsername, "dword", 0, _
		"dword", $LG_INCLUDE_INDIRECT, "ptr", $ptBufPtr, "dword", -1, "ptr", $ptEntriesRead, "ptr", $ptTotalEntries)
	If $aRet[0] Then Return SetError(1, $aRet[0], False)
	Local $iEntriesRead = DllStructGetData($tEntriesRead, 1)
	Local $pBuf = DllStructGetData($tBufPtr, 1)
	Local $sLocalGroupUsersInfo0 = "ptr"
	Local $tLocalGroupUsersInfo0 = DllStructCreate($sLocalGroupUsersInfo0)
	Local $zLocalGroupUsersInfo0 = DllStructGetSize($tLocalGroupUsersInfo0)
	Local $tLocalGroupName
	Local $aLocalGroupNames[1] = [0]
	For $i = 1 To $iEntriesRead
		$tLocalGroupUsersInfo0 = DllStructCreate($sLocalGroupUsersInfo0, $pBuf + ($i - 1) * $zLocalGroupUsersInfo0)
		$tLocalGroupName = DllStructCreate("wchar[256]", DllStructGetData($tLocalGroupUsersInfo0, 1))
		$aLocalGroupNames[0] += 1
		ReDim $aLocalGroupNames[$aLocalGroupNames[0]+1]
		$aLocalGroupNames[$aLocalGroupNames[0]] = DllStructGetData($tLocalGroupName, 1)
	Next
	DllCall("Netapi32.dll", "int", "NetApiBufferFree", "ptr", $pBuf)
	Return $aLocalGroupNames ; array[0] contains number of elements
EndFunc	;_NetUserGetLocalGroups
