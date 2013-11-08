#cs ----------------------------------------------------------------------------

 AutoIt Version: v3.3.6.1
 Author:		 Serj

 Script Function:
	MakeMeAdmin

#ce ----------------------------------------------------------------------------

#NoTrayIcon
; #RequireAdmin
#include <Array.au3>
#include <ButtonConstants.au3>
#include <Constants.au3>
#include <Crypt.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <Misc.au3>
#include <Security.au3>
#include <StaticConstants.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>
Opt("GUIOnEventMode", 1)
Opt("MustDeclareVars", 1)
OnAutoItExitRegister("_Exit")

Init()

Func OnAutoItStart()
	TraySetClick(0)
	TraySetState(2)
EndFunc ; OnAutoItStart

Func Init()
	Global CONST $vCryptKey = DriveGetSerial(@HomeDrive)
	If @error Then
		MsgBox(0, "Ошибка, DriveGetSerial", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local CONST $AdminGroupSID = "S-1-5-32-544"
		; Administrators [S-1-5-32-544], Users [S-1-5-32-545], Guests [S-1-5-32-546], Power Users [S-1-5-32-547]
	Local $aLocalAdminGroupName = _Security__LookupAccountSid($AdminGroupSID)
	If @error Or $aLocalAdminGroupName[2] <> 4 Then
		MsgBox(0, "Ошибка, LocalAdminGroupName", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aCompSID = _Security__LookupAccountName(@ComputerName)
	If @error Or $aCompSID[2] <> 3 Then
		MsgBox(0, "Ошибка, ComputerSID", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aLocalAdminName = _Security__LookupAccountSid($aCompSID[0] & "-500")
	If @error Or $aLocalAdminName[2] <> 1 Then
		MsgBox(0, "Ошибка, LocalAdminName", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Local $aUserEnum = _NetUserEnum()
	If @error Then
		MsgBox(0, "Ошибка, _NetUserEnum", _WinAPI_GetLastErrorMessage())
		_Exit()
	EndIf
	Global CONST $sLocalAdminName = $aLocalAdminName[0]
	Global CONST $sLocalAdminGroupName = $aLocalAdminGroupName[0]
	If $CmdLine[0] == 5 And $CmdLine[1] == "-RUNasLUA" Then
		RUNasLUA($CmdLine[2], $CmdLine[3], $CmdLine[4], $CmdLine[5])
	ElseIf $CmdLine[0] == 3 And $CmdLine[1] == "-app" Then
		MainGUI($aUserEnum, $CmdLine[2], $CmdLine[3])
	ElseIf $CmdLine[0] == 2 And $CmdLine[1] == "-o" Then
		Open($CmdLine[2])
	Else
		MainGUI($aUserEnum)
	EndIf
EndFunc ; Init

Func MainGUI($aUserEnum, $sAppPath = "", $sParameters = "")
	_EnvSet()
	$sAppPath = _ExpandEnvironmentVariable($sAppPath)
	Local $hGUI = GUICreate("Настройки MakeMeAdmin:", 300, 205, -1, -1, BitOr($WS_CAPTION, $WS_SYSMENU, $WS_MINIMIZEBOX))
		GUICtrlCreateGroup("Настройки запуска программы", 5, 0, 290, 90) ; Program options
			Local $nPath = GUICtrlCreateInput($sAppPath, 10, 15, 255, 20, BitOr($ES_AUTOHSCROLL, $ES_READONLY))
			GUICtrlSetTip($nPath, "Объект") ; Path
			Local $nAppPath = GUICtrlCreateButton("...", 270, 15, 20, 20)
			GUICtrlSetTip($nAppPath, "Обзор") ; Browse
			GUICtrlSetOnEvent($nAppPath, "Browse")
			Local $nParameters = GUICtrlCreateInput($sParameters, 10, 40, 280, 20, $ES_AUTOHSCROLL)
			GUICtrlSetTip($nParameters, "Параметры") ; Parameters
			Local $nHash = GUICtrlCreateInput("", 10, 65, 255, 20, BitOr($ES_AUTOHSCROLL, $ES_READONLY))
			GUICtrlSetTip($nHash, "SHA1") ; SHA1
			Local $nHashRefresh = GUICtrlCreateButton(">>", 270, 65, 20, 20)
			GUICtrlSetOnEvent($nHashRefresh, "HashRefresh")
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateGroup("Пароль для: " & @UserName, 5, 90, 290, 40)
			Local $nLUAPassword = GUICtrlCreateInput("", 10, 105, 280, 20, BitOr($ES_PASSWORD, $ES_AUTOHSCROLL))
			GUICtrlSetTip($nLUAPassword, "Пароль для ограниченной учетной записи") ; Password for LUA
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateGroup("Пароль для: " & $sLocalAdminName, 5, 130, 290, 40) ; Admin user account
			Local $nAdminPassword = GUICtrlCreateInput("", 10, 145, 280, 20, BitOr($ES_PASSWORD, $ES_AUTOHSCROLL))
			GUICtrlSetTip($nAdminPassword, "Пароль для встроенной учетной записи администратора") ; Password for build-in administrator account
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		Local $nExec = GUICtrlCreateButton("Выполнить", 5, 175, 65, 25) ; Execute
		GUICtrlSetOnEvent($nExec, "Exec")
		Local $nOpen = GUICtrlCreateButton("Открыть", 80, 175, 65, 25) ; Open
		GUICtrlSetOnEvent($nOpen, "OpenWrapper")
		Local $nSave = GUICtrlCreateButton("Сохранить", 155, 175, 65, 25) ; Save
		GUICtrlSetOnEvent($nSave, "Save")
		Local $nExit = GUICtrlCreateButton("Выйти", 230, 175, 65, 25) ; Exit
		GUICtrlSetOnEvent($nExit, "_Exit")
		GUISetOnEvent($GUI_EVENT_CLOSE, "_Exit", $hGUI)
	GUISetState(@SW_SHOW)
	While 1
		Sleep(1000)
	WEnd
EndFunc ; MainGUI

Func _EnvSet()
	Local $sProfilesDirectory = RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList", "ProfilesDirectory")
	$sProfilesDirectory = _ExpandEnvironmentVariable($sProfilesDirectory)
	EnvSet("ScriptDir", @ScriptDir )
	EnvSet("AppDataCommonDir", @AppDataCommonDir)
	EnvSet("AppDataDir", @AppDataDir)
	EnvSet("CommonFilesDir", @CommonFilesDir)
	EnvSet("ComSpec", @ComSpec)
	EnvSet("DesktopCommonDir", @DesktopCommonDir)
	EnvSet("DesktopDir", @DesktopDir)
	EnvSet("DocumentsCommonDir", @DocumentsCommonDir)
	EnvSet("FavoritesCommonDir", @FavoritesCommonDir)
	EnvSet("FavoritesDir", @FavoritesDir)
	EnvSet("MyDocumentsDir", @MyDocumentsDir)
	EnvSet("ProfilesDirectory", $sProfilesDirectory)
	EnvSet("ProgramFilesDir", @ProgramFilesDir)
	EnvSet("ProgramsCommonDir", @ProgramsCommonDir)
	EnvSet("ProgramsDir", @ProgramsDir)
	EnvSet("StartMenuCommonDir", @StartMenuCommonDir)
	EnvSet("StartMenuDir", @StartMenuDir)
	EnvSet("StartupCommonDir", @StartupCommonDir)
	EnvSet("StartupDir", @StartupDir)
	EnvSet("SystemDir", @SystemDir)
	EnvSet("TempDir", @TempDir)
	EnvSet("UserProfileDir", @UserProfileDir)
	EnvSet("WindowsDir", @WindowsDir)
	EnvSet("WorkingDir", @WorkingDir)
	EnvUpdate()
EndFunc	;_EnvSet

Func Browse()
	Local $sAppPath = FileOpenDialog("Выбрать приложение:", @WorkingDir & "\", "Programs (*.exe;*.com;*.cmd;*.bat)", 1 + 2, "", @GUI_WINHANDLE) ; Choose application
	If Not @error Then
		Local $hHash = ControlGetHandle(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:3]")
		Local $nHash = _WinAPI_GetDlgCtrlID($hHash)
		ControlSetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:1]", $sAppPath)
		Local $aRet = HashWrapper($sAppPath)
		ControlSetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:3]", StringLower(StringTrimLeft($aRet[0], 2)))
		GUICtrlSetTip($nHash, $aRet[1])
	EndIf
EndFunc	;Browse

Func HashRefresh()
	Local $sAppPath = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:1]")
	If FileExists($sAppPath) Then
		Local $hHash = ControlGetHandle(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:3]")
		Local $nHash = _WinAPI_GetDlgCtrlID($hHash)
		Local $aRet = HashWrapper($sAppPath)
		ControlSetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:3]", StringLower(StringTrimLeft($aRet[0], 2)))
		GUICtrlSetTip($nHash, $aRet[1])
	EndIf
EndFunc	;HashRefresh

Func Exec()
	Local $sAppPath = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:1]")
	Local $sParameters = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:2]")
	Local $sAdminPwd = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:5]")
	Local $sLUAPwd = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:4]")
	HashRefresh()
	Local $bHash = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:3]")
	$bHash = "0x" & StringUpper($bHash)
	Local $sRet = RunAsAdminStub($sAppPath, $sParameters, $bHash, $sAdminPwd, $sLUAPwd)
	If @error Then MsgBox(0, "Ошибка", $sRet, 0, @GUI_WINHANDLE)
EndFunc	;Exec

Func OpenWrapper()
	Local $sCryptPath = FileOpenDialog("Выбрать путь:", @WorkingDir & "\", "(*.crypt)", 1 + 2, "", @GUI_WINHANDLE) ; Choose path
	If @error Then
		MsgBox(0, "Ошибка", "Не выбран Crypt файл!", 0, @GUI_WINHANDLE) ; No Crypt file was chosen!
	Else
		Local $sRet = Open($sCryptPath)
		If @error Then MsgBox(0, "Ошибка", $sRet, 0, @GUI_WINHANDLE)
	EndIf
EndFunc	;OpenWrapper

Func Open($sCryptPath)
	If FileExists($sCryptPath) Then
		Local $bCryptString = FileRead($sCryptPath)
		Local $sCryptString = _Crypt_DecryptData($bCryptString, $vCryptKey, $CALG_3DES)
		If @error Then
			Return SetError(1, False, "Ошибка дешифровки!") ; Failed to decrypt data!
		Else
			$sCryptString = BinaryToString($sCryptString)
			Local $aStringSplit = StringSplit($sCryptString, '¶', 1)
			Local $sRet = RunAsAdminStub($aStringSplit[1], $aStringSplit[2], $aStringSplit[3], $aStringSplit[4], $aStringSplit[5])
			If @error Then
				Return SetError(1, False, $sRet)
			Else
				Return True
			EndIf
		EndIf
	Else
		Return SetError(1, False, "Ошибка чтения Crypt файла!") ; Error reading Crypt file!
	EndIf
EndFunc	;Open

Func Save()
	HashRefresh()
	Local $sAppPath = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:1]")
	Local $sParameters = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:2]")
	Local $bHash = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:3]")
	$bHash = "0x" & StringUpper($bHash)
	Local $sAdminPwd = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:5]")
	Local $sLUAPwd = ControlGetText(@GUI_WINHANDLE, "", "[CLASS:Edit; INSTANCE:4]")
	Local $sCryptString = $sAppPath & "¶" & $sParameters & "¶" & $bHash & "¶" & $sAdminPwd & "¶" & $sLUAPwd
	Local $bCryptString = _Crypt_EncryptData($sCryptString, $vCryptKey, $CALG_3DES)
	If @error Then
		MsgBox(0, "Ошибка", "Ошибка шифрования!", 0, @GUI_WINHANDLE) ; Failed to encrypt data!
	Else
		Local $sAppName = _GetAppName($sAppPath)
		Local $sCryptPath = FileSaveDialog("Выбрать путь:", @WorkingDir & "\", "(*.crypt)", 16, $sAppName & ".crypt", @GUI_WINHANDLE) ; Choose path
		If @error Then
			MsgBox(0, "Ошибка", "Crypt файл не сохранен!", 0, @GUI_WINHANDLE) ; Crypt file save error!
		Else
			Local $hCryptPath = FileOpen($sCryptPath, 2 + 8 + 16)
			If $hCryptPath <> -1 Or @error Then
				FileWrite($hCryptPath, $bCryptString)
				If @error Then MsgBox(0, "Ошибка", "Crypt файл не сохранен!", 0, @GUI_WINHANDLE) ; Crypt file save error!
				FileClose($hCryptPath)
			Else
				MsgBox(0, "Ошибка", "Crypt файл не сохранен!", 0, @GUI_WINHANDLE) ; Crypt file save error!
			EndIf
		EndIf
	EndIf
EndFunc	;Save

Func HashWrapper($sAppPath)
	Local $hTimer = TimerInit()
	Local $bAppHash = _Crypt_HashFile($sAppPath, $CALG_SHA1)
	Local $aRet[2]
	If @error Then
		$aRet[0] = ""
		$aRet[1] = "Вычисление контрольной суммы (SHA1) потерпело неудачу." ; SHA1 hash computation failed.
		SetError(1)
		Return $aRet
	Else
		Local $iTimer = TimerDiff($hTimer)
		$aRet[0] = $bAppHash
		$aRet[1] = "Вычисление контрольной суммы (SHA1) заняло - " & Round($iTimer, 2) & " мс." ; SHA1 hash computation took
		SetError(0)
		Return $aRet
	EndIf
EndFunc	;HashWrapper

Func _Exit()
	Exit
EndFunc	;_Exit

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

Func _StripAppName($sAppPath)
	Dim $szDrive, $szDir, $szFName, $szExt
	Local $aPathRet = _PathSplit($sAppPath, $szDrive, $szDir, $szFName, $szExt)
	If FileExists($aPathRet[2]) == 0 Then
		Return @TempDir
	Else
		Return $aPathRet[1] & $aPathRet[2]
	EndIf
EndFunc	;_StripAppName

Func _GetAppName($sAppPath)
	Dim $szDrive, $szDir, $szFName, $szExt
	Local $aPathRet = _PathSplit($sAppPath, $szDrive, $szDir, $szFName, $szExt)
	If IsArray($aPathRet) Then
		Return $aPathRet[3]
	Else
		Return ""
	EndIf
EndFunc	;_GetAppName

Func RunAsAdminStub($sAppPath, $sParameters, $bHash, $sAdminPwd, $sLUAPwd)
	Local $bLUAPwd = _Crypt_EncryptData($sLUAPwd, $vCryptKey, $CALG_3DES)
	If @error Then
		Return SetError(1, False, "Ошибка шифрования!") ; Failed to encrypt data!
	Else
		If FileExists($sAppPath) Then
			Local $bAppHash = _Crypt_HashFile($sAppPath, $CALG_SHA1)
			If @error Then
				Return SetError(1, False, "Ошибка создания контрольной суммы!") ; Failed to create hash checksums!
			Else
				If $bHash == $bAppHash Then
					Local $CmdParam = ' -RUNasLUA "' & $sAppPath & '" "' & $sParameters & '" "' & @UserName & '" "' & $bLUAPwd & '"'
					If @compiled == 1 Then
						RunAs($sLocalAdminName, @ComputerName, $sAdminPwd, 0, @AutoItExe & ' ' & $CmdParam, @ScriptDir)
					Else
						RunAs($sLocalAdminName, @ComputerName, $sAdminPwd, 0, @AutoItExe & ' "' & @ScriptFullPath & '"' & $CmdParam, @ScriptDir)
					EndIf
					If @error Then
						Return SetError(1, False, "RunAs - " & _WinAPI_GetLastErrorMessage())
					Else
						_Exit()
						; Return True
					EndIf
				Else
					Return SetError(1, False, "Контрольные суммы не совпадают!") ; Hash checksums not equal!
				EndIf
			EndIf
		Else
			Return SetError(1, False, "Файл не найден!") ; File not found!
		EndIf
	EndIf
EndFunc	;RunAsAdminStub

Func RUNasLUA($sAppPath, $sParameters, $sLUALogin, $bLUAPwd)
	Local $sLUAPwd = _Crypt_DecryptData($bLUAPwd, $vCryptKey, $CALG_3DES)
	If @error Then
		Return SetError(1, False, "Ошибка дешифровки.") ; Failed to decrypt data.
	Else
		$sLUAPwd = BinaryToString($sLUAPwd)
		If FileExists($sAppPath) And IsAdmin() Then
			If _IsAdmin($sLUALogin) == True Then
				MsgBox(0, "Ошибка", "Указанный пользователь находится в локальной группе администраторов!") ; Specified account is member of local administrator group!
			Else
				Local $sWorkingDir = _StripAppName($sAppPath)
				Local $aRet = _MakeMeAdmin_RunAs($sLUALogin, $sLUAPwd, 0, $sAppPath & " " & $sParameters, $sWorkingDir)
				If @error Then MsgBox(0, "Ошибка _MakeMeAdmin_RunAs-" & @extended, $aRet)
			EndIf
		ElseIf Not FileExists($sAppPath) Then
			MsgBox(0, "Ошибка", "Файл не найден!") ; File not found!
		ElseIf Not IsAdmin() Then
			MsgBox(0, "Ошибка", "Требуются права Администратора!") ; Admin rights are required!
		EndIf
	EndIf
	_Exit()
EndFunc	;RUNasLUA

Func _MakeMeAdmin_RunAs($sUserName, $sPassword, $nLogonFlag, $sExecPath, $sWorkingDir)
	Local $sErrMsg, $nErrExtendedMacro = Int(0)
	_NetLocalGroupAddMember($sUserName, $sLocalAdminGroupName)
	If Not @error Then
		Sleep(100)
		Local $PID = RunAs($sUserName, @ComputerName, $sPassword, $nLogonFlag, $sExecPath, $sWorkingDir)
		If @error And $PID == 0 Then
			$sErrMsg &= _WinAPI_GetLastErrorMessage()
			$nErrExtendedMacro = Int(2)
		EndIf
		Sleep(50)
		_NetLocalGroupDelMembers($sUserName, $sLocalAdminGroupName)
		If @error Then
			$sErrMsg &= _WinAPI_GetLastErrorMessage()
			$nErrExtendedMacro = Int(3)
		EndIf
	Else
		$sErrMsg &= _WinAPI_GetLastErrorMessage()
		$nErrExtendedMacro = Int(1)
	EndIf
	If $nErrExtendedMacro == 0 Then
		Return $PID
	Else
		Return SetError(1, $nErrExtendedMacro, $sErrMsg)
	EndIf
EndFunc	;_MakeMeAdmin_RunAs

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

Func _NetUserEnum($sServer = "") ; array[0] contains number of elements
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
	Return $aUserEnum
EndFunc	;_NetUserEnum

Func _NetUserGetLocalGroups($sUsername, $sServer = "") ; array[0] contains number of elements
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
	Return $aLocalGroupNames
EndFunc	;_NetUserGetLocalGroups
