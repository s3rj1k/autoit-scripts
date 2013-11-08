#cs -------------------------------------------------------------------------------

 AutoIt Version: v3.3.6.1
 Author:		 Serj

 Script Function:
	EncryptedRunAs.

#ce -------------------------------------------------------------------------------

#NoTrayIcon
; #RequireAdmin
#include <Array.au3>
#include <ComboConstants.au3>
#include <Constants.au3>
#Include <Crypt.au3>
#include <EditConstants.au3>
#include <File.au3>
#include <GUIConstantsEx.au3>
#include <WinAPI.au3>
#include <WindowsConstants.au3>

Opt("MustDeclareVars", 1)
Opt("TrayMenuMode", 1)

Local $aUserEnum = _NetUserEnum()
If @error Then
	MsgBox(0, "Error", _WinAPI_GetLastErrorMessage(), 30)
	Exit
EndIf
Local $sUserEnum = _ArrayToString($aUserEnum, "|", 1, $aUserEnum[0]+1)
Global $hKey = DriveGetSerial(@HomeDrive)
If @error Then
	MsgBox(0, "Error", "Crypt key error!", 30)
	Exit
EndIf
Global CONST $EncryptedStringSeparator = "¶"

If $CmdLine[0] == 1 Then
	EncryptedRunAs($CmdLine[1])
ElseIf $CmdLine[0] == 0 Then
	GUI($sUserEnum)
Else
	Exit
EndIf

Func GUI($sUserEnum)
	Local $hWnd = GUICreate("EncryptedRunAs Options:", 420, 420, -1, -1, BitOr($WS_CAPTION, $WS_SYSMENU, $WS_MINIMIZEBOX))
		GUICtrlCreateGroup("Application", 10, 5, 400, 137)
			GUICtrlCreateLabel("Path:", 20, 25, 80, 20)
			Local $AppPath_Input = GUICtrlCreateInput("", 90, 22, 284, 20, BitOr($ES_AUTOHSCROLL, $ES_READONLY))
			Local $Browse_Button = GUICtrlCreateButton("...", 378, 21, 22, 22)
			GUICtrlCreateLabel("Parameters:", 20, 55, 80, 20)
			Local $AppParameters_Input = GUICtrlCreateInput("", 90, 52, 284, 20, $ES_AUTOHSCROLL)
			Local $CMDHelp_Button = GUICtrlCreateButton("?", 378, 51, 22, 22)
			GUICtrlCreateLabel("Execute in:", 20, 85, 80, 20)
			Local $WorkingDir_Combo = GUICtrlCreateCombo("", 90, 82, 284, 100, BitOr($CBS_AUTOHSCROLL, $GUI_SS_DEFAULT_COMBO, $CBS_SORT, $CBS_DISABLENOSCROLL))
			GUICtrlSetData($WorkingDir_Combo, StringTrimRight(@UserProfileDir, StringLen(@UserName) + 1) & "|" & @UserProfileDir & "|" & @TempDir & "|" & @ProgramFilesDir & "|" & @CommonFilesDir & "|" & @WindowsDir & "|" & @SystemDir)
			Local $WorkingDir_Button = GUICtrlCreateButton("»", 378, 81, 22, 22)
			Local $SHA1_CheckBox = GUICtrlCreateCheckbox("SHA1:", 20, 113, 50, 20)
			Local $SHA1_Input = GUICtrlCreateInput("SHA1_Enabled", 90, 112, 284, 20, BitOr($ES_AUTOHSCROLL, $ES_READONLY))
			GUICtrlSetFont($SHA1_Input, 8, 2)
			GUICtrlSetState($SHA1_CheckBox, $GUI_CHECKED)
			Local $SHA1Refresh_Button = GUICtrlCreateButton("»", 378, 111, 22, 22)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateGroup("Authentication", 10, 145, 400, 109)
			GUICtrlCreateLabel("Login:", 20, 165, 50, 20)
			Local $Login_Combo = GUICtrlCreateCombo("", 90, 162, 284, 100)
			GUICtrlSetData($Login_Combo, $sUserEnum, @UserName)
			Local $LogonUserTest_Button = GUICtrlCreateButton("»", 378, 161, 22, 52)
			GUICtrlCreateLabel("Password:", 20, 195, 50, 20)
			Local $Password_Input = GUICtrlCreateInput("", 90, 192, 284, 20, BitOr($ES_PASSWORD, $ES_AUTOHSCROLL))
			GUICtrlCreateLabel("Logon:", 20, 225, 50, 20)
			Local $LogonFlag_Combo = GUICtrlCreateCombo("", 90, 222, 284, 100, $CBS_DROPDOWNLIST)
			GUICtrlSetData($LogonFlag_Combo, "Interactive logon with no profile.|Interactive logon with profile.", "Interactive logon with profile.")
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		GUICtrlCreateGroup("Output", 10, 257, 400, 119)
			Local $Output_Edit = GUICtrlCreateEdit("Please input all data before encryption.", 20, 272, 380, 72, BitOr($ES_AUTOVSCROLL, $WS_VSCROLL, $ES_READONLY))
			Local $Shortcut_CheckBox = GUICtrlCreateCheckbox("Create shortcut to generated Crypt.key file", 20, 350, 340, 20)
			GUICtrlSetFont($Output_Edit, 9, 2)
			GUICtrlSetState($Shortcut_CheckBox, $GUI_CHECKED)
		GUICtrlCreateGroup("", -99, -99, 1, 1)
		Local $Open_Button = GUICtrlCreateButton("Open", 10, 384, 85, 27)
		Local $QuickSave_Button = GUICtrlCreateButton("Save", 105, 384, 85, 27)
		Local $Save_Button = GUICtrlCreateButton("...", 190, 384, 30, 27)
		Local $Test_Button = GUICtrlCreateButton("Test", 230, 384, 85, 27)
		Local $Quit_Button = GUICtrlCreateButton("Quit", 325, 384, 85, 27)
		GUICtrlSetState($QuickSave_Button, $GUI_DEFBUTTON)
	GUISetState(@SW_SHOW)
	While 1
		Local $msg = GUIGetMsg(), $AppPath, $AppParameters, $WorkingDir, $UserName, $Password, $Logon_Combo_Flag, $SHA1_Flag, $AppSHA1, $CreateShortcut
		If $msg == $Browse_Button Then
			$SHA1_Flag = GUICtrlRead($SHA1_CheckBox)
			Local $BrowsePath = FileOpenDialog("Choose application to RunAs...", @WorkingDir & "\", "Programs (*.exe;*.com;*.cmd;*.bat)", 1 + 2, "", $hWnd)
			If @error Then
				GUICtrlSetData($Output_Edit, "No application was chosen!" & @CRLF)
			Else
				GUICtrlSetData($AppPath_Input, $BrowsePath)
				$WorkingDir = _StripAppName($BrowsePath)
				GUICtrlSetData($WorkingDir_Combo, $WorkingDir, $WorkingDir)
				If $SHA1_Flag == 1 Then
					GUICtrlSetData($Output_Edit, "Calculating SHA1 hash..." & @CRLF)
					Local $aRetAppSHA1 = SHA1Wrapper($BrowsePath)
					GUICtrlSetData($SHA1_Input, $aRetAppSHA1[0])
					GUICtrlSetData($Output_Edit, GUICtrlRead($Output_Edit) & $aRetAppSHA1[1])
				Else
					GUICtrlSetData($SHA1_Input, "SHA1_Disabled")
					GUICtrlSetData($Output_Edit, "SHA1 hash disabled." & @CRLF)
				EndIf
			EndIf
		EndIf
		If $msg == $CMDHelp_Button Then
			$AppPath = GUICtrlRead($AppPath_Input)
			If FileExists($AppPath) == 1 Then Run(@ComSpec & " /k", @WorkingDir)
		EndIf
		If $msg == $WorkingDir_Button Then
			$AppPath = GUICtrlRead($AppPath_Input)
			$WorkingDir = _StripAppName($AppPath)
			GUICtrlSetData($WorkingDir_Combo, $WorkingDir, $WorkingDir)
		EndIf
		If $msg == $SHA1_CheckBox Then
			$AppPath = GUICtrlRead($AppPath_Input)
			$SHA1_Flag = GUICtrlRead($SHA1_CheckBox)
			If $SHA1_Flag == 1 Then
				GUICtrlSetData($SHA1_Input, "SHA1_Enabled")
				GUICtrlSetData($Output_Edit, "SHA1 hash enabled." & @CRLF)
				If FileExists($AppPath) == 1 Then
					GUICtrlSetData($Output_Edit, "Calculating SHA1 hash..." & @CRLF)
					Local $aRetAppSHA1 = SHA1Wrapper($AppPath)
					GUICtrlSetData($SHA1_Input, $aRetAppSHA1[0])
					GUICtrlSetData($Output_Edit, GUICtrlRead($Output_Edit) & $aRetAppSHA1[1])
				EndIf
			Else
				GUICtrlSetData($SHA1_Input, "SHA1_Disabled")
				GUICtrlSetData($Output_Edit, "SHA1 hash disabled." & @CRLF)
			EndIf
		EndIf
		If $msg == $SHA1Refresh_Button Then
			$AppPath = GUICtrlRead($AppPath_Input)
			$SHA1_Flag = GUICtrlRead($SHA1_CheckBox)
			If FileExists($AppPath) == 1 And $SHA1_Flag == 1 Then
				GUICtrlSetData($Output_Edit, "Calculating SHA1 hash..." & @CRLF)
				Local $aRetAppSHA1 = SHA1Wrapper($AppPath)
				GUICtrlSetData($SHA1_Input, $aRetAppSHA1[0])
				GUICtrlSetData($Output_Edit, GUICtrlRead($Output_Edit) & $aRetAppSHA1[1])
			ElseIf $SHA1_Flag == 0 Then
				GUICtrlSetData($SHA1_Input, "SHA1_Disabled")
				GUICtrlSetData($Output_Edit, "SHA1 hash disabled." & @CRLF)
			ElseIf $SHA1_Flag == 1 And FileExists($AppPath) == 0 Then
				GUICtrlSetData($SHA1_Input, "SHA1_Enabled")
				GUICtrlSetData($Output_Edit, "SHA1 hash enabled." & @CRLF)
			EndIf
		EndIf
		If $msg == $LogonUserTest_Button Then
			$UserName = GUICtrlRead($Login_Combo)
			$Password = GUICtrlRead($Password_Input)
			Local $LogonUser = _LogonUser($UserName, $Password)
			If @error Then
				GUICtrlSetData($Output_Edit, "LogonUser WinAPI function error." & @CRLF & "Authentication check is impossible." & @CRLF)
			ElseIf $LogonUser == True Then
				GUICtrlSetData($Output_Edit, "LogonUser authentication successful." & @CRLF)
			ElseIf $LogonUser == False Then
				GUICtrlSetData($Output_Edit, "LogonUser authentication failed." & @CRLF)
			EndIf
		EndIf
		If $msg == $Open_Button Then
			Local $OpenCryptFilePath = FileOpenDialog("Open Crypt.key file for EncryptedRunAs...", @WorkingDir & "\", "Crypt file (*.key)", 1 + 2, "", $hWnd)
			If @error Then
				GUICtrlSetData($Output_Edit, "No Crypt.key file was chosen!" & @CRLF)
			Else
				EncryptedRunAs($OpenCryptFilePath)
			EndIf
		EndIf
		If $msg == $Save_Button Or $msg == $QuickSave_Button Or $msg == $Test_Button Then
			$AppPath = GUICtrlRead($AppPath_Input)
			$AppParameters = GUICtrlRead($AppParameters_Input)
			$WorkingDir = GUICtrlRead($WorkingDir_Combo)
			$UserName = GUICtrlRead($Login_Combo)
			$Password = GUICtrlRead($Password_Input)
			$Logon_Combo_Flag = GUICtrlRead($LogonFlag_Combo)
			$SHA1_Flag = GUICtrlRead($SHA1_CheckBox)
			$AppSHA1 = GUICtrlRead($SHA1_Input)
			$CreateShortcut = GUICtrlRead($Shortcut_CheckBox)
			If $Logon_Combo_Flag == "Interactive logon with no profile." Then
				Local $Logon_Flag = 0
			ElseIf $Logon_Combo_Flag == "Interactive logon with profile." Then
				Local $Logon_Flag = 1
			EndIf
			If FileExists($AppPath) == 1 And $AppSHA1 <> "SHA1_Failed" And FileExists($WorkingDir) == 1 Then
				Local $EncryptedString = _Crypt_EncryptData($AppPath & $EncryptedStringSeparator & $AppParameters & $EncryptedStringSeparator & $WorkingDir & $EncryptedStringSeparator & $UserName & $EncryptedStringSeparator & $Password & $EncryptedStringSeparator & $Logon_Flag & $EncryptedStringSeparator & $AppSHA1, $hKey, $CALG_3DES)
				If $msg == $Save_Button Then
					Local $CryptFileSavePath = FileSaveDialog("Choose name and location for crypt file...", $AppPath, "Encrypted key (*.key)", 2 + 16, "crypt.key", $hWnd)
					If @error <> 1 Then
						Local $SaveCryptKeyFileRet = SaveCryptKeyFile($AppPath, $CryptFileSavePath, $EncryptedString, $CreateShortcut, $hWnd)
						GUICtrlSetData($Output_Edit, $SaveCryptKeyFileRet)
					Else
						GUICtrlSetData($Output_Edit, "Crypt file save cancelled." & @CRLF)
					EndIf
				EndIf
				If $msg == $QuickSave_Button Then
					Local $SaveCryptKeyFileRet = SaveCryptKeyFile($AppPath, 0, $EncryptedString, $CreateShortcut, $hWnd)
					GUICtrlSetData($Output_Edit, $SaveCryptKeyFileRet)
				EndIf
				If $msg == $Test_Button Then
					$CreateShortcut = 0
					Local $SaveCryptKeyFileRet = SaveCryptKeyFile($AppPath, @TempDir & '\crypt.key', $EncryptedString, $CreateShortcut, $hWnd)
					GUICtrlSetData($Output_Edit, $SaveCryptKeyFileRet)
					Local $EncryptedRunAsRet = EncryptedRunAs(@TempDir & '\crypt.key')
					If $EncryptedRunAsRet == True Then
						GUICtrlSetData($Output_Edit, GUICtrlRead($Output_Edit) & "Test-run successful." & @CRLF)
					Else
						GUICtrlSetData($Output_Edit, GUICtrlRead($Output_Edit) & "Test-run failed." & @CRLF)
					EndIf
					FileDelete(@TempDir & '\crypt.key')
					WinActivate($hWnd)
				EndIf
			Else
				Local $errmsg = ""
				If FileExists($AppPath) == 0 Then $errmsg &= "Specified file does not exist." & @CRLF
				If FileExists($WorkingDir) == 0 Then $errmsg &= "Specified working directory does not exist." & @CRLF
				If $AppSHA1 == "SHA1_Failed" Then $errmsg &= "SHA1 hash computation failed." & @CRLF & "Please disable SHA1 hash computation!" & @CRLF
				GUICtrlSetData($Output_Edit, $errmsg)
			EndIf
		EndIf
		If $msg == $Quit_Button Or $msg == $GUI_EVENT_CLOSE Then Exit
	WEnd
	GUIDelete()
EndFunc	;GUI

Func SHA1Wrapper($AppPath)
	Local $hTimer = TimerInit()
	Local $AppSHA1 = _Crypt_HashFile($AppPath, $CALG_SHA1)
	Local $aRet[2]
	If @error Then
		$aRet[0] = "SHA1_Failed"
		$aRet[1] = "SHA1 hash computation failed." & @CRLF
		SetError(1)
		Return $aRet
	Else
		Local $iTimer = TimerDiff($hTimer)
		$aRet[0] = $AppSHA1
		$aRet[1] = "SHA1 hash computation completed successfully." & @CRLF & "SHA1 took " & $iTimer & " ms" & @CRLF
		SetError(0)
		Return $aRet
	EndIf
EndFunc	;SHA1Wrapper

Func _StripAppName($AppPath)
	Dim $szDrive, $szDir, $szFName, $szExt
	Local $aPathRet = _PathSplit($AppPath, $szDrive, $szDir, $szFName, $szExt)
	If FileExists($aPathRet[2]) == 0 Then
		Return @TempDir
	Else
		Return $aPathRet[1] & $aPathRet[2]
	EndIf
EndFunc	;_StripAppName

Func _GetAppName($AppPath)
	Dim $szDrive, $szDir, $szFName, $szExt
	Local $aPathRet = _PathSplit($AppPath, $szDrive, $szDir, $szFName, $szExt)
	If IsArray($aPathRet) Then
		Return $aPathRet[3]
	Else
		Return "AppName"
	EndIf
EndFunc	;_GetAppName

Func _GetAppExtension($AppPath)
	Dim $szDrive, $szDir, $szFName, $szExt
	Local $aPathRet = _PathSplit($AppPath, $szDrive, $szDir, $szFName, $szExt)
	If IsArray($aPathRet) Then
		Return $aPathRet[4]
	Else
		Return ""
	EndIf
EndFunc	;_GetAppExtension

Func SaveCryptKeyFile($AppPath, $CryptFileSavePath, $EncryptedString, $CreateShortcut, $hWnd)
	Local $sRet
	If $CryptFileSavePath == 0 Then
		$CryptFileSavePath = @UserProfileDir & "\" & _GetAppName($AppPath) & '.crypt.key'
	Endif
	Local $CryptFileOpenTest = FileOpen($CryptFileSavePath, 16 + 2)
	If $CryptFileOpenTest = -1 Then
		$sRet &= "Unable to create crypt file." & @CRLF
	Else
		FileWrite($CryptFileOpenTest, $EncryptedString)
		If @error Then
			$sRet &= "File cannot be written to." & @CRLF
		EndIf
		FileClose($CryptFileOpenTest)
		If $CreateShortcut == 1 Then
			Local $AppExtension = _GetAppExtension($AppPath)
			If $AppExtension == ".exe" Then
				Local $IconPath = $AppPath
				Local $IconIndex = 0
			Else
				Local $IconPath = @SystemDir & "\shell32.dll"
				Local $IconIndex = 2
			EndIf
			Local $ShortcutSavePath = FileSaveDialog("Choose name and location for shortcut...", @DesktopDir, "Shortcut (*.lnk)", 2 + 16, _GetAppName($AppPath) & ".lnk", $hWnd)
			If @error Then
				$sRet &= "Shortcut creation canceled." & @CRLF
			Else
				If @Compiled == 0 Then
					Local $ShortcutArgs = '"' & @ScriptFullPath & '" "' & $CryptFileSavePath & '"'
				ElseIf @Compiled == 1 Then
					Local $ShortcutArgs = '"' & $CryptFileSavePath & '"'
				EndIf
				FileCreateShortcut(@AutoItExe, $ShortcutSavePath, @ScriptDir, $ShortcutArgs, "", $IconPath, "", $IconIndex, @SW_SHOWNORMAL)
				If @error Then
					$sRet &= "Shortcut creation failed." & @CRLF
				EndIf
				If FileExists($ShortcutSavePath) == 1 Then $sRet &= "Shortcut created successfully." & @CRLF
			EndIf
		EndIf
	EndIf
	If FileExists($CryptFileSavePath) == 1 Then $sRet &= "Crypt file created successfully." & @CRLF
	Return $sRet
EndFunc	;SaveCryptKeyFile

Func EncryptedRunAs($CryptFilePath)
	Local $CryptFileOpenTest = FileOpen($CryptFilePath, 16)
	If $CryptFileOpenTest = -1 Then
		MsgBox(0, "Error", "Unable to open crypt file.", 30)
		Return False
		Exit
	Else
		Local $EncryptedString = FileRead($CryptFileOpenTest)
		If @error Then
			MsgBox(0, "Error", "File cannot be opened in read mode.", 30)
			Return False
			Exit
		EndIf
		FileClose($CryptFileOpenTest)
		Local $DecryptedString = BinaryToString(_Crypt_DecryptData($EncryptedString, $hKey, $CALG_3DES))
		Local $DecryptedArray = StringSplit($DecryptedString, $EncryptedStringSeparator, 1)
		If $DecryptedArray[0] == 7 Then
			Local $AppPath = $DecryptedArray[1]
			Local $AppParameters = $DecryptedArray[2]
			Local $WorkingDir = $DecryptedArray[3]
			Local $UserName = $DecryptedArray[4]
			Local $Password = $DecryptedArray[5]
			Local $Logon_Flag = $DecryptedArray[6]
			Local $AppSHA1Str = $DecryptedArray[7]
			If $AppParameters == "" Then
				Local $AppPathANDParameters = $AppPath
			Else
				Local $AppPathANDParameters = $AppPath & " " & $AppParameters
			EndIf
			If $AppSHA1Str == "SHA1_Disabled" Then
				RunAs($UserName, @ComputerName, $Password, $Logon_Flag, $AppPathANDParameters, $WorkingDir)
				If @error Then
					MsgBox(0, "Error", "RunAs failed:" & @CRLF & @TAB & _WinAPI_GetLastErrorMessage(), 30)
					Return False
					Exit
				EndIf
			Else
				Local $AppSHA1Gen = _Crypt_HashFile($AppPath, $CALG_SHA1)
				If @error Then
					MsgBox(0, "Error", "SHA1 hash computation failed.", 30)
					Return False
					Exit
				EndIf
				If $AppSHA1Str == $AppSHA1Gen Then
					RunAs($UserName, @ComputerName, $Password, $Logon_Flag, $AppPathANDParameters, $WorkingDir)
					If @error Then
						MsgBox(0, "Error", "RunAs failed:" & @CRLF & @TAB & _WinAPI_GetLastErrorMessage(), 30)
						Return False
						Exit
					EndIf
				Else
					MsgBox(0, "Error", "File integrity check failed." & @CRLF & "SHA1 file hash is diffrent from encoded in crypt file.", 30)
					Return False
					Exit
				EndIf
			EndIf
		Else
			MsgBox(0, "Error", "Decryption failed. Error in crypt file integrity.", 30)
			Return False
			Exit
		EndIf
	EndIf
	Return True
EndFunc	;EncryptedRunAs

Func _NetUserEnum($sServer = "")
	Local $tBufPtr = DllStructCreate("ptr")
	Local $tEntriesRead = DllStructCreate("dword")
	Local $tTotalEntries = DllStructCreate("dword")
	Local $aRet = DllCall("Netapi32.dll", "int", "NetUserEnum", "wstr", $sServer, "dword", 1, "dword", 2, "ptr", DllStructGetPtr($tBufPtr), "dword", -1, "ptr", DllStructGetPtr($tEntriesRead), "ptr", DllStructGetPtr($tTotalEntries), "ptr", 0 )
	If $aRet[0] Then Return SetError(1, $aRet[0], False)
	Local Const $UF_ACCOUNTDISABLE = 0x2
	Local $iEntriesRead = DllStructGetData($tEntriesRead,1)
	Local $pBuf = DllStructGetData($tBufPtr,1)
	Local $aUserEnum[1] = [0]
	Local $sUserInfo1 = "ptr;ptr;dword;dword;ptr;ptr;dword;ptr"
	Local $tUserInfo1 = DllStructCreate ($sUserInfo1)
	Local $zUserInfo1 = DllStructGetSize($tUserInfo1)
	For $i=1 To $iEntriesRead
		$tUserInfo1 = DllStructCreate($sUserInfo1, $pBuf+($i-1)*$zUserInfo1)
		Local $tName = DllStructCreate("wchar[256]", DllStructGetData($tUserInfo1,1))
		Local $tFlag = DllStructGetData($tUserInfo1,7)
		If BitAnd($tFlag, $UF_ACCOUNTDISABLE)=0 Then
			$aUserEnum[0] += 1
			ReDim $aUserEnum[$aUserEnum[0]+1]
			$aUserEnum[$aUserEnum[0]] = DllStructGetData($tName,1)
		EndIf
	Next
	DllCall("Netapi32.dll", "int", "NetApiBufferFree", "ptr", $pBuf)
	Return $aUserEnum ; $aUserEnum[0] contains number of elements
EndFunc	;_NetUserEnum

Func _LogonUser($sUsername, $sPassword, $sServer = '.')
	Local $stToken = DllStructCreate("int")
	Local $aRet = DllCall("advapi32.dll", "int", "LogonUser", _
			"str", $sUsername, "str", $sServer, "str", $sPassword, "dword", 3, "dword", 0, "ptr", DllStructGetPtr($stToken))
	; Local $hToken = DllStructGetData($stToken, 1)
	If @error Then
		Local $ErrorMsg = _WinAPI_GetLastErrorMessage()
		Return SetError(1, @error, $ErrorMsg)
	ElseIf $aRet[0] <> 0 Then
		Return True ; Returns True if user exists
	ElseIf $aRet[0] == 0 Then
		Return False
	EndIf
EndFunc	;_LogonUser
