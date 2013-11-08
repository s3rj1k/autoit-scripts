#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.0.0
 Author:		 Serj

 Script Function:
	Remove print/printto, edit Context Menu Handlers.

#ce ----------------------------------------------------------------------------
#RequireAdmin

Local $nLogFile = FileOpen(@ScriptName & ".log", 1)
If $nLogFile = -1 Then
	MsgBox(0, "Error", "Unable to open logfile.")
	Exit
EndIf
FileWriteLine($nLogFile, "[" & @MDAY & "/" & @MON & "/" & @YEAR & "-" & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & "]" & @CRLF)
Local $nTimerInit = TimerInit()
Local $j = 0, $i1 = 0, $i2 = 0
While -1
	Local $sRegEnum = RegEnumKey("HKEY_CLASSES_ROOT\", $i1 + 1)
	If @error <> 0 Then ExitLoop
	$i1 += 1
	If RegDelete("HKEY_CLASSES_ROOT\" & $sRegEnum & "\shell\print") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\" & $sRegEnum & "\shell\print" & @CRLF)
	EndIf
	If RegDelete("HKEY_CLASSES_ROOT\" & $sRegEnum & "\shell\printto") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\" & $sRegEnum & "\shell\printto" & @CRLF)
	EndIf
	If RegDelete("HKEY_CLASSES_ROOT\" & $sRegEnum & "\shell\edit") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\" & $sRegEnum & "\shell\edit" & @CRLF)
	EndIf
WEnd
While -1
	$sRegEnum = RegEnumKey("HKEY_CLASSES_ROOT\SystemFileAssociations", $i2 + 1)
	If @error <> 0 Then ExitLoop
	$i2 += 1
	If RegDelete("HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\print") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\print" & @CRLF)
	EndIf
	If RegDelete("HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\printto") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\printto" & @CRLF)
	EndIf
	If RegDelete("HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\edit") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\edit" & @CRLF)
	EndIf
	If RegDelete("HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\MediaInfo") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CLASSES_ROOT\SystemFileAssociations\" & $sRegEnum & "\shell\MediaInfo" & @CRLF)
	EndIf
WEnd
Local $nTimerDiff = Round(TimerDiff($nTimerInit) / 1000, 6)
MsgBox(4096, "Info:", "Subkeys processed: " & $i1 + $i2 & @CR & "Time elapsed: " & $nTimerDiff & " sec" & @CR & "Keys deleted: " & $j)
FileWriteLine($nLogFile, "Subkeys processed: " & $i1 + $i2 & @CRLF)
FileWriteLine($nLogFile, "Time elapsed: " & $nTimerDiff & " sec" & @CRLF)
FileWriteLine($nLogFile, "Keys deleted: " & $j & @CRLF)
FileClose($nLogFile)
Exit
