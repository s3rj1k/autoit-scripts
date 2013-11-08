#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.0.0
 Author:		 Serj

 Script Function:
	Clean OpenWith MRU.

#ce ----------------------------------------------------------------------------

Local $nLogFile = FileOpen(@ScriptName & ".log", 1)
If $nLogFile = -1 Then
	MsgBox(0, "Error", "Unable to open logfile.")
	Exit
EndIf
FileWriteLine($nLogFile, "[" & @MDAY & "/" & @MON & "/" & @YEAR & "-" & @HOUR & ":" & @MIN & ":" & @SEC & ":" & @MSEC & "]" & @CRLF)
Local $nTimerInit = TimerInit()
Local $j = 0, $i = 0
While -1
	Local $sRegEnum = RegEnumKey("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\", $i + 1)
	If @error <> 0 Then ExitLoop
	$i += 1
	If RegDelete("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & $sRegEnum & "\OpenWithList") Then
		$j +=1
		FileWriteLine($nLogFile, "Deleted: " & "HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\" & $sRegEnum & "\OpenWithList" & @CRLF)
	EndIf
WEnd
Local $nTimerDiff = Round(TimerDiff($nTimerInit) / 1000, 6)
MsgBox(4096, "Info:", "Subkeys processed: " & $i & @CR & "Time elapsed: " & $nTimerDiff & " sec" & @CR & "Keys deleted: " & $j)
FileWriteLine($nLogFile, "Subkeys processed: " & $i & @CRLF)
FileWriteLine($nLogFile, "Time elapsed: " & $nTimerDiff & " sec" & @CRLF)
FileWriteLine($nLogFile, "Keys deleted: " & $j & @CRLF)
FileClose($nLogFile)
Exit