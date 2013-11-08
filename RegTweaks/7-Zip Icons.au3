#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.0.0
 Author:		 Serj

 Script Function:
	7-Zip Icons.

#ce ----------------------------------------------------------------------------

Local $nTimerInit = TimerInit()
Local $j = 0, $i = 0
While -1
	Local $sRegEnum = RegEnumKey("HKEY_CLASSES_ROOT\", $i + 1)
	If @error <> 0 Then ExitLoop
	$i += 1
	If StringInStr($sRegEnum, "7-Zip.") <> 0 Then
		If RegWrite("HKEY_CLASSES_ROOT\" & $sRegEnum & "\DefaultIcon", "", "REG_EXPAND_SZ", "%SystemRoot%\system32\cabview.dll,-1") == 1 Then $j +=1	
	EndIf
WEnd
Local $nTimerDiff = Round(TimerDiff($nTimerInit) / 1000, 6)
MsgBox(4096, "Info:", "Processed: " & $i & @CR & "Time elapsed: " & $nTimerDiff & " sec" & @CR & "Changed: " & $j)
Exit
