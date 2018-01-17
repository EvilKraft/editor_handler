if WScript.Arguments.Count = 0 then
	Set ObjShell = CreateObject("Shell.Application")
	ObjShell.ShellExecute "wscript.exe", """" & WScript.ScriptFullName & """" & " RunAsAdministrator", , "runas", 1
	WScript.Quit
end if

' from http://dwarf1711.blogspot.com/2007/10/vbscript-urldecode-function.html
Function URLDecode(ByVal str)
	Dim intI, strChar, strRes
	str = Replace(str, "+", " ")
	For intI = 1 To Len(str)
		strChar = Mid(str, intI, 1)
		If strChar = "%" Then
			If intI + 2 < Len(str) Then
				strRes = strRes & Chr(CLng("&H" & Mid(str, intI+1, 2)))
				intI = intI + 2
			End If
		Else
			strRes = strRes & strChar
		End If
	Next
	URLDecode = strRes
End Function

' File Browser via HTA
' Author:   Rudi Degrande, modifications by Denis St-Pierre and Rob van der Woude
' Features: Works in Windows Vista and up (Should also work in XP).
'           Fairly fast.
'           All native code/controls (No 3rd party DLL/ XP DLL).
' Caveats:  Cannot define default starting folder.
'           Uses last folder used with MSHTA.EXE stored in Binary in [HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\ComDlg32].
'           Dialog title says "Choose file to upload".
' Source:   http://social.technet.microsoft.com/Forums/scriptcenter/en-US/a3b358e8-15&alig;-4ba3-bca5-ec349df65ef6
Function SelectFile( )
    Dim objExec, strMSHTA, wshShell

    SelectFile = ""

    ' For use in HTAs as well as "plain" VBScript:
    strMSHTA = "mshta.exe ""about:" & "<" & "input type=file id=FILE>" _
             & "<" & "script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
             & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);" & "<" & "/script>"""
    ' For use in "plain" VBScript only:
    ' strMSHTA = "mshta.exe ""about:<input type=file id=FILE>" _
    '          & "<script>FILE.click();new ActiveXObject('Scripting.FileSystemObject')" _
    '          & ".GetStandardStream(1).WriteLine(FILE.value);close();resizeTo(0,0);</script>"""

    Set wshShell = CreateObject( "WScript.Shell" )
    Set objExec = wshShell.Exec( strMSHTA )

    SelectFile = objExec.StdOut.ReadLine( )

    Set objExec  = Nothing
    Set wshShell = Nothing
End Function
                              

if WScript.Arguments(0) = "RunAsAdministrator" then
	if MsgBox("Do you want do install editor: url scheme handler?",4+vbSystemModal,"editor_handler setup") = 6 then
  
		inputBoxText = "1 - PHPStorm"                       & chr(13) & _
					   "2 - Sublime or Eclipse or NetBeans" & chr(13) & _
					   "3 - Notepad++"                      & chr(13) & _
					   "4 - PHPEd"                          & chr(13) & _
					   "5 - SciTE"                          & chr(13) & _
					   "6 - EmEditor"                       & chr(13) & _
					   "7 - PSPad Editor"                   & chr(13) & _
					   "8 - gVim"                           & chr(13) & _
					   "0 - other text editor (no automatic line number highlighting)"
  
		editor_no = InputBox(inputBoxText, "Select your text editor", 1)
		
		Select Case editor_no
			Case 1 editor = "PHPStorm"
			Case 2 editor = "Sublime or Eclipse or NetBeans"      
			Case 3 editor = "Notepad++"
			Case 4 editor = "PHPEd"
			Case 5 editor = "SciTE"
			Case 6 editor = "EmEditor"
			Case 7 editor = "PSPad Editor"
			Case 8 editor = "gVim"
			Case Else 
				editor = "text editor" 
				editor_no = 0
		End Select

		temp = MsgBox ("Click on OK to open the file selection window and pick the "& editor &" EXE file (problably somewhere in your Program Files directory).",0+vbSystemModal,"Next step")
		filename = SelectFile()
    
		if filename = "" then
			temp = MsgBox ("Aborted installation",0+vbSystemModal,"Aborted")    
			WScript.Quit
		end if

		Set objFSO=CreateObject("Scripting.FileSystemObject")
		temp = CreateObject("WScript.Shell").ExpandEnvironmentStrings("%Temp%")

		outFile="tmp.reg"
		Set objFile = objFSO.CreateTextFile(temp&"\"&outFile,True)
		objFile.Write "Windows Registry Editor Version 5.00" & vbCrLf
		objFile.Write "" & vbCrLf
		objFile.Write "[HKEY_CLASSES_ROOT\editor]" & vbCrLf
		objFile.Write "@="&chr(34)&"URL:editor Protocol"&chr(34) & vbCrLf
		objFile.Write ""&chr(34)&"URL Protocol"&chr(34)&"="&chr(34)&chr(34) & vbCrLf
		objFile.Write "" & vbCrLf
		objFile.Write "[HKEY_CLASSES_ROOT\editor\shell]" & vbCrLf
		objFile.Write "" & vbCrLf
		objFile.Write "[HKEY_CLASSES_ROOT\editor\shell\open]" & vbCrLf
		objFile.Write "" & vbCrLf
		objFile.Write "[HKEY_CLASSES_ROOT\editor\shell\open\command]" & vbCrLf
		objFile.Write "@="&chr(34)&"\"&chr(34)&"wscript.exe"&"\"&chr(34)&" "&"\"&chr(34)&Replace(WScript.ScriptFullName,"\","\\")&"\"&chr(34)&" "& editor_no &" \"&chr(34)&Replace(filename,"\","\\")&"\"&chr(34)&" %1"&chr(34)&"" & vbCrLf
		objFile.Close

		Set ObjShell = CreateObject("Shell.Application")
		ObjShell.ShellExecute "regedit.exe", "/S """ & temp & "\" & outFile & """" & " RunAsAdministrator", , "runas", 1

		temp = MsgBox ("Successfully installed editor: url scheme handler", 0, "Success")
		WScript.Quit   

	end if
else
	str = URLDecode(WScript.Arguments(2))
	Set re = New RegExp
	re.Pattern = "editor://open/\?file=(.+)&line=([0-9]+)"
	re.IgnoreCase = True
	re.Global = False
	Set matches = re.Execute(str)
	
	If matches.Count > 0 Then
		Set match = matches(0)
		If match.SubMatches.Count > 0 Then
		    Set ObjShell = CreateObject("Shell.Application")
		  
			file = chr(34) & match.SubMatches(0) & chr(34)
			line = match.SubMatches(1)
		  
			params = ""
			Select Case WScript.Arguments(0)
				Case 1    params = " --line " & line & " " & file
				Case 2    params = file & ":"   & line
				Case 3    params = file & " -n" & line
				Case 4    params = file & " --line=" & line
				Case 5    params = chr(34) & "-open:" & match.SubMatches(0) & chr(34) & " -goto:" & line
				Case 6    params = file & " /l " & line
				Case 7    params = " -" & line & file
				Case 8    params = file & " +" & line
				Case Else params = file
			End Select
			
			ObjShell.ShellExecute Wscript.Arguments(1), params, , "open", 1
		End If
	End If
end if