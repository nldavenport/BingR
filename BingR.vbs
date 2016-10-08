' Copyright � 1984-2015 Nathan Davenport. All rights reserved.
' Free for all users, but leave in this header
' http://www.????.com
' ====================================================
' Language:	VBScript
' Name:		xxxx.vbs
' Author:	Nathan Davenport
' Date:		mm/yyyy
'
' Description:	
' Comment:	
' Keywords:
' ====================================================
Option Explicit
'On Error Resume Next

' Variable Declarations
' ----------------------------------------------------
Dim i ' Counter
Dim StartTime 		: StartTime = Now() ' start script timer

Dim KB	: KB = 1024
Dim MB	: MB = KB * 1024
Dim GB	: GB = MB * 1024

Dim blnReqArgs		: blnReqArgs = False ' Does this app require Command line arguments?
Dim blnReqDomain	: blnReqDomain = False ' Does the App require the PC to be on a domain?
Dim blnHelp			: blnHelp = False
Dim blnVerbose		: blnVerbose = False
Dim blnDebug		: blnDebug = False
Dim blnLog			: blnLog = False
Dim blnQuiet		: blnQuiet = False
Dim blnAskRegistry	: blnAskRegistry = False

' Object Declarations
' ----------------------------------------------------
Dim objFSO		: Set objFSO = CreateObject("Scripting.FileSystemObject")
'Dim objNet		: Set objNet = CreateObject("WScript.Network")
Dim objSysInfo	: Set objSysInfo = CreateObject( "WinNTSystemInfo" )
Dim objShell	: Set objShell = CreateObject("WScript.Shell")
Dim objSysEnv	: Set objSysEnv = objShell.Environment
Dim objUsrEnv	: Set objUsrEnv = objShell.Environment("User")
Dim objPrcEnv	: Set objPrcEnv = objShell.Environment("Process")

Dim strComputer	: strComputer = "." ' Default to localhost
Dim strUserName	: strUserName = objShell.ExpandEnvironmentStrings( "%USERNAME%" )

' Constant Declarations
' ----------------------------------------------------
Const TypeBinary = 1 ' Binary data
Const TypeText = 2 ' Default, Text data

Const ForReading = 1 'Open a file for reading only. You can't write to this file.
Const ForWriting = 2 'Open a file for writing.
Const ForAppending = 8 'Open a file and write to the end of the file.

Const OverWriteFiles = True

Const TristateUseDefault = -2 'Opens the file using the system default.
Const TristateTrue = -1 'Opens the file as Unicode.
Const TristateFalse = 0 'Opens the file as ASCII.

' Special folders: objFSO.GetSpecialFolder(Value)
' ----------------------------------------------------
Const WindowsFolder   = 0
Const SystemFolder    = 1
Const TemporaryFolder = 2

' ====================================================
' Main Body
' ====================================================
Dim strPath			: strPath = ".\" ' set path to exe location
Dim strTempPath		: strTempPath = "" ' set temp path to temp folder as default
Dim strScriptName	: strScriptName = Left(WScript.ScriptName,Len(WScript.ScriptName)-4)
Dim strLogFilename, objLogFile
Dim strOutputFilename, objOutputFile

'strComputer = objNet.ComputerName ' Set to PC name - override default
Dim objWMI		: Set objWMI = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Dim intMaxLen	: intMaxLen = 15
Dim IntMinTime	: IntMinTime = 20
Dim intMaxTime	: intMaxTime = 120
Dim fn, un, pw

DominantArguments ("ini")

If blnVerbose Or blnDebug Or blnLog Then Echo "", "Start Time: " & StartTime
If blnVerbose Or blnDebug Or blnLog Then Echo "", OS("cap") & OS("ver") ' Debug Get OS version

If fn = "bing" Then
	SearchBing (30)
Else
	Echo "", RandomWord (15)
End If
	
CleanUp
WScript.Quit

' ====================================================
' * Place App specific Functions & Subroutines Here *
' Use the blnDebug or blnLog code below as needed
' If blnDebug Then Echo "text"
' ====================================================

' RandomCharacter : select a random character from a subset
' ----------------------------------------------------
Function RandomCharacter(chartype)
	Dim Vowels		: Vowels = "aeiouy "
	Dim Consenants	: Consenants = "bcdfghjklmnpqrstvwxyz "
	Dim UpperCase	: UpperCase = Ucase (Consenants) & UCase (Vowels)
	Dim LowerCase	: LowerCase = Lcase (Consenants) & lCase (Vowels)
	Dim MixedCase	: MixedCase = UpperCase & LowerCase
	Dim Numbers		: Numbers = "1234567890 "
	Dim SpecialChars : SpecialChars = "!@#$%^&*()-=_+[]{}\|;:',./<>? "
	Dim AllowedChars : AllowedChars = MixedCase & Numbers & SpecialChars
	
If blnDebug Then Echo "", "Entered RandomhCharacter"
	Randomize
	Select Case chartype
		Case "V" ' Vowels
		RandomCharacter = Mid(Vowels, Int(Rnd() * Len(Vowels)+1),1)		
		Case "C" ' Consenants
		RandomCharacter = Mid(Consenants, Int(Rnd() * Len(Consenants)+1),1)		
		Case "UC" ' Upper Case
		RandomCharacter = Mid(UpperCase, Int(Rnd() * Len(UpperCase)+1),1)		
		Case "LC" ' Lower Case
		RandomCharacter = Mid(LowerCase, Int(Rnd() * Len(LowerCase)+1),1)		
		Case "MC" ' Mixed Case
		RandomCharacter = Mid(MixedCase, Int(Rnd() * Len(MixedCase)+1),1)		
		Case "N" ' Numbers
		RandomCharacter = Mid(Numbers, Int(Rnd() * Len(Numbers)+1),1)		
		Case "SC" ' Special Characters
		RandomCharacter = Mid(SpecialChars, Int(Rnd() * Len(SpecialChars)+1),1)		
		Case "AC" ' Allowed Characters
		RandomCharacter = Mid(AllowedChars, Int(Rnd() * Len(AllowedChars)+1),1)		
	End Select
End Function

' RandomWord : create a random word (set of characters)
' ----------------------------------------------------
Function RandomWord(Length)
	Dim strTempWord
	Dim charType
	
If blnDebug Then Echo "", "Entered RandomWord"
	strTempWord = ""
	
	Do While Len(strTempWord) < Length
		'		test = Len(strTempWord)
		Randomize
		
		charType = Int(Rnd() * 3)
		Select Case charType
			Case 0,2
			strTempWord = strTempWord & RandomCharacter ("C")
			Case 1
			strTempWord = strTempWord & RandomCharacter ("V")
		End Select
	Loop
	RandomWord = strTempWord
End Function

' RandomURL : Random Bing search usng random word
' ----------------------------------------------------
Function RandomURL
	Dim randomLength, randomSearch
	
If blnDebug Then Echo "", "Entered RandomURL"
	Randomize
	randomLength = Int(intMaxLen * Rnd()) + 2
	randomSearch = Int(4 * Rnd())
	
'	SearchWord = RandomWord (randomLength)
	Select Case randomSearch
		Case 0
		RandomURL = "https://www.bing.com/Search?q=" & RandomWord (randomLength) ' & "&FORM=ML10NS&CREA=ML10NS" ' Search Bing
		Case 1
		RandomURL = "https://www.bing.com/images/Search?q=" & RandomWord (randomLength) ' & "&FORM=ML10NS&CREA=ML10NS" ' Search Bing images
		Case 2
		RandomURL = "https://www.bing.com/videos/Search?q=" & RandomWord (randomLength) ' & "&FORM=ML10NS&CREA=ML10NS" ' Seaech Bing videos
		Case 3
		RandomURL = "https://www.bing.com/news/Search?q=" & RandomWord (randomLength) ' & "&FORM=ML10NS&CREA=ML10NS" ' Search Bing news
		Case 4
		RandomURL = "https://www.bing.com/maps/Search?q=" & RandomWord (randomLength) ' & "?FORM=ML10NS&CREA=ML10NS" ' Search Bing maps
	End Select
	Echo "", RandomURL
End Function

' SearchBing : Search Bing site and collect rewards
' ----------------------------------------------------
Function SearchBing (intSearches)
	Dim i
	Dim objIE, Elem
	Dim RanURL
	
If blnDebug Then Echo "", "Entered SearchBing"
	On Error Resume next
	
	For i = 0 To intSearches
		If FindIE(objIE, 1) Then ' First time is a search for existing instance of IE with any title
			objIE.Navigate2 RandomURL, 2048
			If Err.Number <> 0 Then
				Echo "", "Error:, " & Err.Number & ", " & Err.Description
				Err.Clear
			End If			
			Do
				WScript.Sleep 2000 'Sleeps for 2 seconds
			Loop While objIE.Busy
			WScript.sleep Int((intMaxTime-IntMinTime+1)*Rnd()+intMinTime) * 100
		Else
			Set objIE = CreateObject("InternetExplorer.Application")
			With objIE
				.Navigate("https://login.live.com/login.srf")
				.Visible = True
			End With
			Do
				WScript.Sleep 2000 'Sleeps for 2 seconds
			Loop While objIE.Busy
			If un <> "" Then
				Set Elem = objIE.Document.getElementByID ("i0116")
				Elem.Value = un
			End If
			If pw <> "" Then
				Set Elem = objIE.Document.getElementByID ("i0118")
				Elem.Value = pw
			End If
			objIE.Document.getElementByID("idSIButton9").click
			Do
				WScript.Sleep 2000 'Sleeps for 2 seconds
			Loop While objIE.Busy
			objIE.Navigate("https://www.bing.com/rewards/dashboard")
			Do
				WScript.Sleep 2000 'Sleeps for 2 seconds
			Loop While objIE.Busy
		End If
	Next
	'Clears the object memory
	Set objIE = Nothing
	
	On Error Goto 0

End Function

' FindIE : Make sure IE is running
' ----------------------------------------------------
Function FindIE(obj, nCase)
	Dim w, btest, bFound

If blnDebug Then Echo "", "Entered FindIE"
	With CreateObject("shell.application")
		bFound = False
		For Each w In .windows
			If InStr(LCase(TypeName(w.document)),"htmldocument") <> 0 Then
				Select Case nCase
					Case 1 'First time is a search for existing instance of IE with any title
					btest = w.document.title <> ""
					Case 2 'Second time look for the "about:blank" window (no title)
					btest = w.document.title = "Sign in to your Microsoft account"
					Case 3 'Third time look for "" a blank window (no title)
					btest = w.document.title = ""
				End Select
				If btest Then Set obj = w : bFound = True : Exit For
			End If
		Next
	End With
	FindIE = bFound
End Function

' ====================================================
' * Functions & Subroutines
' ====================================================

' DominantArguments : Choose dominant argument list command line or ini file
' ----------------------------------------------------
Sub DominantArguments (strChoice)
	Dim objIniFile
	Dim colArgs
	
	If strTempPath = "" Then 
		strTempPath = objFSO.GetSpecialFolder(TemporaryFolder).Path & "\"
	Else
	End If

	If strPath = "" Then 
		strPath = Wscript.ScriptFullName
	Else
	End If

	Select Case LCase(strChoice)
	Case "ini"
		If WScript.Arguments.Count <> 0 Then
			Set colArgs = WScript.Arguments.Named
			fn = LCase(colArgs.Item("fn"))
			un = colArgs.Item("un")
			pw = colArgs.Item("pw")
 			For i = 0 To WScript.Arguments.Count - 1
				ProcessArguments (WScript.Arguments(i))
			Next
		End If
'		strPath = Wscript.ScriptFullName & "\"
		If objFSO.FileExists (strPath & strScriptName & ".ini") Then
			Set objIniFile = objFSO.OpenTextFile (strPath & strScriptName & ".ini", 1, False)
			Do Until objIniFile.AtEndOfStream
				ProcessArguments (objIniFile.ReadLine)
			Loop
			objIniFile.Close
			Set objIniFile = Nothing
		End If	
	Case Else
		If objFSO.FileExists (strPath & strScriptName & ".ini") Then
			Set objIniFile = objFSO.OpenTextFile (strPath & strScriptName & ".ini", 1, False)
			Do Until objIniFile.AtEndOfStream
				ProcessArguments (objIniFile.ReadLine)
			Loop
			objIniFile.Close
			Set objIniFile = Nothing
		End If	
		If WScript.Arguments.Count <> 0 Then
			Set colArgs = WScript.Arguments.Named
			fn = lcsse(colArgs.Item("fn"))
			un = colArgs.Item("un")
			pw = colArgs.Item("pw")
			For i = 0 To WScript.Arguments.Count - 1
				ProcessArguments (WScript.Arguments(i))
			Next
		End If
	End Select
	
	If WScript.Arguments.Count = 0 Then
		If blnReqArgs Then
			Echo Usage (WScript.ScriptName)
			WScript.Quit 1
		End If
	End If
End Sub

' ProcessArguments : Process arguments one at a time
' ----------------------------------------------------
Sub ProcessArguments (strArgument)
	blnHelp          = blnHelp        Or (strArgument="/?") Or (strArgument="-?")
	blnDebug         = blnDebug       Or (strArgument="/d") Or (strArgument="/D")
	blnLog           = blnLog         Or (strArgument="/l") Or (strArgument="/L")
	blnVerbose       = blnVerbose     Or (strArgument="/v") Or (strArgument="/V")
	blnQuiet         = blnQuiet       Or (strArgument="/q") Or (strArgument="/Q")
	blnAskRegistry   = blnAskRegistry Or (strArgument="/r") Or (strArgument="/R")
	
' 	strArgument = LCase (strArgument)
	Select Case Left (strArgument, 3)
		Case "fn:"
			fn = LCase(Mid (strArgument, 4, Len(strArgument) -3))
		Case "un:"
			un = Mid (strArgument, 4, Len(strArgument) -3)
		Case "pw:"
			pw = Mid (strArgument, 4, Len(strArgument) -3)
		Case Else
	End Select
	If blnHelp Then	' Show CommandLine Arguments and Quit
		Echo Usage (WScript.ScriptName)
		blnHelp = False
	End If
End Sub

' Usage : display program details and commandline arguments
' ----------------------------------------------------
Function Usage (strScriptName)
	Usage = Left (strScriptName, Len(strScriptName)-4) & "[" & Right (strScriptName, 4) & "]" & " -[arguments]" & vbCrLf _
	& "Copyright � 1984-" & Year(Now()) &" Nathan Davenport. All rights reserved." & vbCrLf _
	& vbCrLf _
	& "Description:" & vbCrLf _
	& "~Detail Description~" & vbCrLf _
	& vbCrLf _
	& "Arguments:  " & vbCrLf _
	& vbTab & "/? or -?		- Show help" & vbCrLf _
	& vbTab & "Arg1 - description" & vbCrLf _
	& vbTab & "Arg2 - description" & vbCrLf _
	& vbCrLf _
	& "Note:  " & vbCrLf _
	& "~Notes~" & vbCrLf _
	& vbTab & "Supply double quotes """" for any missing parameters."
End Function

' Cleanup : Close files, clear memory & prepare to close app
' ----------------------------------------------------
Sub CleanUp
	If strOutputFilename <> "" Then
		objOutputFile.Close
	End If
	
	If blnVerbose Or blnDebug Or blnLog Then
		Echo "", strScriptName & " Finished"
		Dim StopTime : StopTime = Now()
		Echo "", "Script Time: " & SecondsToTime(DateDiff("s", StartTime, StopTime))
		If strLogFilename <> "" Then objLogFile.Close
	End If
	
	If Not IsEmpty(objWMI) Then Set objWMI = Nothing
	If Not IsEmpty(objFSO) Then Set objFSO = Nothing
	If Not IsEmpty(objNet) Then Set objNet = Nothing
	If Not IsEmpty(objShell) Then Set objShell = Nothing
	If Not IsEmpty(objSysEnv) Then Set objSysEnv = Nothing
	If Not IsEmpty(objUsrEnv) Then Set objUsrEnv = Nothing
	If Not IsEmpty(objPrcEnv) Then Set objPrcEnv = Nothing
End Sub

' Echo : display a message & write log entry if set
' ----------------------------------------------------
Sub Echo (output, strMessage)
	If Not blnQuiet Then
		WScript.echo strMessage ' This should be the only line with wscript.echo
	End If
	
	If output <> "" Then
		If strOutputFilename = "" Then ' Open output file
			strOutputFilename = strPath & strScriptName & "." & output
			Set objOutputFile = objFSO.OpenTextFile(strOutputFilename, ForAppending, True)	
		End If
		objOutputFile.Write strMessage
	End If
	
	If blnLog Then 
		If strLogFilename = "" Then ' Open log file
			strLogFilename = strPath & strScriptName & ".log"
			'			strLogFilename = strPath & strScriptName & " - " & DTStamp ("both", False) & ".log"
			Set objLogFile = objFSO.OpenTextFile(strLogFilename, ForAppending, True)
		End If
		objLogFile.WriteLine DTStamp ("both", True) & "," & strMessage
	End If
End Sub

' ZeroPad : Pads number with leading zeros
' ----------------------------------------------------
Function ZeroPad (strNumber,intLength)
	strNumber = Trim(strNumber)
	
	If Len(strNumber) > intLength Then
		ZeroPad = strNumber
		Exit Function
	End If
	
	ZeroPad = String(intLength-Len(strNumber),"0") & strNumber
End Function

' SecondsToTime : Convert seconds to time format
' ----------------------------------------------------
Function SecondsToTime (intSeconds)
	Dim hours, minutes, seconds
	
	' calculates whole hours (like a div operator)
	hours = ZeroPad(intSeconds \ 3600,2)
	
	' calculates the remaining number of seconds
	intSeconds = intSeconds Mod 3600
	
	' calculates the whole number of minutes in the remaining number of seconds
	minutes = ZeroPad(intSeconds \ 60,2)
	
	' calculates the remaining number of seconds after taking the number of minutes
	seconds = ZeroPad(intSeconds Mod 60,2)
	
	' returns as a string
	SecondsToTime = hours & ":" & minutes & ":" & seconds
End Function

' DTStamp : formats Date and Time for logging & Filenames
' ----------------------------------------------------
Function DTStamp (DorT, Delim)
	Dim strDate
	Dim strYear 
	Dim strMonth 
	Dim strDay 
	Dim strHour 
	Dim strMinute 
	Dim strSecond 
	Dim DayDelim
	Dim TimeDelim
	Dim BothDelim
	
	strDate = Now()
	
	Select Case LCase(DorT)
		Case "b", "both"
		strYear = Year (strDate)
		strMonth = ZeroPad (Month (strDate),2)
		strDay = ZeroPad (Day (strDate),2)
		DayDelim = "-"
		
		strHour = ZeroPad (Hour (strDate),2)
		strMinute = ZeroPad (Minute (strDate),2)
		strSecond = ZeroPad (Second (strDate),2)
		TimeDelim = ":"
		BothDelim = " "
		Case "d", "date"
		strYear = Year (strDate)
		strMonth = ZeroPad (Month (strDate),2)
		strDay = ZeroPad (Day (strDate),2)
		DayDelim = "-"
		Case "t", "time"
		strHour = ZeroPad (Hour (strDate),2)
		strMinute = ZeroPad (Minute (strDate),2)
		strSecond = ZeroPad (Second (strDate),2)
		TimeDelim = ":"
		Case Else
	End Select
	
	If Delim Then
		DTStamp = strYear & DayDelim & strMonth & DayDelim & strDay & BothDelim & strHour & TimeDelim & strMinute & TimeDelim & strSecond
	Else
		DTStamp = strYear & strMonth & strDay & strHour & strMinute & strSecond
	End If
End Function

' OS : Get OS name and version
' ----------------------------------------------------
Function OS(opt)
	Dim colItems, objItem
	Dim OSName
	
If blnDebug Then Echo "", "Entered OS"
	Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem")
	For Each objItem In colItems
		Select Case LCase(opt)
			Case "name"
			OSName = Split(objItem.Name,"|") 
			OS = OSName(0)
			Case "ver"
			OS = objItem.Version
			Case "cap"
			OS = objItem.Caption
			Case "num"
			OS = objItem.BuildNumber
			Case "typ"
			OS = objItem.OSType
			Case "csv"
			OSName = Split(objItem.Name,"|") 
			OS = OS & "," & OSName(0)
			OS = OS & "," & objItem.Version
			OS = OS & "," & objItem.Caption
			OS = OS & "," & objItem.BuildNumber
			OS = OS & "," & objItem.BuildType
			OS = OS & "," & objItem.OSType
			OS = OS & "," & objItem.OtherTypeDescription
			OS = OS & "," & objItem.ServicePackMajorVersion & "." & objItem.ServicePackMinorVersion
			Case Else
			OSName = Split(objItem.Name,"|") 
			OS = "Name : " & OSName(0) & vbCrLf &_
			"Build : " & objItem.BuildNumber & vbCrLf &_
			"Install Date : " & Left(objItem.InstallDate,4) & "-" & Mid(objItem.InstallDate,5,2) & "-" & Right(objItem.InstallDate,2) & vbCrLf  
		End Select
	Next
End Function