#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent	; Keep the script running until the user exits it
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx
;CoordMode, Mouse, Window
;CoordMode, Pixel, Screen

Global SelectedFile := A_ScriptDir . "\config\ID.txt"
ConfigFile := A_ScriptDir . "\config\EKconfig.txt"
Global LogFile := A_ScriptDir . "\log\Log.txt"
ComIDFile := A_ScriptDir . "\config\ComID.txt"
Global ComID := 0

If !FileExist(LogFile)
{
	If !FileExist(A_ScriptDir . "\log")
		FileCreateDir, % A_ScriptDir . "\log"
}
LogEntry("Application Startup")

Menu, FileMenu, Add, Save Config, SaveConfig
Menu, FileMenu, Add, Exit, ExitSub
Menu, OptionsMenu, Add, AlwaysOnTop, OnTop
Menu, MyMenu, Add, &File, :FileMenu
Menu, MyMenu, Add, &Options, :OptionsMenu
Gui, Menu, MyMenu

GuiTitle=EliteKeep Extraction
Gui, Add, Tab3,, Query|Save Calls
Gui, add, Button, gB1 vB1 x15 y35, Query
Gui, add, Edit, w50 h100 r1 x85 y35 vB01X
Gui, add, Edit, w50 h100 r1 x140 y35 vB01Y
Gui, add, Button, gB1P x195 y35, GoTo
;Gui, add, Button, gB2 vB2 x15 y65, Edit
;Gui, add, Edit, w50 h100 r1 x85 y65 vB02X
;Gui, add, Edit, w50 h100 r1 x140 y65 vB02Y
Gui, add, Button, gB3 vB3 x15 y65, SegID
Gui, add, Edit, w50 h100 r1 x85 y65 vB03X
Gui, add, Edit, w50 h100 r1 x140 y65 vB03Y
;Gui, add, Button, gB4 vB4 x15 y95, SaveRun
;Gui, add, Edit, w50 h100 r1 x85 y95 vB04X
;Gui, add, Edit, w50 h100 r1 x140 y95 vB04Y
Gui, add, Button, gB5 vB5 x15 y95, GroupBy:
Gui, add, Edit, w50 h100 r1 x85 y95 vB05X
Gui, add, Edit, w50 h100 r1 x140 y95 vB05Y
Gui, add, Button, gB6 vB6 x15 y125, Top Result
Gui, add, Edit, w50 h100 r1 x85 y125 vB06X
Gui, add, Edit, w50 h100 r1 x140 y125 vB06Y
Gui, add, Button, gB7 vB7 x15 y155, Save Calls
Gui, add, Edit, w50 h100 r1 x85 y155 vB07X
Gui, add, Edit, w50 h100 r1 x140 y155 vB07Y

Gui, add, Button, w100 x300 y50 hwndhbuttonrunidprep vRunIDPrep gRunIDPrep, RunIDPrep
Gui, add, Button, w100 x300 y80 hwndhbuttonrunrobot vRunRobot gRunRobot, RunRobot

Gui, add, Text, x255 y115, General Timeout
Gui, add, Edit, w30 x340 y110 vTimeout, 50
Gui, add, Text, x375 y115, (sec)

Gui, add, Text, x245 y140, SaveCalls Timeout
Gui, add, Edit, w30 x340 y135 vSaveTimeout, 0
Gui, add, Text, x375 y140, (sec)

Gui, add, radio, x250 y170 gSegID vSegID Group checked, SegmentID
Gui, add, radio, x250 y195 gComID vComID, CompleteID

Gui, add, radio, x340 y170 gEngage vEngage Group, Engage 6.4+
Gui, add, radio, x340 y195 gNIM vNIM, NIM, 6.3

Gui, add, Text, x180 y220, Window Title:
Gui, add, Edit, x250 y215 vTitle, Application Suite

Gui, Add, Progress, x20 y245 w400 h18 vMyProgress
Gui, Add, Text, x15 y270 w300 vLoadingTxt

Gui, Tab, 2
Gui, add, Button, gB8 vB8 x15 y35, Loc Input
Gui, add, Edit, w50 h100 r1 x85 y35 vB08X
Gui, add, Edit, w50 h100 r1 x140 y35 vB08Y
Gui, add, Button, gB9 vB9 x15 y65, WAV Radio
Gui, add, Edit, w50 h100 r1 x85 y65 vB09X
Gui, add, Edit, w50 h100 r1 x140 y65 vB09Y
;Gui, add, Button, gB10 vB10 x15 y95, Save Btn
;Gui, add, Edit, w50 h100 r1 x85 y95 vB10X
;Gui, add, Edit, w50 h100 r1 x140 y95 vB10Y
Gui, add, Button, gB11 vB11 x15 y95, Three Dots
Gui, add, Edit, w50 h100 r1 x85 y95 vB11X
Gui, add, Edit, w50 h100 r1 x140 y95 vB11Y
Gui, add, Edit, w300 h100 r1 x15 y125 vUNCpath, C:\temp

GuiControl, Hide, B8
GuiControl, Hide, B08X
GuiControl, Hide, B08Y
GuiControl, Hide, B9
GuiControl, Hide, B09X
GuiControl, Hide, B09Y
GuiControl, Hide, B11
GuiControl, Hide, B11X
GuiControl, Hide, B11Y

Array := []
Loop, read, %ConfigFile%
{
	Array := StrSplit(A_LoopReadLine, ":")
	XorY := SubStr(Array[1], 4, 1)
	Num := SubStr(Array[1], 2, 2)
	If (XorY = "X")
	{
		B%Num%X := Array[2]
		GuiControl,, B%Num%X, % B%Num%X
	}
	Else
	{
		B%Num%Y := Array[2]
		GuiControl,, B%Num%Y, % B%Num%Y
	}
}
Gui, -dpiscale
Gui, Show,, EliteKeep Extraction
Return

OnTop:
CheckMarkToggle(A_ThisMenuItem, A_ThisMenu)
Return

CheckMarkToggle(MenuItem, MenuName)
{
	global
	%MenuItem%Flag := !%MenuItem%Flag ; Toggles the variable every time the function is called
	If (%MenuItem%Flag)
	{
		Menu, %MenuName%, Check, %MenuItem%
		Gui, 1: +AlwaysOnTop
		LogEntry("Setting AlwaysOnTop")
	}
	Else
	{
		Menu, %MenuName%, UnCheck, %MenuItem%
		Gui, 1: -AlwaysOnTop
		LogEntry("Removing AlwaysOnTop")
	}
}

SaveConfig:
LogEntry("Getting values from input boxes...")
Loop, 11
{
	If (A_Index < 10)
	{
		GuiControlGet, B0%A_Index%X
		GuiControlGet, B0%A_Index%Y
	}
	If (A_Index >= 10)
	{
		GuiControlGet, B%A_Index%X
		GuiControlGet, B%A_Index%Y
	}
}
If FileExist(ConfigFile)
{
	LogEntry("Existing config file, asking user to overwrite")
	MsgBox,4,Delete Config, Do you want to overwrite the existing config?
}
IfMsgBox Yes
{
	LogEntry("Permission granted, deleting existing configuration file: " . ConfigFile)
	FileDelete, %ConfigFile%
}
IfMsgBox No
{
	LogEntry("Permission denied, returning with no changes.")
	Return
}
If !FileExist(A_ScriptDir . "\config")
{
	LogEntry("Config folder missing, likely first run - creating missing config folder")
	FileCreateDir, % A_ScriptDir . "\config"
}
Loop, 11
{
	If (A_Index < 10)
	{
		tmpX := B0%A_Index%X
		tmpY := B0%A_Index%Y
		Sleep, 50
		FileAppend, % "B0" . A_Index . "X:" . tmpX . "`n", %ConfigFile%
		Sleep, 100
		FileAppend, % "B0" . A_Index . "Y:" . tmpY . "`n", %ConfigFile%
	}
	If (A_Index >= 10)
	{
		tmpX := B%A_Index%X
		tmpY := B%A_Index%Y
		Sleep, 50
		FileAppend, % "B" . A_Index . "X:" . tmpX . "`n", %ConfigFile%
		Sleep, 100
		FileAppend, % "B" . A_Index . "Y:" . tmpY . "`n", %ConfigFile%
	}
}
If FileExist(ConfigFile)
	MsgBox,, Save Config, Save Config Task Complete
Else
	MsgBox,, Save Error, Error running Save Config - permissions to config folder?
Return

SegID:
GuiControl, Show, B3
GuiControl, Show, B03X
GuiControl, Show, B03Y
Return

ComID:
GuiControl, Hide, B3
GuiControl, Hide, B03X
GuiControl, Hide, B03Y
Return

Engage:
GuiControl, Hide, B11
GuiControl, Hide, B11X
GuiControl, Hide, B11Y
GuiControl, Show, B8
GuiControl, Show, B08X
GuiControl, Show, B08Y
GuiControl, Show, B9
GuiControl, Show, B09X
GuiControl, Show, B09Y
Return

NIM:
GuiControl, Show, B11
GuiControl, Show, B11X
GuiControl, Show, B11Y
GuiControl, Hide, B8
GuiControl, Hide, B08X
GuiControl, Hide, B08Y
GuiControl, Hide, B9
GuiControl, Hide, B09X
GuiControl, Hide, B09Y
Return

B1:		; Query
GuiControl, Text, B1, Working
TrayTip, Waiting..., Waiting for query X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B01X, B01Y
GuiControl,, B01X, %B01X%
GuiControl,, B01Y, %B01Y%
GuiControl, Text, B1, Query
HideTrayTip()
MsgBox,,Coords, %B01X% %B01Y%
LogEntry("Query coordinates logged: " . B01X . "," . B01Y)
Return

B2:		; Edit
GuiControl, Text, B2, Working
TrayTip, Waiting..., Waiting for Edit X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B02X, B02Y
GuiControl,, B02X, %B02X%
GuiControl,, B02Y, %B02Y%
GuiControl, Text, B2, Edit
HideTrayTip()
MsgBox,,Coords, %B02X% %B02Y%
LogEntry("Edit coordinates logged: " . B02X . "," . B02Y)
Return

B3:		; SegID
GuiControl, Text, B3, Working
TrayTip, Waiting..., Waiting for SegID X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B03X, B03Y
GuiControl,, B03X, %B03X%
GuiControl,, B03Y, %B03Y%
GuiControl, Text, B3, SegID
HideTrayTip()
MsgBox,,Coords, %B03X% %B03Y%
LogEntry("SegID coordinates logged: " . B03X . "," . B03Y)
Return

B4:		; SaveRun
GuiControl, Text, B4, Working
TrayTip, Waiting..., Waiting for SaveRun X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B04X, B04Y
GuiControl,, B04X, %B04X%
GuiControl,, B04Y, %B04Y%
GuiControl, Text, B4, SaveRun
HideTrayTip()
MsgBox,,Coords, %B04X% %B04Y%
LogEntry("Save&Run coordinates logged: " . B04X . "," . B04Y)
Return

B5:		; GroupBy
GuiControl, Text, B5, Working
TrayTip, Waiting..., Waiting for GroupBy X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B05X, B05Y
GuiControl,, B05X, %B05X%
GuiControl,, B05Y, %B05Y%
GuiControl, Text, B5, GroupBy
HideTrayTip()
MsgBox,,Coords, %B05X% %B05Y%
LogEntry("GroupBy Box coordinates logged: " . B05X . "," . B06Y)
Return


B6:		; Top Result
GuiControl, Text, B6, Working
TrayTip, Waiting..., Waiting for Top Result X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B06X, B06Y
GuiControl,, B06X, %B06X%
GuiControl,, B06Y, %B06Y%
GuiControl, Text, B6, Top Result
HideTrayTip()
MsgBox,,Coords, %B06X% %B06Y%
LogEntry("Top Result coordinates logged: " . B06X . "," . B06Y)
Return

B7:		; Save Calls
GuiControl, Text, B7, Working
TrayTip, Waiting..., Waiting for Save Calls X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B07X, B07Y
GuiControl,, B07X, %B07X%
GuiControl,, B07Y, %B07Y%
GuiControl, Text, B7, Save Calls
HideTrayTip()
MsgBox,,Coords, %B07X% %B07Y%
LogEntry("Save Calls Button coordinates logged: " . B07X . "," . B07Y)
Return

B8:		; Loc Input
GuiControl, Text, B8, Working
TrayTip, Waiting..., Waiting for Loc Input X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B08X, B08Y
GuiControl,, B08X, %B08X%
GuiControl,, B08Y, %B08Y%
GuiControl, Text, B8, Loc Input
HideTrayTip()
MsgBox,,Coords, %B08X% %B08Y%
LogEntry("Loc Input in Save calls tab coordinates logged: " . B08X . "," . B08Y)
Return

B9:		; WAV Radio
GuiControl, Text, B9, Working
TrayTip, Waiting..., Waiting for WAV Radio X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B09X, B09Y
GuiControl,, B09X, %B09X%
GuiControl,, B09Y, %B09Y%
GuiControl, Text, B9, WAV Radio
HideTrayTip()
MsgBox,,Coords, %B09X% %B09Y%
LogEntry("WAV radio button coordinates logged: " . B09X . "," . B09Y)
Return

B10:	; Save Btn
GuiControl, Text, B10, Working
TrayTip, Waiting..., Waiting for Save Btn X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B10X, B10Y
GuiControl,, B10X, %B10X%
GuiControl,, B10Y, %B10Y%
GuiControl, Text, B10, Save Btn
HideTrayTip()
MsgBox,,Coords, %B10X% %B10Y%
LogEntry("Save button coordinates logged: " . B10X . "," . B10Y)
Return

B11:	; 3dots
GuiControl, Text, B11, Working
TrayTip, Waiting..., Waiting for three dots X`,Y coords, 30, 1
KeyWait, Enter, D
MouseGetPos, B11X, B11Y
GuiControl,, B11X, %B11X%
GuiControl,, B11Y, %B11Y%
GuiControl, Text, B11, Three Dots
HideTrayTip()
MsgBox,,Coords, %B11X% %B11Y%
LogEntry("Save button coordinates logged: " . B11X . "," . B11Y)
Return

B1P:
Gui, Submit, NoHide
WinActivate, %Title%
MouseMove, %B01X%, %B01Y%
MouseClick, R, %B01X%, %B01Y%
LogEntry("Right click on: (" . B01X . "," . B01Y . ")")
Return

RunIDPrep:
GuiControl, Text, RunIDPrep, RunningPrep
LogEntry("Run ID Prep Start (from file:) " . SelectedFile)
If !FileExist(SelectedFile)
{
	LogEntry("Missing " . SelectedFile . " file path...")
	MsgBox,, File Error, Missing %SelectedFile% file path...
	GuiControl, Text, RunIDPrep, RunIDPrep
	Return
}
If !FileExist(A_ScriptDir . "\Processing")
{
	LogEntry("Processing folder missing, likely first run - creating missing Processing folder")
	FileCreateDir, % A_ScriptDir . "\Processing"
}
LogEntry("Count number of lines")
FileRead, File, %SelectedFile%
StringReplace File, File, `n, `n, All UseErrorLevel
TotalLines := ErrorLevel
TotalLines++	; Add one as it was always one short since the last line did not have a `n at the end... Of course if there is an emtpy line it will count the extra emtpy line.  Oh well.
LogEntry("Total number of SegID's is: " . TotalLines)
GuiControl, +Range0-%TotalLines%, MyProgress
Gui, Show
FileRead, InputFile, %SelectedFile%
IDArray := StrSplit(InputFile, "`n")
GuiControl, Text, LoadingTxt, Array Complete: %TotalLines%
LogEntry("Loaded ID's from " . SelectedFile)
LogEntry("Before FuncLoop")
GuiControl, Text, LoadingTxt, Concatenation...
FuncLoop(IDArray)
LogEntry("After FuncLoop")
GuiControl, Text, LoadingTxt, Concatenation...Complete
GuiControl, Text, RunIDPrep, Run ID Prep
Return

RunRobot:
Gui, Submit, NoHide
GuiControl, Text, RunRobot, RunningRobot
countfiles := 0
Loop, Files, Processing\*.txt
   countfiles += 1
If (countfiles < 1)
{
	MsgBox,, Robot Error, Run Robot failed - did you prep the ID's first?
	LogEntry("FAILURE - countfiles var (" . countfiles . ") less than one...Returning")
	GuiControl, Text, RunRobot, RunRobot
	Return
}
If (!Engage and !NIM)
{
	MsgBox,, Select System, Please select system version
	LogEntry("FAILURE - User did not select system version...Returning")
	GuiControl, Text, RunRobot, RunRobot
	Return
}
If !FileExist(UNCPath)
{
	LogEntry("Attempting to create save calls path: " . UNCPath)
	FileCreateDir, %UNCPath%
	If !FileExist(UNCPath)
	{
		MsgBox,, Save Path, Save calls path cannot be created.
		LogEntry("FAILURE - Save calls path is not available")
		GuiControl, Text, RunRobot, RunRobot
		Return
	}
}
GroupAdd,WindowGroup,, styledButton4
GroupAdd,WindowGroup, Save Calls
GuiControlGet, Timeout
GuiControlGet, SaveTimeout
Timeoutms := Timeout*1000
SaveTimeoutms := SaveTimeout*1000
TotalTime := A_TickCount
TotalElapsed := 0
GuiControl, +Range0-%countfiles%, MyProgress
GuiControl, , MyProgress, 0
GuiControl, Text, LoadingTxt, Robot Working: 0 of %countfiles%
LogEntry("Robot Starting...0 of " . countfiles)
Loop, %countfiles%
{
	Gui, Submit, NoHide
	DriveSpaceFree, FreeSpaceArchive, %UNCPath%
	LogEntry("Free Space: " . FreeSpaceArchive . " MB in " . UNCPath)
	If (FreeSpaceArchive < 1024)
	{
		LogEntry("Free Space LOW: " . FreeSpaceArchive . " MB in " . UNCPath)
		MsgBox,4,Free Space, Warning: Free space is low on %UNCPath% - %FreeSpaceArchive% MB.  End Task?
		IfMsgBox Yes
		{
			LogEntry("User chose to stop based on low disk space")
			GuiControl, Text, RunRobot, RunRobot
			GuiControl, Text, LoadingTxt, Robot Ended. %A_Index% of %countfiles%
			Return
		}
	}
	CurFile := 
	Loop, Files, Processing\*.txt
	{
		CurFile := A_LoopFilePath
		break
	}
	FileRead, IDvar, %CurFile%
	LoopTime := A_TickCount
	LoopElapsed := 0
	RightClick :=
	GroupByBoxColor :=
	LogEntry("Waiting for '" . Title . "' to become active")
	TrayTip, Waiting..., Waiting for %Title%, %Timeout%, 1
	WinWaitActive, %Title%,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for " . Title)
		MsgBox,, Timeout, Timeout, please focus the %Title% window
		TrayTip, Waiting..., Waiting for %Title%, %Timeout%, 1
		WinWaitActive, %Title%,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
		HideTrayTip()
	}
	HideTrayTip()
	If (SegID = 1)
	{
		Sleep, 400
		LogEntry("SegmentID selected, running routine for SegmentID")
		MouseClick, R, %B01X%, %B01Y%  ; Right click on query
		LogEntry("Right click on query at (" . B01X . "`," . B01Y . ")")
		TrayTip, Query, Right-Click on Query, %Timeout%, 1
		LogEntry("Before While loop - waiting for menu")
		StartTime := A_TickCount
		ElapsedTime := 0
		ErrorTimeout := 0
		Sleep, 500
		Loop 5
		{
			Sleep 30
			Send {Down}
		}
		Send {Enter}
		HideTrayTip()
		SkipClickEdit:
		LogEntry("Waiting for Advanced Query window")
		TrayTip, Waiting..., Waiting for Advanced Query Window, %Timeout%, 1
		WinWait, Advanced Query,,%Timeout%
		WinActivate, Advanced Query
		WinWaitActive, Advanced Query,,%Timeout%
		If ErrorLevel
		{
			LogEntry("Timeout waiting for Advanced Query Window")
			MsgBox,, Timeout, Timeout waiting for Advanced Query Window, was edit pressed?  Manually edit query and press OK.
			TrayTip, Waiting..., Waiting for Advanced Query Window, %Timeout%, 1
			WinWaitActive, Advanced Query,,%Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
			HideTrayTip()
		}
		HideTrayTip()
		MouseClick,, %B03X%, %B03Y% ; Click on SegID box
		LogEntry("Click on SegID Box")
		MouseGetPos,,,,SegIDControl
		ControlSetText,%SegIDControl%, % IDvar, Advanced Query
		LogEntry("Inserting " . IDvar)
		TrayTip, Insert, Inserting SegID's, %Timeout%, 1
		LogEntry("Click on 'Save&Run'")
		ControlClick, Save & Run, Advanced Query
		HideTrayTip()
		LogEntry("Waiting for '" . Title . "' to become active")
		TrayTip, Waiting..., Waiting for %Title%, %Timeout%, 1
		WinWait, %Title%,,%Timeout%
		WinActivate, %Title%
		WinWaitActive, %Title%,,%Timeout%
		If ErrorLevel
		{
			LogEntry("Timeout waiting for " . Title)
			MsgBox,, Timeout, Timeout waiting for %Title%, please focus the %Title% (was 'Save & Run' pressed?)
			TrayTip, Waiting..., Waiting for %Title%, %Timeout%, 1
			WinWaitActive, %Title%,,%Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
			HideTrayTip()
		}
		HideTrayTip()
	}
	If (ComID = 1)
	{
		LogEntry("CompleteID selected, running routine for CompleteID")
		If FileExist(ComIDFile)
			FileDelete, %ComIDFile%
		FileCopy, %CurFile%, %ComIDFile%
		if FileExist("GECQueryUpdater.exe")
		{
			LogEntry("Running CompleteID script to modify DB")
			TrayTip, Waiting..., Waiting for script to push to DB, %Timeout%, 1
			RunWait %ComSpec% /c ""GECQueryUpdater.exe" "%ComIDFile%"",,Hide
			HideTrayTip()
		}
		else
		{
			MsgBox,, Missing Binary, Missing GECQueryUpdater.exe
			FileSelectFile, GECQueryFileLoc, 1, , Select GECQueryUpdater Location, *.exe
			LogEntry("Running CompleteID script to modify DB")
			TrayTip, Waiting..., Waiting for script to push to DB, %Timeout%, 1
			RunWait %ComSpec% /c ""%GECQueryFileLoc%" "%ComIDFile%"",,Hide
			HideTrayTip()
		}
		LogEntry("Left clicking on updated query at (" . B01X . "," B01Y . ")")
		MouseClick, L, %B01X%, %B01Y%  ; Left click on query
		StartTime := A_TickCount
		ElapsedTime := 0
		ErrorTimeout := 0
		TrayTip, Waiting..., Waiting for query to start, %Timeout%, 1
		While (GroupByBoxColor != 0xCED3D6 or GroupByBoxColor != 0xD6D3CE)
		{
			If (ElapsedTime > Timeoutms)
			{
				LogEntry("Elapsed time (" . ElapsedTime . ") > Timeoutms (" . Timeoutms . ") , breaking loop")
				ErrorTimeout := 1
				break
			}
			PixelGetColor, GroupByBoxColor, %B05X%, %B05Y%
			LogEntry("Got Color " . GroupByBoxColor . " at coodinate (" . B05X . "`," . B05Y . ")")
			ElapsedTime := A_TickCount - StartTime
			LogEntry("Time elapsed: " . ElapsedTime)
			If (GroupByBoxColor = 0xFFFFFF)
			{
				If (ElapsedTime < 500)
					continue
				Else
					break
			}
		}
		HideTrayTip()
	}
	StartTime := A_TickCount
	ElapsedTime := 0
	ErrorTimeout := 0
	TrayTip, Waiting..., Waiting for query to finish, %Timeout%, 1
	While (GroupByBoxColor != 0xFFFFFF) ; Wait for GroupBy to go white (indicates query is done)
	{
		If (ElapsedTime > Timeoutms)
		{
			LogEntry("Elapsed time (" . ElapsedTime . ") > Timeoutms (" . Timeoutms . ") , breaking loop")
			ErrorTimeout := 1
			break
		}
		PixelGetColor, GroupByBoxColor, %B05X%, %B05Y%
		LogEntry("Got Color " . GroupByBoxColor . " at coodinate (" . B05X . "`," . B05Y . ")")
		ElapsedTime := A_TickCount - StartTime
		LogEntry("Time elapsed: " . ElapsedTime)
	}
	If (ErrorTimeout = 1)
		MsgBox,, Error Timeout, Please ensure query ran successfully, timer breaking loop required to move along.
	HideTrayTip()
	LogEntry("Clicking on top result (" . B06X . "`," . B06Y . ")")
	MouseClick,, %B06X%, %B06Y%
	LogEntry("Sending ctrl-a")
	Send ^a
	TrayTip, Click, Clicking on Save Calls button..., 10, 1
	LogEntry("Click Save Calls (" . B07X . "`," . B07Y . ")")
	MouseClick,, %B07X%, %B07Y%
	LogEntry("Wait for either Save Calls window OR the 'No Recordings Found' window")
	WinWait, ahk_group WindowGroup,, %Timeout%
	LogEntry("Check which window we have...")
	If WinActive(,"styledButton4")
	{
		LogEntry("In IF statement for info box - means there are no recordings")
		WinActivate,, Information
		Send, {Esc}
		GuiControl,, MyProgress, %A_Index%
		GuiControl, Text, LoadingTxt, Robot Completed: %A_Index% of %countfiles%
		If !FileExist(A_ScriptDir . "\Failed")
			FileCreateDir, % A_ScriptDir . "\Failed"
		FileMove, %CurFile%, %A_ScriptDir%\Failed\
		continue
	}
	HideTrayTip()
	LogEntry("Wait for SaveCalls dialog")
	TrayTip, Waiting..., Waiting for Save Calls Dialog box..., %Timeout%, 1
	WinWait, Save Calls,,%Timeout%
	Sleep, 200
	WinActivate, Save Calls
	WinWaitActive, Save Calls,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Save Calls Dialog")
		MsgBox,, Timeout, Timeout waiting for Save Calls Dialog - ensure calls are selected and press 'save calls' button, then press OK.
		TrayTip, Waiting..., Waiting for Save Calls Dialog box..., %Timeout%, 1
		WinWaitActive, Save Calls,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
		HideTrayTip()
	}
	HideTrayTip()
	If (Engage)
	{
		LogEntry("User selected Engage 6.4+, going down that save calls path...")
		LogEntry("Click 'location' inputbox (" . B08X . "`," . B08Y . ")")
		MouseClick,, %B08X%, %B08Y%
		MouseGetPos,,,,LocControl
		LogEntry("ClassNN - " . LocControl)
		ControlSetText, %LocControl%, %UNCPath%
		LogEntry("Inject save calls path (" . UNCPath . ") to location input box")
		MouseClick,, %B09X%,%B09Y%
		LogEntry("Click on WAV radio btn")
		TrayTip, Click, Clicked Save, %Timeout%, 1
		ControlClick, Save, Save Calls
		LogEntry("Click on Save Btn")
		MouseGetPos,,,,InfoControl
		LogEntry("Move mouse to Location input, checking for 'info' screen...")
		Sleep, 500
		HideTrayTip()
		InfoBox := WinExist("ahk_exe PresentationHost.exe","styledButton4")
		LogEntry("InfoBox: " . InfoBox . " should only populate when there is the 'info' window, in which case we need to hit esc")
		If (InfoBox != "0x0")
		{
			LogEntry("In IF statement for InfoBox")
			WinActivate,, Information
			Send, {Esc}
		}
		LogEntry("Waiting for 'Saving' Dialog box")
		TrayTip, Waiting..., Waiting for Save Calls to show..., %Timeout%, 1
		WinWait, Saving,,%Timeout%
		WinActivate, Saving
		WinWaitActive, Saving,,%Timeout%
		If ErrorLevel
		{
			LogEntry("Timeout waiting for Saving Dialog")
			MsgBox,, Timeout, Timeout waiting for Saving Dialog - was 'Save Calls' pressed?
			TrayTip, Waiting..., Waiting for Save Calls to show..., %Timeout%, 1
			WinWaitActive, Saving,,%Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
			HideTrayTip()
		}
		HideTrayTip()
		LogEntry("Waiting for 'Open File Location' button on 'Saving/Done' dialog (indicates task is complete)")
		TrayTip, Waiting..., Waiting for Save Calls to Complete..., 5, 1
		WinWait,, Open File Location,%Timeout%
		WinActivate,, Open File Location
		WinWaitActive,, Open File Location,%Timeout%
		If ErrorLevel
		{
			LogEntry("Timeout waiting for 'Open File Location' Button on 'Save Calls'")
			MsgBox,, Timeout, Timeout waiting for 'Open File Location' button after save calls - press OK after it is complete
			WinWaitActive,, Open File Location,%Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, will move to next step.
				GoTo WaitAppEnd
			}
		}
		HideTrayTip()
		Sleep, SaveTimeoutms
		Send, {Tab}
		LogEntry("Tab once")
		Send, {Tab}
		LogEntry("Tab twice")
		Send, {Space}
		LogEntry("Space to close the window...")
		WaitAppEnd:
		TrayTip, Waiting..., Wait for %Title%, %Timeout%, 1
		WinWait, %Title%,,%Timeout%
		WinActivate, %Title%
		WinWaitActive, %Title%,,%Timeout%
		If ErrorLevel
		{
			LogEntry("Timeout waiting for " . Title)
			MsgBox,, Timeout, Please focus the %Title% window
			WinWaitActive, %Title%,,%Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
		}
		HideTrayTip()
	}
	If (NIM)
	{
		LogEntry("User selected the legacy NIM/6.3 interface, going down that save calls path...")
		WinActivate, Save Calls		
		LogEntry("Tab 3x to get to file prefix checkbox")
		Loop, 3
		{
			Send, {Tab}
			Sleep, 30
		}
		LogEntry("Send space to select file prefix box")
		Send, {Space}
		LogEntry("Click on three dot button")
		MouseClick,, %B11X%, %B11Y%
		LogEntry("Waiting for 'Save as' dialog")
		TrayTip, Waiting..., Waiting for Save As dialog..., %Timeout%, 1
		WinWaitActive, Save as,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout, Timeout waiting for 'save as' dialog
			TrayTip, Waiting..., Waiting for Save As dialog..., %Timeout%, 1
			WinWaitActive, Save as,, %Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
			HideTrayTip()
		}
		HideTrayTip()
		WinActivate, Save as
		UNCPathNIM := UNCPath "\Call1.wav"
		LogEntry("Injecting path: " . UNCPathNIM)
		ControlSetText, Edit1, %UNCPathNIM%
		Send, {Enter}
		Sleep, 300
		InfoBox := WinExist("","Information")
		LogEntry("InfoBox: " . InfoBox . " should only populate when there is the 'info' window, in which case we need to hit esc")
		If (InfoBox != "0x0")
		{
			LogEntry("In IF statement for InfoBox")
			WinActivate,, Information
			Send, {Esc}
		}
		WaitforClose:
		LogEntry("Waiting for 'Save' button to show (indicating save calls is ready)")
		SaveCallsText :=
		WinGetText, SaveCallsText, Save Calls
		Loop, parse, SaveCallsText, `r
		{
			If (A_Index = 1)
			{
				If (A_LoopField = "Save")
					GoTo Closer
				Else
					GoTo WaitforClose
			}
		}
		Closer:
		LogEntry("Clicking on Save btn")
		ControlClick, Save, Save Calls
		WaitforSave:
		LogEntry("Waiting for 'Close' button to show (indicating save calls is complete)")
		SaveCallsText :=
		WinGetText, SaveCallsText, Save Calls
		Loop, parse, SaveCallsText, `r
		{
			If (A_Index = 1)
			{	
				If (A_LoopField = "Close")
					GoTo Close
				Else
					GoTo WaitforSave
			}
		}
		Close:
		Sleep, SaveTimeoutms
		LogEntry("Clicking on Close btn")
		ControlClick, Close, Save Calls
	}
	EndStep:
	LoopElapsed := A_TickCount - LoopTime
	LoopElapsed := LoopElapsed / 1000
	LoopElapsed := Round(LoopElapsed)
	LogEntry("Completed loop: " . A_Index . " of " . countfiles . " in " . LoopElapsed . "sec.")
	If !FileExist(A_ScriptDir . "\Processed")
		FileCreateDir, % A_ScriptDir . "\Processed"
	FileMove, %CurFile%, %A_ScriptDir%\Processed\
	TrayTip, Loop End, End of loop %A_Index% of %countfiles%, %Timeout%, 1
	GuiControl,, MyProgress, %A_Index%
	GuiControl, Text, LoadingTxt, Robot Completed: %A_Index% of %countfiles%
}
TotalElapsed := A_TickCount - TotalTime
TotalElapsed := TotalElapsed / 1000
TotalElapsed := Round(TotalElapsed)
LogEntry("Task Complete in " . TotalElapsed . "sec.  Go have a beer.")
MsgBox,, Complete, Done in %TotalElapsed%sec
GuiControl, Text, RunRobot, RunRobot
Return

FuncLoop(CompleteIDArray)
{
	LogEntry("In FuncLoop - Setting up files")
	GuiControlGet, ComID
	NewIDArray := []
	IDArray := []
	ID :=
	Increment := 1
	for k, v in CompleteIDArray
	{
		NewIDArray.Push(v)
		GuiControl, , MyProgress, %A_Index%
		ArrayLength := CompleteIDArray.length()
		GuiControl, Text, LoadingTxt, Concatenation...  %k% of %ArrayLength%
		;check if we are at the 20th item OR if we are at the last item (which could be a weird number)
		if (NewIDArray.length()=20 || a_index = CompleteIDArray.length())
		{
			if (ComID = 0)
			{
				for k2, v2 in NewIDArray
				{
					ID := ID NewIDArray[k2] ";"
				}
			}
			if (ComID = 1)
			{
				for k2, v2 in NewIDArray
				{
					ID := ID NewIDArray[k2] "`n"
				}
			}
			ProcessingFile := "Processing\" . Increment . ".txt"
			If FileExist(ProcessingFile)
				FileDelete, %ProcessingFile%
			FileAppend, %ID%, %ProcessingFile%
			LogEntry("Add to IDArray at index: " . Increment . " Create " . IDArray[Increment])
			Increment++
			ID:=
			NewIDArray:=[] ; clear the list to get the next 20
		}
	}
	LogEntry("End of FuncLoop - Return var")
	return IDArray
}

LogEntry(Message)
{
	Sleep, 100
	FileAppend, %A_MM%/%A_DD%/%A_YYYY% %A_Hour%:%A_Min%:%A_Sec%.%A_MSec% - %Message%`n, %LogFile%
	Sleep, 100
}

HideTrayTip()
{
    TrayTip  ; Attempt to hide it the normal way.
    if SubStr(A_OSVersion,1,3) = "10." {
        Menu Tray, NoIcon
        Sleep 200  ; It may be necessary to adjust this sleep.
        Menu Tray, Icon
    }
}

ExitSub:
LogEntry("Exit Application")
GuiClose:
GuiEscape:
ExitApp
Exit