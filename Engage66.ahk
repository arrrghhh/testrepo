#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent	; Keep the script running until the user exits it
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx
;CoordMode, Mouse, Window
;CoordMode, Pixel, Screen

SelectedFile := A_ScriptDir . "\config\ID.txt"
ConfigFile := A_ScriptDir . "\config\EKconfig.txt"
Global LogFile := A_ScriptDir . "\log\Log.txt"

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
Gui, add, Button, gB1 x15 y35, Query
Gui, add, Edit, w50 h100 r1 x85 y35 vB01X
Gui, add, Edit, w50 h100 r1 x140 y35 vB01Y
Gui, add, Button, gB1P x195 y35, GoTo
Gui, add, Button, x15 y65 gB2, Edit
Gui, add, Edit, w50 h100 r1 x85 y65 vB02X
Gui, add, Edit, w50 h100 r1 x140 y65 vB02Y
Gui, add, Button, x15 y95 gB3, SegID
Gui, add, Edit, w50 h100 r1 x85 y95 vB03X
Gui, add, Edit, w50 h100 r1 x140 y95 vB03Y
Gui, add, Button, x15 y125 gB4, Save&Run
Gui, add, Edit, w50 h100 r1 x85 y125 vB04X
Gui, add, Edit, w50 h100 r1 x140 y125 vB04Y
Gui, add, Button, x15 y155 gB5, GroupBy:
Gui, add, Edit, w50 h100 r1 x85 y155 vB05X
Gui, add, Edit, w50 h100 r1 x140 y155 vB05Y
Gui, add, Button, x15 y185 gB6, Top Result
Gui, add, Edit, w50 h100 r1 x85 y185 vB06X
Gui, add, Edit, w50 h100 r1 x140 y185 vB06Y
Gui, add, Button, x15 y215 gB7, Save Calls
Gui, add, Edit, w50 h100 r1 x85 y215 vB07X
Gui, add, Edit, w50 h100 r1 x140 y215 vB07Y

Gui, add, Button, w100 x300 y50 hwndhbuttonrunidprep vRunIDPrep gRunIDPrep, RunIDPrep
Gui, add, Button, w100 x300 y80 hwndhbuttonrunrobot vRunRobot gRunRobot, RunRobot

Gui, add, Text, x255 y115, General Timeout
Gui, add, Edit, w30 x340 y110 vTimeout, 50
Gui, add, Text, x375 y115, (sec)

Gui, add, Text, x245 y140, SaveCalls Timeout
Gui, add, Edit, w30 x340 y135 vSaveTimeout, 6
Gui, add, Text, x375 y140, (sec)

Gui, add, radio, x340 y170 vEngage, Engage 6.4+
Gui, add, radio, x340 y195 vNIM, NIM, 6.3

Gui, add, Edit, x250 y215 vTitle, Application Suite

Gui, Add, Progress, x8 y245 w400 h18 vMyProgress
Gui, Add, Text, x15 y270 w300 vLoadingTxt

Gui, Tab, 2
Gui, add, Button, gB8 x15 y35, Loc Input
Gui, add, Edit, w50 h100 r1 x85 y35 vB08X
Gui, add, Edit, w50 h100 r1 x140 y35 vB08Y
Gui, add, Button, gB9 x15 y65, WAV Radio
Gui, add, Edit, w50 h100 r1 x85 y65 vB09X
Gui, add, Edit, w50 h100 r1 x140 y65 vB09Y
Gui, add, Button, gB10 x15 y95, Save Btn
Gui, add, Edit, w50 h100 r1 x85 y95 vB10X
Gui, add, Edit, w50 h100 r1 x140 y95 vB10Y
Gui, add, Button, gB11 x15 y125, Three Dots
Gui, add, Edit, w50 h100 r1 x85 y125 vB11X
Gui, add, Edit, w50 h100 r1 x140 y125 vB11Y
Gui, add, Button, gB12 x15 y160, Save Btn
Gui, add, Edit, w50 h100 r1 x85 y160 vB12X
Gui, add, Edit, w50 h100 r1 x140 y160 vB12Y
Gui, add, Edit, w300 h100 r1 x15 y195 vUNCpath, C:\temp

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
Loop, 12
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
Loop, 12
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

B1:		; Query
KeyWait, Enter, D
MouseGetPos, B01X, B01Y
MsgBox,,Coords, %B01X% %B01Y%
GuiControl,, B01X, %B01X%
GuiControl,, B01Y, %B01Y%
LogEntry("Query coordinates logged: " . B01X . "," . B01Y)
Return

B2:		; Edit
KeyWait, Enter, D
MouseGetPos, B02X, B02Y
MsgBox,,Coords, %B02X% %B02Y%
GuiControl,, B02X, %B02X%
GuiControl,, B02Y, %B02Y%
LogEntry("Edit coordinates logged: " . B02X . "," . B02Y)
Return

B3:		; SegID
KeyWait, Enter, D
MouseGetPos, B03X, B03Y
MsgBox,,Coords, %B03X% %B03Y%
GuiControl,, B03X, %B03X%
GuiControl,, B03Y, %B03Y%
LogEntry("SegID coordinates logged: " . B03X . "," . B03Y)
Return

B4:		; SaveRun
KeyWait, Enter, D
MouseGetPos, B04X, B04Y
MsgBox,,Coords, %B04X% %B04Y%
GuiControl,, B04X, %B04X%
GuiControl,, B04Y, %B04Y%
LogEntry("Save&Run coordinates logged: " . B04X . "," . B04Y)
Return

B5:		; GroupBy
KeyWait, Enter, D
MouseGetPos, B05X, B05Y
MsgBox,,Coords, %B05X% %B05Y%
GuiControl,, B05X, %B05X%
GuiControl,, B05Y, %B05Y%
LogEntry("GroupBy Box coordinates logged: " . B05X . "," . B06Y)
Return


B6:		; Top Result
KeyWait, Enter, D
MouseGetPos, B06X, B06Y
MsgBox,,Coords, %B06X% %B06Y%
GuiControl,, B06X, %B06X%
GuiControl,, B06Y, %B06Y%
LogEntry("Top Result coordinates logged: " . B06X . "," . B06Y)
Return

B7:		; Save Calls
KeyWait, Enter, D
MouseGetPos, B07X, B07Y
MsgBox,,Coords, %B07X% %B07Y%
GuiControl,, B07X, %B07X%
GuiControl,, B07Y, %B07Y%
LogEntry("Save Calls Button coordinates logged: " . B07X . "," . B07Y)
Return

B8:		; Loc Input
KeyWait, Enter, D
MouseGetPos, B08X, B08Y
MsgBox,,Coords, %B08X% %B08Y%
GuiControl,, B08X, %B08X%
GuiControl,, B08Y, %B08Y%
LogEntry("Loc Input in Save calls tab coordinates logged: " . B08X . "," . B08Y)
Return

B9:		; WAV Radio
KeyWait, Enter, D
MouseGetPos, B09X, B09Y
MsgBox,,Coords, %B09X% %B09Y%
GuiControl,, B09X, %B09X%
GuiControl,, B09Y, %B09Y%
LogEntry("WAV radio button coordinates logged: " . B09X . "," . B09Y)
Return

B10:	; Save Btn
KeyWait, Enter, D
MouseGetPos, B10X, B10Y
MsgBox,,Coords, %B10X% %B10Y%
GuiControl,, B10X, %B10X%
GuiControl,, B10Y, %B10Y%
LogEntry("Save button coordinates logged: " . B10X . "," . B10Y)
Return

B11:	; 3dots
KeyWait, Enter, D
MouseGetPos, B11X, B11Y
MsgBox,,Coords, %B11X% %B11Y%
GuiControl,, B11X, %B11X%
GuiControl,, B11Y, %B11Y%
LogEntry("Save button coordinates logged: " . B11X . "," . B11Y)
Return

B12:	; Save btn
KeyWait, Enter, D
MouseGetPos, B12X, B12Y
MsgBox,,Coords, %B12X% %B12Y%
GuiControl,, B12X, %B12X%
GuiControl,, B12Y, %B12Y%
LogEntry("Save button coordinates logged: " . B12X . "," . B12Y)
Return

B1P:
WinActivate, %Title%
MouseMove, %B01X%, %B01Y%
MouseClick, R, %B01X%, %B01Y%
Return

RunIDPrep:
GuiControl, Text, RunIDPrep, RunningPrep
Global globIDArray := []
LogEntry("Run ID Prep Start (from file:) " . SelectedFile)
If !FileExist(SelectedFile)
{
	LogEntry("Missing " . SelectedFile . " file path...")
	MsgBox,, File Error, Missing %SelectedFile% file path...
	GuiControl, Text, RunIDPrep, RunIDPrep
	Return
}
LogEntry("Count number of lines")
FileRead File, %SelectedFile%
StringReplace File, File, `n, `n, All UseErrorLevel
TotalLines := ErrorLevel
TotalLines++	; Add one as it was always one short since the last line did not have a `n at the end... Of course if there is an emtpy line it will count the extra emtpy line.  Oh well.
LogEntry("Total number of SegID's is: " . TotalLines)
GuiControl, +Range0-%TotalLines%, MyProgress
Gui, Show
Loop, read, %SelectedFile% ; Loading ID's from file
{
	GuiControl, , MyProgress, %A_Index% 
	globIDArray[A_Index] := A_LoopReadLine
	Sleep, 10
	GuiControl, Text, LoadingTxt, Working on %A_Index% of %TotalLines%
}
GuiControl, Text, LoadingTxt, Array Complete: %TotalLines%
LogEntry("Loaded CompleteID's from " . SelectedFile)
Global globNewIDArray :=
LogEntry("Before FuncLoop")
GuiControl, Text, LoadingTxt, Concatenation...
globNewIDArray := FuncLoop(globIDArray)
LogEntry("After FuncLoop - globNewIDArray")
GuiControl, Text, LoadingTxt, Concatenation...Complete
GuiControl, Text, RunIDPrep, Run ID Prep
Return

RunRobot:
Gui, Submit, NoHide
GuiControl, Text, RunRobot, RunningRobot
If (globIDArray[1] = "")
{
	MsgBox,, Robot Error, Run Robot failed - did you prep the ID's first?
	LogEntry("FAILURE - Array slot 1 empty...Returning")
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
TotalIDs := globIDArray.Length()
TotalArray := globNewIDArray.Length()
GuiControl, +Range0-%TotalArray%, MyProgress
GuiControl, , MyProgress, 0
GuiControl, Text, LoadingTxt, Robot Starting: 0 of %TotalArray%
LogEntry("Robot Starting...0 of " . TotalArray)
Loop, % globNewIDArray.Length()
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
			GuiControl, Text, LoadingTxt, Robot Ended. %A_Index% of %TotalArray%
			Return
		}
	}
	LoopTime := A_TickCount
	LoopElapsed := 0
	RightClick :=
	GroupByBoxColor :=
	LogEntry("Waiting for '" . Title . "' to become active")
	TrayTip, Waiting..., Waiting for %Title%, 3, 1
	WinWaitActive, %Title%,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for " . Title)
		MsgBox,, Timeout, Timeout, please focus the %Title% window
		TrayTip, Waiting..., Waiting for %Title%, 3, 1
		WinWaitActive, %Title%,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
	MouseClick, R, %B01X%, %B01Y%  ; Right click on query
	LogEntry("Right click on query at (" . B01X . "`," . B01Y . ")")
	TrayTip, Query, Right-Click on Query, 3, 1
	LogEntry("Before While loop - waiting for menu")
	StartTime := A_TickCount
	ElapsedTime := 0
	ErrorTimeout := 0
	;While (RightClick != 0x979797 && RightClick != 0xCCCCCC) ; Wait for menu to show
	;{
		;If (ElapsedTime > Timeoutms)
		;{
			;LogEntry("Elapsed time (" . ElapsedTime . ") > Timeoutms (" . Timeoutms . ") , breaking loop")
			;ErrorTimeout := 1
			;break
		;}
		;PixelGetColor, RightClick, %B01X%, %B01Y%
		;LogEntry("Got Color " . RightClick . " at coodinate (" . B01X . "," . B01Y . ")")
		;ElapsedTime := A_TickCount - StartTime
		;LogEntry("Time elapsed: " . ElapsedTime)
	;}
	;If (ErrorTimeout = 1)
	;{
		;MsgBox,, Timeout, Please edit query manually, timer breaking loop required to move along.
		;GoTo SkipClickEdit
	;}
	Sleep, 500
	Loop 5
	{
		Sleep 30
		Send {Down}
	}
	Send {Enter}
	SkipClickEdit:
	LogEntry("Waiting for Advanced Query window")
	TrayTip, Waiting..., Waiting for Advanced Query Window, 3, 1
	WinWait, Advanced Query,,%Timeout%
	WinActivate, Advanced Query
	WinWaitActive, Advanced Query,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Advanced Query Window")
		MsgBox,, Timeout, Timeout waiting for Advanced Query Window, was edit pressed?  Manually edit query and press OK.
		TrayTip, Waiting..., Waiting for Advanced Query Window, 3, 1
		WinWaitActive, Advanced Query,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
	MouseClick,, %B03X%, %B03Y% ; Click on SegID box
	LogEntry("Click on SegID Box")
	MouseGetPos,,,,SegIDControl
	ControlSetText,%SegIDControl%, % globNewIDArray[A_Index], Advanced Query
	TrayTip, Insert, Inserting SegID's, 3, 1
	LogEntry("Inserting " . globNewIDArray[A_Index])
	LogEntry("Click on 'Save&Run'")
	MouseClick,, %B04X%, %B04Y%	; Click on 'Save&Run'
	LogEntry("Waiting for '" . Title . "' to become active")
	TrayTip, Waiting..., Waiting for %Title%, 3, 1
	WinWait, %Title%,,%Timeout%
	WinActivate, %Title%
	WinWaitActive, %Title%,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for " . Title)
		MsgBox,, Timeout, Timeout waiting for %Title%, please focus the %Title% (was 'Save & Run' pressed?)
		TrayTip, Waiting..., Waiting for %Title%, 3, 1
		WinWaitActive, %Title%,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
	StartTime := A_TickCount
	ElapsedTime := 0
	ErrorTimeout := 0
	TrayTip, Waiting..., Waiting for query to finish, 3, 1
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
	LogEntry("Clicking on top result (" . B06X . "`," . B06Y . ")")
	MouseClick,, %B06X%, %B06Y%
	LogEntry("Sending ctrl-a")
	Send ^a
	LogEntry("Click Save Calls (" . B07X . "`," . B07Y . ")")
	MouseClick,, %B07X%, %B07Y%
	LogEntry("Wait for either Save Calls window OR the 'No Recordings Found' window")
	WinWait, ahk_group WindowGroup,, %Timeout%
	LogEntry("Check which window we have...")
	If WinActive("","styledButton4")
	{
		LogEntry("In IF statement for info box - means there are no recordings")
		WinActivate,, Information
		Send, {Esc}
		GuiControl,, MyProgress, %A_Index%
		GuiControl, Text, LoadingTxt, Robot Completed: %A_Index% of %TotalArray%
		continue
	}
	LogEntry("Wait for SaveCalls dialog")
	TrayTip, Waiting..., Waiting for Save Calls Dialog box..., 3, 1
	WinWait, Save Calls,,%Timeout%
	Sleep, 200
	WinActivate, Save Calls
	WinWaitActive, Save Calls,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Save Calls Dialog")
		MsgBox,, Timeout, Timeout waiting for Save Calls Dialog - ensure calls are selected and press 'save calls' button, then press OK.
		TrayTip, Waiting..., Waiting for Save Calls Dialog box..., 3, 1
		WinWaitActive, Save Calls,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
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
		MouseClick,, %B10X%,%B10Y%
		LogEntry("Click on Save Btn")
		TrayTip, Click, Clicked Save, 3, 1
		MouseMove, %B08X%, %B08Y%
		LogEntry("Move mouse to Location input, checking for 'info' screen...")
		MouseGetPos,,,,InfoControl
		Sleep, 500
		InfoBox := WinExist("ahk_exe PresentationHost.exe","styledButton4")
		LogEntry("InfoBox: " . InfoBox . " should only populate when there is the 'info' window, in which case we need to hit esc")
		If (InfoBox != "0x0")
		{
			LogEntry("In IF statement for InfoBox")
			WinActivate,, Information
			Send, {Esc}
		}
		LogEntry("Waiting for 'Saving' Dialog box")
		TrayTip, Waiting..., Waiting for Save Calls to show..., 3, 1
		WinWait, Saving,,%Timeout%
		WinActivate, Saving
		WinWaitActive, Saving,,%Timeout%
		If ErrorLevel
		{
			LogEntry("Timeout waiting for Saving Dialog")
			MsgBox,, Timeout, Timeout waiting for Saving Dialog - was 'Save Calls' pressed?
			TrayTip, Waiting..., Waiting for Save Calls to show..., 3, 1
			WinWaitActive, Saving,,%Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
		}
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
		Sleep, SaveTimeoutms
		Send, {Tab}
		LogEntry("Tab once")
		Send, {Tab}
		LogEntry("Tab twice")
		Send, {Space}
		LogEntry("Space to close the window...")
		WaitAppEnd:
		TrayTip, Waiting..., Wait for %Title%, 3, 1
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
		WinWaitActive, Save as,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout, Timeout waiting for 'save as' dialog
			TrayTip, Waiting..., Waiting for Save As dialog..., 3, 1
			WinWaitActive, Save as,, %Timeout%
			If ErrorLevel
			{
				MsgBox,, Timeout 2, Second timeout, please start the robot again.
				GuiControl, Text, RunRobot, RunRobot
				Return
			}
		}
		WinActivate, Save as
		UNCPathNIM := UNCPath "\Call1.wav"
		LogEntry("Injecting path: " . UNCPathNIM)
		ControlSetText, Edit1, %UNCPathNIM%
		Sleep, 300
		Send, {Enter}
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
		LogEntry("Clicking on Close btn at: (" . B12X . "," . B12Y . ")")
		MouseClick,, %B12X%, %B12Y%
	}
	EndStep:
	LoopElapsed := A_TickCount - LoopTime
	LoopElapsed := LoopElapsed / 1000
	LoopElapsed := Round(LoopElapsed)
	LogEntry("Completed loop: " . A_Index . " of " . globNewIDArray.Length() . " in " . LoopElapsed . "sec.")
	TrayTip, Loop End, End of loop %A_Index% of %TotalArray%, 3, 1
	GuiControl,, MyProgress, %A_Index%
	GuiControl, Text, LoadingTxt, Robot Completed: %A_Index% of %TotalArray%
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
	LogEntry("In FuncLoop - Setting up vars")
	NewIDArray := []
	IDArray := []
	ID :=
	Increment := 1
	TotalArray := CompleteIDArray.length()
	for k, v in CompleteIDArray
	{
		NewIDArray.Push(v)
		GuiControl, , MyProgress, %A_Index% 
		GuiControl, Text, LoadingTxt, Concatenation...%k% of %TotalArray%
		;check if we are at the 20th item OR if we are at the last item (which could be a weird number)
		if (NewIDArray.length()=20 || a_index = CompleteIDArray.length())
		{
			LogEntry("Entered IF statement (more than 20 items or end of list)")
			for k2, v2 in NewIDArray
			{
				ID := ID NewIDArray[k2] ";"
			}
			LogEntry("Create " . ID)
			IDArray[Increment] := ID
			LogEntry("Add to IDArray at index: " . Increment)
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