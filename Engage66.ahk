#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Persistent	; Keep the script running until the user exits it
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
SetTitleMatchMode, RegEx
CoordMode, Mouse, Screen
CoordMode, Pixel, Screen

SelectedFile := "CompleteID.txt"
ConfigFile := "EKconfig.txt"

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
Gui, add, Edit, w300 h100 r1 x15 y125 vUNCpath, C:\temp
Gui, -dpiscale

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
	}
	Else
	{
		Menu, %MenuName%, UnCheck, %MenuItem%
		Gui, 1: -AlwaysOnTop
	}
}

SaveConfig:
Loop, 10
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
IfExist EKConfig.txt
	MsgBox,4,Delete Config, Do you want to overwrite the existing config?
IfMsgBox Yes
	FileDelete, EKConfig.txt
IfMsgBox No
	Return
Loop, 10
{
	If (A_Index < 10)
	{
		tmpX := B0%A_Index%X
		tmpY := B0%A_Index%Y
		Sleep, 50
		FileAppend, % "B0" . A_Index . "X:" . tmpX . "`n", EKConfig.txt
		Sleep, 100
		FileAppend, % "B0" . A_Index . "Y:" . tmpY . "`n", EKConfig.txt
	}
	If (A_Index >= 10)
	{
		tmpX := B%A_Index%X
		tmpY := B%A_Index%Y
		Sleep, 50
		FileAppend, % "B" . A_Index . "X:" . tmpX . "`n", EKConfig.txt
		Sleep, 100
		FileAppend, % "B" . A_Index . "Y:" . tmpY . "`n", EKConfig.txt
	}
}
Return

B1:		; Query
KeyWait, Enter, D
MouseGetPos, B01X, B01Y
MsgBox, %B01X% %B01Y%
GuiControl,, B01X, %B01X%
GuiControl,, B01Y, %B01Y%
Return

B2:		; Edit
KeyWait, Enter, D
MouseGetPos, B02X, B02Y
MsgBox, %B02X% %B02Y%
GuiControl,, B02X, %B02X%
GuiControl,, B02Y, %B02Y%
Return

B3:		; SegID
KeyWait, Enter, D
MouseGetPos, B03X, B03Y
MsgBox, %B03X% %B03Y%
GuiControl,, B03X, %B03X%
GuiControl,, B03Y, %B03Y%
Return

B4:		; SaveRun
KeyWait, Enter, D
MouseGetPos, B04X, B04Y
MsgBox, %B04X% %B04Y%
GuiControl,, B04X, %B04X%
GuiControl,, B04Y, %B04Y%
Return

B5:		; GroupBy
KeyWait, Enter, D
MouseGetPos, B05X, B05Y
MsgBox, %B05X% %B05Y%
GuiControl,, B05X, %B05X%
GuiControl,, B05Y, %B05Y%
Return


B6:		; Top Result
KeyWait, Enter, D
MouseGetPos, B06X, B06Y
MsgBox, %B06X% %B06Y%
GuiControl,, B06X, %B06X%
GuiControl,, B06Y, %B06Y%
Return

B7:		; Save Calls
KeyWait, Enter, D
MouseGetPos, B07X, B07Y
MsgBox, %B07X% %B07Y%
GuiControl,, B07X, %B07X%
GuiControl,, B07Y, %B07Y%
Return

B8:		; Loc Input
KeyWait, Enter, D
MouseGetPos, B08X, B08Y
MsgBox, %B08X% %B08Y%
GuiControl,, B08X, %B08X%
GuiControl,, B08Y, %B08Y%
Return

B9:		; WAV Radio
KeyWait, Enter, D
MouseGetPos, B09X, B09Y
MsgBox, %B09X% %B09Y%
GuiControl,, B09X, %B09X%
GuiControl,, B09Y, %B09Y%
Return

B10:	; Save Btn
KeyWait, Enter, D
MouseGetPos, B10X, B10Y
MsgBox, %B10X% %B10Y%
GuiControl,, B10X, %B10X%
GuiControl,, B10Y, %B10Y%
Return

RunIDPrep:
GuiControl, Text, RunIDPrep, RunningPrep
Global globIDArray := []
LogEntry("Get ID's from file: " . SelectedFile)
IfNotExist, %SelectedFile%
{
	LogEntry("Missing " . SelectedFile . " file path...")
	MsgBox, Missing %SelectedFile% file path...
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
GuiControl, Text, RunRobot, RunningRobot
Loop, 10
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
GuiControlGet, UNCPath
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
Loop, % globNewIDArray.Length()
{
	LoopTime := A_TickCount
	LoopElapsed := 0
	RightClick :=
	GroupByBoxColor :=
	LogEntry("Waiting for 'Application Suite' to become active")
	TrayTip, Waiting..., Waiting for AppSuite, 3, 1
	WinWaitActive, Application Suite,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Application Suite")
		MsgBox,, Timeout, Timeout, please focus the Application Suite
		TrayTip, Waiting..., Waiting for AppSuite, 3, 1
		WinWaitActive, Application Suite,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
	Sleep, 100
	MouseClick, R, %B01X%, %B01Y%  ; Right click on query
	LogEntry("Right click on query at (" . B01X . "`," . B01Y . ")")
	Sleep, 100
	TrayTip, Query, Right-Click on Query, 3, 1
	LogEntry("Before While loop - waiting for menu")
	StartTime := A_TickCount
	ElapsedTime := 0
	ErrorTimeout := 0
	While (RightClick != 0x979797 && RightClick != 0xCCCCCC) ; Wait for menu to show
	{
		If (ElapsedTime > Timeoutms)
		{
			LogEntry("Elapsed time (" . ElapsedTime . ") > Timeoutms (" . Timeoutms . ") , breaking loop")
			ErrorTimeout := 1
			break
		}
		PixelGetColor, RightClick, %B01X%, %B01Y%
		LogEntry("Got Color " . RightClick . " at coodinate (" . B01X . "," . B01Y . ")")
		Sleep, 50
		ElapsedTime := A_TickCount - StartTime
		LogEntry("Time elapsed: " . ElapsedTime)
	}
	If (ErrorTimeout = 1)
	{
		MsgBox,, Timeout, Please edit query manually, timer breaking loop required to move along.
		GoTo SkipClickEdit
	}
	MouseClick,, %B02X%, %B02Y%   ; Click on 'Edit' menu
	LogEntry("Clicked on 'Edit'")
	SkipClickEdit:
	LogEntry("Waiting for Advanced Query window")
	TrayTip, Waiting..., Waiting for Advanced Query Window, 3, 1
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
	Sleep, 100
	MouseClick,, %B03X%, %B03Y% ; Click on SegID box
	LogEntry("Click on SegID Box")
	MouseGetPos,,,,SegIDControl
	ControlSetText,%SegIDControl%, % globNewIDArray[A_Index], Advanced Query
	TrayTip, Insert, Inserting SegID's, 3, 1
	LogEntry("Inserting " . globNewIDArray[A_Index])
	Sleep, 100
	LogEntry("Click on 'Save&Run'")
	MouseClick,, %B04X%, %B04Y%	; Click on 'Save&Run'
	LogEntry("Waiting for 'Application Suite' to become active")
	TrayTip, Waiting..., Waiting for AppSuite, 3, 1
	WinWaitActive, Application Suite,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Application Suite")
		MsgBox,, Timeout, Timeout waiting for Application Suite, please focus the Application Suite (was 'Save & Run' pressed?)
		TrayTip, Waiting..., Waiting for AppSuite, 3, 1
		WinWaitActive, Application Suite,,%Timeout%
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
		Sleep, 100
		ElapsedTime := A_TickCount - StartTime
		LogEntry("Time elapsed: " . ElapsedTime)
	}
	If (ErrorTimeout = 1)
		MsgBox,, Error Timeout, Please ensure query ran successfully, timer breaking loop required to move along.
	LogEntry("Clicking on top result (" . B06X . "`," . B06Y . ")")
	MouseClick,, %B06X%, %B06Y%
	Sleep 300
	LogEntry("Sending ctrl-a")
	Sleep 100
	Send ^a
	Sleep, 300
	LogEntry("Click Save Calls (" . B07X . "`," . B07Y . ")")
	MouseClick,, %B07X%, %B07Y%
	LogEntry("Wait for SaveCalls dialog")
	TrayTip, Waiting..., Waiting for Save Calls Dialog box..., 3, 1
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
	LogEntry("Click 'location' inputbox (" . B08X . "`," . B08Y . ")")
	MouseClick,, %B08X%, %B08Y%
	Sleep 100
	MouseGetPos,,,,LocControl
	LogEntry("ClassNN - " . LocControl)
	ControlSetText, %LocControl%, %UNCPath%
	LogEntry("Inject save calls path (" . UNCPath . ") to location input box")
	MouseClick,, %B09X%,%B09Y%
	LogEntry("Click on WAV radio btn")
	Sleep 100
	MouseClick,, %B10X%,%B10Y%
	LogEntry("Click on Save Btn")
	TrayTip, Click, Clicked Save, 3, 1
	Sleep, 500
	MouseMove, %B08X%, %B08Y%
	LogEntry("Move mouse to Location input, checking for 'info' screen...")
	MouseGetPos,,,,InfoControl
	Sleep, 100
	LogEntry("Getting control under mouse")
	InfoBox :=
	ControlGet, InfoBox,Hwnd,,,, styledButton4
	Sleep, 100
	LogEntry("InfoBox: " . %InfoBox% . " should only populate when there is the 'info' window, in which case we need to hit esc")
	If (InfoBox != "")
	{
		LogEntry("In IF statement for InfoBox")
		TrayTip, Esc, Esc window for Screen Calls, 3, 1
		Send, {Esc}
	}
	LogEntry("Waiting for 'Saving' Dialog box")
	TrayTip, Waiting..., Waiting for Save Calls to show..., 3, 1
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
	LogEntry("Waiting for 'Close' button on 'Saving' dialog (indicates task is complete)")
	TrayTip, Waiting..., Waiting for Save Calls to Complete..., 5, 1
	WinWaitActive,, Close,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Close Button on 'Save Calls'")
		MsgBox,, Timeout, Timeout waiting for 'Close' button after save calls - press OK after it is complete
		WinWaitActive,, Close,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
	Sleep, SaveTimeoutms
	Send, {Tab}
	Sleep, 50
	LogEntry("Tab once")
	Send, {Tab}
	Sleep, 50
	LogEntry("Tab twice")
	Send, {Space}
	LogEntry("Space to close the window...")
	TrayTip, Waiting..., Wait for AppSuite, 3, 1
	WinWaitActive, Application Suite,,%Timeout%
	If ErrorLevel
	{
		LogEntry("Timeout waiting for Application Suite")
		MsgBox,, Timeout, Please focus the Application Suite
		WinWaitActive, Application Suite,,%Timeout%
		If ErrorLevel
		{
			MsgBox,, Timeout 2, Second timeout, please start the robot again when ready.
			GuiControl, Text, RunRobot, RunRobot
			Return
		}
	}
	LoopElapsed := A_TickCount - LoopTime
	LoopElapsed := LoopElapsed / 1000
	LoopElapsed := Round(LoopElapsed)
	LogEntry("Completed loop: " . A_Index . " of " . globNewIDArray.Length() . " in " . LoopElapsed . "sec.")
	TrayTip, Loop End, End of loop %A_Index% of %TotalArray%, 3, 1
	GuiControl, , MyProgress, %A_Index%
	GuiControl, Text, LoadingTxt, Robot Completed: %A_Index% of %TotalArray%
}
TotalElapsed := A_TickCount - TotalTime
TotalElapsed := TotalElapsed / 1000
TotalElapsed := Round(TotalElapsed)
LogEntry("Task Complete in " . TotalElapsed . "sec.  Go have a beer.")
MsgBox, Done in %TotalElapsed%sec
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
			Sleep, 50
			LogEntry("Create " . ID)
			IDArray[Increment] := ID
			Sleep, 50
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
	FileAppend, %A_MM%/%A_DD%/%A_YYYY% %A_Hour%:%A_Min%:%A_Sec%.%A_MSec% - %Message%`n, Log.txt
}

ExitSub:
LogEntry("Exit Application")
GuiClose:
GuiEscape:
ExitApp
Exit