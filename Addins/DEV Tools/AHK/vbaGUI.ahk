/*	Info
AUTHOR: 
	Anastasiou Alex
		anastasioualex@gmail.com
		https://github.com/alexofrhodes <-- Repos
		https://alexofrhodes.github.io	<-- Blog
		https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
		
PURPOSE: GUI to run excel macros
*/


;#IfWinActive, ahk_exe EXCEL.EXE			;if any excel window is active
;#IfWinActive ahk_class wndclass_desked_gsk ;if vbeditor window is active


#SingleInstance, force
SetWorkingDir %A_ScriptDir%
#include ini-editor.ahk

#Include Class_ImageButton.ahk
EStyle := [[0, 0x80F0F0F0, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2] ; normal
		, [0, 0x80C6E9F4, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2] ; hover
		, [0, 0x8086D0E7, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2] ; pressed
		, [0, 0x80F0F0F0, , , 8, 0xFFF0F0F0, 0x8046B8DA, 2]]


;Custom Tray Icon
I_Icon = %A_WorkingDir%\vbaGUI.ico ;dAKirby309 (Michael) at https://icon-icons.com/icon/excel-mac/23559
IfExist, %I_Icon%
	Menu, Tray, Icon, %I_Icon%

;Event for Tray icon left click
OnMessage(0x404, "AHK_NOTIFYICON")
AHK_NOTIFYICON(wParam, lParam)
{
    if (lParam = 0x201) ; WM_LBUTTONDOWN
    {
        gosub start
        return 0
    }
}

;Load Configuration
LoadOptions:
	IniRead, myHotkey, config.ini, Settings, myHotkey
	; IniRead, WorkbookName, config.ini, Settings, WorkbookName    ;moved to runExcelMacro
	IniRead, MenuFile, config.ini, Settings, MenuFile
	IniRead, ItemsPerColumn, config.ini, Settings, ItemsPerColumn
	Hotkey, %myHotkey%,Start
	IniRead, fontSize, config.ini, Settings, fontSize
	Gui,Font, s%fontSize%, Arial
	
	IniRead, xPos, config.ini, Settings, xPos, 100
	IniRead, yPos, config.ini, Settings, yPos, 100
	
return ;if you comment this out then the gui will show at startup (hotkey still works)

Start:

;get menu headers for dropdown menu
GetGuis:
{	
	Guis:=

	;put text file's content into a variable
	FileRead, content, %MenuFile%

	;Split into an aray of its lines
	for each, line in StrSplit(content, "`n", "`r")
	{
		if trim(line) =
			continue			;<--- means to skip this element of the loop
		
		FirstCharacter:= substr(line,1,1)	
		if firstCharacter in !,-,.,;
			Continue
		
		;Choose the non indented elements as headers for the menu
		if not (firstCharacter = A_Tab) 
		{
			if Guis=
			{
				Guis:= Guis . line
			}else{
				Guis:= Guis . "|" . line
			}
		}
	}
}

;Setup GUI with a dropdown to allow switching
ChooseGui:
{
	
	if (ChooseGui="") 
		ChooseGui:=0

	gosub SavePos

	Gui, Submit
	Gui, Destroy

	/*	
		Add a dropdown to the GUI
		The selection change will create the corresponding menu
		AltSubmit passes the element's index instead of text to the variable fo the control
	*/

	Gui,Font, s%fontSize%, Arial

	Gui, Add, DropDownList, section choose%ChooseGui% AltSubmit gChooseGui vChooseGui, %Guis%
	gui, add, button, xs section gMenuSettings,  Options
	gui, add, Button, ys gEditFile, Menu
	gui, add, button, ys gReloadMe,Reload

	Gui, Add, Text, x5 h0 w250 0x10

	;sections resets the x y position for subsequent controls
	Gui, Add, Text, xs section
}

;Create GUI according to dropdown selection (by index)
LoadMenu:
{
	
	guicontrolget, DDSelection, , ChooseGui, text
	counter:=0
	MenuFound:=0
	for each, line in StrSplit(content, "`n", "`r")
		{
			if (line=DDSelection)
				{
					MenuFound++
					continue
				}	
				
			if (MenuFound=1)
				{

				FirstCharacter:= substr(line,1,1)
				if FirstCharacter in !,.
					continue
				if not (firstCharacter=A_Tab) && (trim(line)<>"")
					Break
				Position:= InStr(line, ";")
				if Position>0
					line:=SubStr(line, 1, Position - 1)
				Position:= InStr(line, "!")		;because ezMenu uses ! key to switch the GoSub
				if Position>0
					line:=SubStr(line, 1, Position - 1)
				line:=trim(line)
				
				;alternatively to force new column leave empty line (resets counter for the rest of the controls in the new column)
				if line =
				{
					Gui, Add, Text,ys
					counter:=0
					continue
				;or if controls placed in active column reached set limit, start new column
				}else if (counter=ItemsPerColumn)
				{
					counter:=0
					Gui, Add, Text,ys 
				}
				
				/*
				Add a button with the caption = the text of the line, in this case 
				the name of a VBA procedure to run with RunExcelMacro 
				if the workbook (WorkbookName) which was set at the top is open and contains said Procedure
				*/
				Gui, Add, Button, gRunExcelMacro hwndBtn, %line%

				ImageButton.Create(Btn, line, EStyle*) ; define o estilo inicial

				counter++
			}			
		}
}

FinalizeGUI:
{
	;Add dividing line and exit button
	Gui, Add, Text, x5 h0 w250 0x10

	;Hotkey Info
	; Gui,Font, s10, Arial
	gui, Add, Text, xs cRed,  Current Hotkey = %myHotkey%
	; Gui,Font, s8, Arial

	;AUTHOR links
	Gui, Add, Link,xs section, 	<a href="https://github.com/alexofrhodes/">							GitHub		</a> 
	Gui, Add, Link,ys, 			<a href="https://alexofrhodes.github.io/">							Blog		</a> 
	Gui, Add, Link,ys, 			<a href="https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg">	YouTube		</a> 

	;Gui options
	Gui, +AlwaysOnTop ;-Border +resize 

	Gui, +hwnd_hwnd
	; Gui, Color, 0x000000
	
	
	gosub LoadOptions
	Gui, Show, x%xpos% y%ypos%
	
	; WinSet, Transparent, 225, % "ahk_id " _hwnd

	return 
}



/*	NOTES
	Gui, Add, Text, x5 y5 w150 0x10  ;Horizontal Line > Etched Gray
	Gui, Add, Text, x5 y5 h150 0x11  ;Vertical Line > Etched Gray
	Gui, Add, Text, x5 y155 w150 h1 0x7  ;Horizontal Line > Black
	Gui, Add, Text, x155 y5 w1 h150 0x7  ;Vertical Line > Black

	Gui, Add, Button, gRunExcelMacro, [&1] abc ; the last part is the Sub name, [&1] and space before sub name act as accelerator, must be unique key
	Gui, Add, Button, gRunExcelMacro, abcx
	Gui, Add, Button, ys gRunExcelMacro, [&3] Place tags  ;ys starts new column
	Gui, Add, Button, gRunExcelMacro, [&4] Open Notepad
*/

RunExcelMacro:
{
	try {
		XL := Excel_Get()
	} catch {
		MsgBox, 16,, Can't obtain Excel! 
		return
	}
	;MsgBox, 64,, Excel obtained successfully!   ;for debugging purposes
	
	IniRead, WorkbookName, config.ini, Settings, WorkbookName

	;a space is allowed in the following format: [&1] MacroName
	;to allow a GUI accelerator between the braces eg. [accelerator]
	if instr(A_GuiControl, A_Space){
		StringSplit, Procedure, A_GuiControl, %A_Space%
		macro:= "'" . WorkbookName . "'" . "!" . Procedure2 
	}else{
		macro:= "'" . WorkbookName . "'" . "!" . A_GuiControl
	}
	
	try {
		XL.Run(macro)  
	} catch {
		MsgBox, 16,, Can't find the macro %A_GuiControl% in %WorkbookName%
	}
	
	return
}

Excel_Get(WinTitle:="ahk_class XLMAIN", Excel7#:=1) 
{
	/*
		Excel_Get by jethrow (modified)
		Forum:    https://autohotkey.com/boards/viewtopic.php?f=6&t=31840
		Github:   https://github.com/ahkon/MS-Office-COM-Basics/blob/master/Examples/Excel/Excel_Get.ahk
		
		References
		https://autohotkey.com/board/topic/88337-ahk-failure-with-excel-get/?p=560328
		https://autohotkey.com/board/topic/76162-excel-com-errors/?p=484371
		https://autohotkey.com/boards/viewtopic.php?p=134048#p134048
	*/
	
	
	static h := DllCall("LoadLibrary", "Str", "oleacc", "Ptr")
	WinGetClass, WinClass, %WinTitle%
	if !(WinClass == "XLMAIN")
		return "Window class mismatch."
	ControlGet, hwnd, hwnd,, Excel7%Excel7#%, %WinTitle%
	if (ErrorLevel)
		return "Error accessing the control hWnd."
	VarSetCapacity(IID_IDispatch, 16)
	NumPut(0x46000000000000C0, NumPut(0x0000000000020400, IID_IDispatch, "Int64"), "Int64")
	if DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", hWnd, "UInt", -16, "Ptr", &IID_IDispatch, "Ptr*", pacc) != 0
		return "Error calling AccessibleObjectFromWindow."
	window := ComObject(9, pacc, 1)
	if ComObjType(window) != 9
		return "Error wrapping the window object."
	Loop
		try return window.Application
	catch e
		if SubStr(e.message, 1, 10) = "0x80010001"
			ControlSend, Excel7%Excel7#%, {Esc}, %WinTitle%
	else
		return "Error accessing the application object."
}


ReloadMe:
	gosub SavePos
	Reload
	Sleep 1000 ; If successful, the reload will close this instance during the Sleep, so the line below will never be reached.
	MsgBox, 4,, The script could not be reloaded. Would you like to open it for editing?
	IfMsgBox, Yes, Edit
return

MenuSettings:
	IniSettingsEditor("vbaGUI", "config.ini")
Return

EditFile:
	run %MenuFile%
return

;Close GUI when exit button pressed or ESC pressed. This doesn't stop the script's execution.
GuiEscape:
	gosub SavePos
	Gui, Destroy
return

GuiClose:
	gosub SavePos
	Gui, Destroy
return

SavePos:
    Gui +lastfound
    WinGetPos, xPos, yPos
	if (xPos <= 0) 
		return
    IniWrite, %xPos%, config.ini, Settings, xPos
    IniWrite, %yPos%, config.ini, Settings, yPos
Return