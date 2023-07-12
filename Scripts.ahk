#SingleInstance Force
; ---------------------------------
; Hotkey Modifier Symbols
; ---------------------------------
; # Windows key
; ! Alt
; ^ Ctrl
; + Shift
; < Left key (in combination with Alt, Ctrl and Shift)
; > Right key (in combination with Alt, Ctrl and Shift)
; <^> AltGr

; ---------------------------------
; General
; --------------------------------- 

#^WheelUp::Volume_Up
#^WheelDown::Volume_Down

;#^PgUp::Send {Volume_Up}
;#^PgDn::Send {Volume_Down}
;#^Pause::Send {Volume_Mute}

; Default windows position
^+0::
	WinGetPos, X, Y, W, H, A
	WinMove, A,,717,31,1200,800
return

; ---------------------------------
; Open new Email in Outlook
; ---------------------------------

;^!m::
;	try
;		outlookApp := ComObjActive("Outlook.Application")
;	catch
;		outlookApp := ComObjCreate("Outlook.Application")
;	MailItem := outlookApp.CreateItem(0)
;	MailItem.Display
; return
 
; ---------------------------------
; Hotstrings
; ---------------------------------
:*:@@::emilio.zamorano@intrum.com

:*:timestamp_::  ; This hotstring replaces "_timestamp" with the current date and time via the commands below.
FormatTime, CurrentDateTime,, dd/MM/yyyy HH:mm ; It will look like 01/09/2005 3:53 PM
SendInput %CurrentDateTime%
return


;-------------------------------------------------------------------------------
;  Open KeePass
;-------------------------------------------------------------------------------
^#k::
Goto, KeePass

;-------------------------------------------------------------------------------
;  Open Worktime Calculator
;-------------------------------------------------------------------------------
^#t::
Goto, Worktime

;-------------------------------------------------------------------------------
;  Press windows+q to quit active window
;-------------------------------------------------------------------------------

#q::Send, !{F4}


;-------------------------------------------------------------------------------
; Open QuickAccess
;------------------------------------------------------------------------------- 

^#-:: ; hotkey Ctrl-Win-F1

Gui, Add, Tab3,, RDPs | Work 
Gui, Add, Button, xm+5 y50  w110 gFi422helsdb92, % "FI422HELSDB92"
Gui, Add, Button, xm+5 y+5 w110 gFi422helsdb139, % "FI422HELSDB139"
Gui, Add, Button, xm+5 y+5 w110 gFi422helstd200, % "FI422HELSTD200"
Gui, Add, Button, xm+5 y+5 w110 gFI422HELSAS297, % "FI422HELSAS297"
Gui, Add, Button, xm+5 y+5 w110 gEBWVSRV6r, % "EBW-VSRV6-Remote"
Gui, Add, Button, xm+5 y+5 w110 gEBWVSRV6, % "EBW-VSRV6"
Gui, Tab, 2
Gui, Add, Button, xm+5 y50  w110 gGFOS, % "GFOS"
Gui, Add, Button, xm+5 y+5  w110 gTeams, % "Teams"
Gui, Add, Button, xm+5 y+5  w110 gKeePass, % "KeePass"
Gui, Add, Button, xm+5 y+5  w110 gAnaconda, % "Anaconda"
Gui, Add, Button, xm+5 y+5  w110 gNotepadpp, % "Notepad++"
Gui, Add, Button, xm+5 y+5  w110 gWorktime, % "WorktimeCalculator"
Gui, Show, AutoSize, % "QuickAccess"
return

;GuiClose:
;ExitApp

Fi422helsdb92:
	Run "C:\Users\zamorem\Documents\RDP\fi422helsdb92.RDP"
	GoTo, GuiClose
return

Fi422helsdb139:
	Run "C:\Users\zamorem\Documents\RDP\fi422helsdb139.RDP"
	GoTo, GuiClose
return

Fi422helstd200:
	Run "C:\Users\zamorem\Documents\RDP\fi422helstd200.RDP"
	GoTo, GuiClose
return

FI422HELSAS297:
	Run "C:\Users\zamorem\Documents\RDP\FI422HELSAS297.RDP"
	GoTo, GuiClose
return

EBWVSRV6r:
	Run "C:\Program Files (x86)\Citrix\ICA Client\SelfServicePlugin\SelfService.exe" -launch -reg "Software\Microsoft\Windows\CurrentVersion\Uninstall\ctx-prod-d-b45aab95@@Kolohonka XD7.15-Prod.Remote Desktop Co-1" -startmenuShortcut
	GoTo, GuiClose
return

EBWVSRV6:
	Run "C:\Users\zamorem\Documents\RDP\IKA3-Reporting.RDP"
	GoTo, GuiClose
return

GFOS:
	Run http://10.1.1.10:8080/gfos/login.xhtml?PERS=0397&MAND=LIN&WERK=2000
	GoTo, GuiClose
return

KeePass:
	Run "C:\Program Files (x86)\KeePass2x\KeePass.exe" 
	GoTo, GuiClose
return 

Anaconda:
	Run %windir%\System32\cmd.exe "/K" C:\Users\zamorem\Software\Anaconda\Scripts\activate.bat C:\Users\zamorem\Software\Anaconda
	GoTo, GuiClose
return 

Teams:
	Run C:\Users\zamorem\AppData\Local\Microsoft\Teams\Update.exe --processStart "Teams.exe"
	GoTo, GuiClose
return

Notepadpp:
	Run C:\Users\zamorem\Software\npp.8.0.portable.x64\notepad++.exe
	GoTo, GuiClose
return

Worktime:
	Run "C:\Users\zamorem\Software\WorktimeCalculator\WorktimeCalculator.exe"
	GoTo, GuiClose
return
	
GuiClose:
Gui, Destroy

GuiEscape:
Gui, Destroy

;ExitApp 
