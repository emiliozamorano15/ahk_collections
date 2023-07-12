; W:\Germany\BI_Operation\00_AR\04_Members\09_EZ\AutoHotkey_1.1.35.00
; ---------------------------------
; Initial Setup
; ---------------------------------

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force
#Hotstring EndChars `n


; ---------------------------------
; Hotstrings
; ---------------------------------

::ALO::
	SendInput KI ist arbeitslos. Derzeit Zahlungsunfähig.
	task=ALO
	gosub track
	Return
	
::KG::
	SendInput Bezieht Krankengeld. Ist derzeit Zahlungsunfähig.
	task=KG
	gosub track
	Return
	
::SB RU::
	SendInput Sprachbarriere. Kunde spricht nur Rumänisch. RR mit Dolmetscher nicht möglich. In die Filiale verwiesen.
	task=SB RU
	gosub track
	Return
	
::SB BG::
	SendInput Sprachbarriere. Kunde spricht nur Bulgarisch. RR mit Dolmetscher nicht möglich. In die Filiale verwiesen.
	task=SB BG
	gosub track
	Return
	
::SB ESP::
	SendInput Sprachbarriere. Kunde spricht nur Spanisch. RR mit Dolmetscher nicht möglich. In die Filiale verwiesen.
	task=SB ESP
	gosub track
	Return
	
::SB RUS::
	SendInput Sprachbarriere. Kunde spricht nur Russisch. RR mit Dolmetscher nicht möglich. In die Filiale verwiesen.
	task=SB RUS
	gosub track
	Return
	
::ZU::
	SendInput KI kann nicht zahlen. Nennt keine weitere Informationen.
	task=ZU
	gosub track
	Return
	
::RZV WTR::
	SendInput Wünscht RZV. Per WTR weitergeleitet.
	task=RZV WTR
	gosub track
	Return
	
::RZV::
	SendInput Wünscht RZV. Mitgeteilt dass RZV geprüft wird und dies eine unverbindliche Anfrage ist.
	task=RZV
	gosub track
	Return
	
::FI::
	SendInput Möchte an Telefon nicht sprechen. Wird sich an seinen Bankberater wenden.
	task=FI
	gosub track
	Return
	
;--------------------------
; Hotstrings with GUI
;--------------------------

::AM3::
	Gui, Destroy
	Gui, Add, DropDownList, w150 vEditContents, Early||Mid
	; Gui, Add, Button, w110 x25 +default gGetContentsAM3
	Gui, Add, Button, w110 x25 +default gGetContentsAM3, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, AM3
	Return

	GetContentsAM3:
		Gui, Submit
		Gui, Destroy
		SendInput  Wünscht RZV. Nach Rücksprache mit %EditContents% RZV abgelehnt. Kunde kann nicht zahlen.
		task=AM3
		gosub track
	Return
	
::TIF::
	Gui, Destroy
	Gui, Add, Edit, w150 h20 -WantReturn vEditContents, Datum
	Gui, Add, Button, w110 x25 +default gGetContentsTiF, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, TIF
	Return

	GetContentsTiF:
		Gui, Submit
		Gui, Destroy
		SendInput Hat einen Termin in der Filiale am %EditContents%
		task=TIF
		gosub track
	Return
	
::SB::
	Gui, Destroy
	Gui, Add, DropDownList, w150 vEditContents, Polnisch||Italienisch|Französich
	Gui, Add, Button, w110 x25 +default gGetContentsSB, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, SB
	Return

	GetContentsSB:
		Gui, Submit
		Gui, Destroy
		SendInput Sprachbarriere. Kunde spricht nur %EditContents%. RR mit Dolmetscher nicht möglich. In die Filiale verwiesen.
		task=SB
		gosub track
	Return
		
::SUB::
	Gui, Destroy
	Gui, Add, Edit, w150 h20 -WantReturn vEditContents, SUB Name
	Gui, Add, Button, w110 x25 +default gGetContentsSUB, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, SUB
	Return

	GetContentsSUB:
		Gui, Submit
		Gui, Destroy
		SendInput Ist bei SUB %EditContents%. Dieser wird sich bei uns melden.
		task=SUB
		gosub track
	Return
	
::NL:: 
	Gui, Destroy
	Gui, Add, Edit, w150 h20 -WantReturn vEditContents, Datum
	Gui, Add, Edit, w150 h20 -WantReturn vEditContents2, Ort
	Gui, Add, Button, w110 x25 +default gGetContentsNL, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, NL
	Return

	GetContentsNL:
		Gui, Submit
		Gui, Destroy
		SendInput Teilt mit, dass Kundin am %EditContents% in %EditContents2% verstorben ist.
		task=NL
		gosub track
	Return
	
::WTR::
	Gui, Destroy
	Gui, Add, DropDownList, w150 vEditContents, Early||Mid
	Gui, Add, Button, w110 x25 +default gGetContentsWTR, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, ADR WTR
	Return

	GetContentsWTR:
		Gui, Submit
		Gui, Destroy
		SendInput Warmtransfer an %EditContents% RZV abgelehnt. Kunde kann nicht zahlen.
		task=WTR
		gosub track
	Return
	
::AWTR::
	Gui, Destroy
	Gui, Add, DropDownList, w150 vEditContents, Early||Mid
	Gui, Add, Button, w110 x25 +default gGetContentsADRWTR, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, AWTR
	Return

	GetContentsADRWTR:
		Gui, Submit
		Gui, Destroy
		SendInput Wegen Adressänderung an %EditContents%, per WTR weitergeleitet.
		task=ADR WTR
		gosub track
	Return
		
::ADR::
	Gui, Destroy
	Gui, Add, DropDownList, vEditContents, Early||Mid
	Gui, Add, Edit, w100 h20 -WantReturn vEditContents2, Adresse
	Gui, Add, Button, w110 x25 +default gGetContentsADR, OK
	Gui, +ToolWindow
	Gui, Show, Center w175, ADR
	Return

	GetContentsADR:
		Gui, Submit
		Gui, Destroy
		SendInput Teilt neue Adresse mit: %EditContents2%. In  %EditContents% niemand erreicht.
		task=ADR
		gosub track
	Return

::HACK::
	SendInput, Gehacktes Konto, ULA bereits in der Filiale {–} Email gesendet
	task=HACK
	gosub track	
Return

;--------------------------

track: ;calculates how many times AHK is pressed and gone to the end

	date:=Substr(A_Now, 1,8)
	time:=Substr(A_Now, 9,14)
	
	path= \\groupad1\data\Germany\EDV Bereitstellung\Robotics\logs\vaiana_kurzeln.csv
	FileAppend, %date%;%time%;%A_ScriptName%;%A_Username%;%task%`n, %path%
return