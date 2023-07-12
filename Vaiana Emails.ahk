#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance force

;===================================================================================================
;; CTRL + ALT + M
^!m::

Gui, Add, Button, xm+5  w110 gHacking, % "Hacking"
Gui, Add, Button, xm+5 y+5 w110 gUeber, % "Überziehung"
Gui, +ToolWindow ;; Remove Min and Max Buttons in Window
Gui, Show,,Vaiana E-Mails
GuiControl, Focus, Hacking
; Gui, Destroy
return

;;--------------------------
;; Hacking E-Mail

Hacking:
	RAC=""
	Gui, Destroy
	Gui, Font, bold
	Gui, Add, Text, ,E-Mail Optionen:
	Gui, Font, norm
	Gui, Add, DropDownList, w200 vvarEmail, CRW-CF.Telefonie-Early@db.com||crw-cf.telefonie-mid@db.com
	Gui, Add, CheckBox, vOptAttachment, Datei hinzufügen?

	Gui, Font, bold
	Gui, Add, Text, ,Vorfall Daten:
	Gui, Font, norm
	Gui, Add, Text, ,Router Nr:
	Gui, Add, Edit, w200 vRAC, 
	Gui, Add, Text, ,Name des KN:
	Gui, Add, Edit, w200 vName, 
	Gui, Add, Text, ,Konto gehackt am:
	Gui, Add, MonthCal, w400 h200 vHackDatum

	Gui, Add, Text, ,Anzeige? Ja/Nein
	Gui, Add, DropDownList, w200 vAnzeige, Ja||Nein
	Gui, Add, Text, ,AZ der Anzeige:
	Gui, Add, Edit, w200 vAzAnzeige,
	Gui, Add, Text, ,Weitere Infos zum Vorfall:
	Gui, Add, Edit, w400 h200 vInfos,
	Gui, Add, Button, w50 x175 gOK, OK
	Gui, +ToolWindow ;; Remove Min and Max Buttons in Window
	Gui, Show, Center Autosize, Vaiana-Hacking
	Return

	OK:
	Gui, Submit
	Gui, Destroy

	;; Format the dates for the Email
	TodayDate := A_DDDD . ", " . A_DD . "." . A_MM . "." . A_YYYY
	FormatTime, HackDatum2, %HackDatum%, dd/MM/yyyy

	;; Open Outlook if not already open
	try
		outlookApp := ComObjActive("Outlook.Application")
	catch
		outlookApp := ComObjCreate("Outlook.Application")

	;; Start a new Email
	olMailItem := 0 
	MailItem := outlookApp.CreateItem(olMailItem) 
	MailItem.Display
	;; Delete signature 
	Sleep, 500
	
	Send, %A_Tab%
	Send, %A_Tab%
	Send, %A_Tab%
	SEND, ^a
	Send, {delete}

	Sleep, 500
	;; Fill the recipients
	MailItem.TO :=varEmail

	;; Subject of the E-Mail
	subj= Betrugsfall Router Nr: %RAC% 
	MailItem.Subject := subj ;Betrugsfall RouterNr ;%RAC%

	;; Body of the E-mail
	MailItem.BodyFormat := 2 ;olFormatHTML 

	HTMLBody =
	
	<p>Hallo,</p>
	<p>KN teilt mit, dass das Konto gehackt wurde.</p>
	<ul>
	<li><em>Router Nr.</em>: %RAC%</li>
	<li><em>Name des KN / BVM</em>: %Name%</li>
	<li><em>Konto gehackt am</em>: %HackDatum2%</li>
	<li><em>Anzeige Ja/Nein?</em> %Anzeige%</li>
	<li><em>AZ der Anzeige</em>: %AzAnzeige%</li>
	<li><em>Weitere Infos zum Vorfall</em>: </li>
	<p>%Infos%</p>
	</ul>
	<p>Wir bitten um Pr&uuml;fung. Vielen Dank</p>
	<p>Mit freundlichen Gr&uuml;&szlig;e</p>
	<p><strong>Das DB Team</strong></p>
	<p><strong><em><br /></em></strong>Telefon: +49 6252 672 209</p>
	<p>Intrum Deutschland GmbH<br />Donnersbergstra&szlig;e 1<br />64646 Heppenheim<br />Germany<br />Zentrale: +49 6252 672 0 <br />Fax: +49 6252 672 230</p>
	<p><strong><a href="https://www.intrum.com/de/de/unsere_loesungen/">intrum.de</a></strong></p>
	<p>Sitz der Gesellschaft: Heppenheim</p>
	<p>Gesch&auml;ftsf&uuml;hrer: Marc Knothe, Yvonne Wagner</p>
	<p>Handelsregister Darmstadt HRB 87484</p>
	<p>USt-IdNr. DE177092311<br />Registrierter Inkassodienstleister nach &sect; 10 Abs. 1 Nr. 1 RDG</p>
	<p>This e-mail and any attachments are confidential and may<br />also be privileged. If you are not the named recipient, please<br />notify the sender immediately and do not disclose the contents<br />to another person, use it for any purpose, or store or copy the<br />information in any medium. Thank you for your cooperation.</p>
	<p>Please consider the environment before printing this e-mail</p>
	<p><a href="https://www.intrum.de/datenschutz">Information about how we process personal data</a></p>

	MailItem.HTMLBody := HTMLBody

	;; Attachments (if defined in the previous menu)
	if (OptAttachment=1){
		FileSelectFile, var_attachment,,,Datei auswählen
		MailItem.Attachments.Add(var_attachment) 
	}

	;; Show the E-Mail ready to be sent
	MailItem.Display ;Make email visible 

	Gui, Destroy
	task=Vaiana Hacking
	gosub track
	Return

;;--------------------------
;; Überziehungsemail

Ueber:
	RAC=""
	Gui, Destroy
	Gui, Font, bold
	Gui, Add, Text, ,E-Mail Optionen:
	Gui, Font, norm
	Gui, Add, DropDownList, w200 vvarEmail, CRW-CF.Telefonie-Early@db.com||crw-cf.telefonie-mid@db.com

	Gui, Font, bold
	Gui, Add, Text, ,Vorfall Daten:
	Gui, Font, norm
	Gui, Add, Text, ,Router Nr:
	Gui, Add, Edit, w200 vRAC, 
	Gui, Add, Text, ,KI:
	Gui, Add, Edit, w200 vKi, 
	Gui, Add, Text, ,Geburtstagsdatum:
	Gui, Add, Edit, w200 vGeb, 
	Gui, Add, Text, ,Rückstand:
	Gui, Add, Edit, w200 vRueckstand, 
	Gui, Add, Text, ,Saldo:
	Gui, Add, Edit, w200 vSaldo, 
	Gui, Add, Text, ,Limit:
	Gui, Add, Edit, w200 vLimit, 
	Gui, Add, Button, w50 x75 gOK2, OK
	GuiControl, Focus, RAC
	Gui, +ToolWindow ;; Remove Min and Max Buttons in Window
	Gui, Show, Center Autosize, Vaiana-Überziehung
	Return

	OK2:
	Gui, Submit
	Gui, Destroy

	;; Format the dates for the Email
	TodayDate := A_DDDD . ", " . A_DD . "." . A_MM . "." . A_YYYY

	;; Open Outlook if not already open
	try
		outlookApp := ComObjActive("Outlook.Application")
	catch
		outlookApp := ComObjCreate("Outlook.Application")

	;; Start a new Email
	olMailItem := 0 
	MailItem := outlookApp.CreateItem(olMailItem) 
	MailItem.Display
	;; Delete signature 
	Sleep, 500
	
	Send, %A_Tab%
	Send, %A_Tab%
	Send, %A_Tab%
	SEND, ^a
	Send, {delete}

	Sleep, 500
	;; Fill the recipients
	MailItem.TO :=varEmail

	;; Subject of the E-Mail

	subj= Überziehung >1000€ Router Nr: %RAC% 
	MailItem.Subject := subj ;Überziehung >1000€, RouterNr %RAC% 

	;; Body of the E-mail
	MailItem.BodyFormat := 2 ;olFormatHTML 

	HTMLBody =
	
	<p>Hallo,</p>
	<p>bitte prüfen, ob eine Kreditkarte existiert und gegeben falls bitte die Karte sperren.</p>
	<ul>
	<li><em>Router Nr.</em>: %RAC%</li>
	<li><em>KI</em>: %Ki%</li>
	<li><em>Geburtstagsdatum</em>: %Geb%</li>
	<li><em>Rückstand</em>: %Rueckstand%</li>
	<li><em>Saldo</em>: %Saldo%</li>
	<li><em>Limit</em>: %Limit%</li>
	</ul>
	<p>Vielen Dank</p>
	<p>Mit freundlichen Gr&uuml;&szlig;e</p>
	<p><strong>Das DB Team</strong></p>
	<p><strong><em><br /></em></strong>Telefon: +49 6252 672 209</p>
	<p>Intrum Deutschland GmbH<br />Donnersbergstra&szlig;e 1<br />64646 Heppenheim<br />Germany<br />Zentrale: +49 6252 672 0 <br />Fax: +49 6252 672 230</p>
	<p><strong><a href="https://www.intrum.com/de/de/unsere_loesungen/">intrum.de</a></strong></p>
	<p>Sitz der Gesellschaft: Heppenheim</p>
	<p>Gesch&auml;ftsf&uuml;hrer: Marc Knothe, Yvonne Wagner</p>
	<p>Handelsregister Darmstadt HRB 87484</p>
	<p>USt-IdNr. DE177092311<br />Registrierter Inkassodienstleister nach &sect; 10 Abs. 1 Nr. 1 RDG</p>
	<p>This e-mail and any attachments are confidential and may<br />also be privileged. If you are not the named recipient, please<br />notify the sender immediately and do not disclose the contents<br />to another person, use it for any purpose, or store or copy the<br />information in any medium. Thank you for your cooperation.</p>
	<p>Please consider the environment before printing this e-mail</p>
	<p><a href="https://www.intrum.de/datenschutz">Information about how we process personal data</a></p>

	MailItem.HTMLBody := HTMLBody

	;; Show the E-Mail ready to be sent
	MailItem.Display ;Make email visible 

	Gui, Destroy
	task=Vaiana Ueberziehung
	gosub track
	Return
;--------------------------

GuiEscape:
Gui, Destroy

track: ;calculates how many times AHK is pressed and gone to the end

	date:=Substr(A_Now, 1,8)
	time:=Substr(A_Now, 9,14)
	
	path= \\groupad1\data\Germany\EDV Bereitstellung\Robotics\logs\vaiana_kurzeln.csv
	FileAppend, %date%;%time%;%A_ScriptName%;%A_Username%;%task%`n, %path%
return