SetNumLockState, AlwaysOn
CapsLock::End


;Privacy key, Minimizes chrome and pauses media

^#w::
closeChrome()
{
	switchToChrome()
	Send, #{Down}
	Sleep, 100
	Send, {Media_Play_Pause}
}

;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;										bring app to front
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
^F1::
switchToChrome()
{
IfWinNotExist, ahk_exe chrome.exe
	Run, chrome.exe

if WinActive("ahk_exe chrome.exe")
	Sendinput ^{tab}
else
	WinActivate ahk_exe chrome.exe
}
;----------------------------------------------------------------------------------------------------

^F2::
switchToExplorer(){
IfWinNotExist, ahk_class CabinetWClass
	Run, explorer.exe
GroupAdd, taranexplorers, ahk_class CabinetWClass
if WinActive("ahk_exe explorer.exe")
	GroupActivate, taranexplorers, r
else
	WinActivate ahk_class CabinetWClass 
	;you have to use WinActivatebottom if you didn't create a window group.
}
;-----------------------------------------------------------------------------------------------

^F3::
switchWordWindow()
{
 Process, Exist, WINWORD.EXE
 ;msgbox errorLevel `n%errorLevel%
	 If errorLevel = 0
		 Run, WINWORD.EXE
	 else
	 {
	GroupAdd, taranwords, ahk_class OpusApp
	if WinActive("ahk_class OpusApp")
		GroupActivate, taranwords, r
	else
		WinActivate ahk_class OpusApp
	 }
}
;--------------------------------------------------------------------------------------------------
^F4::
switchExcelWindow()
{
 Process, Exist, EXCEL.EXE
 ;msgbox errorLevel `n%errorLevel%
	 If errorLevel = 0
		 Run, EXCEL.EXE
	 else
	 {
	GroupAdd, taranexcel, ahk_class XLMAIN
	if WinActive("ahk_class XLMAIN")
		GroupActivate, taranexcel, r
	else
		WinActivate ahk_class XLMAIN
	 }
}
;------------------------------------------------------------------------------------------------------

^F5::
switchAccutermWindow()
{
IfWinNotExist, ahk_exe Atwin2k2.exe
	Run, Atwin2k2.exe

if WinActive("ahk_exe Atwin2k2.exe")
	Sendinput ^{tab}
else
	WinActivate ahk_exe Atwin2k2.exe
}
;-------------------------------------------------------------------------------------------------------------

^F6::
switchToOutlook()
{
 Process, Exist, OUTLOOK.EXE
 ;msgbox errorLevel `n%errorLevel%
	 If errorLevel = 0
		 Run, OUTLOOK.EXE
	 else
	 {
	GroupAdd, taranmail, ahk_class rctrl_renwnd32
	if WinActive("ahk_class rctrl_renwnd32")
		GroupActivate, taranmail, r
	else
		WinActivate ahk_class rctrl_renwnd32
	 }
}
;----------------------------------------------------------------------------------------------------------------------

^F7::
switchToFirefox()
{
IfWinNotExist, ahk_exe firefox.exe
	Run, firefox.exe
	
if WinActive("ahk_exe firefox.exe")
	Sendinput ^{tab}
else
	WinActivate ahk_exe firefox.exe
}


;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;												Open instance of Vendor Maint or Prod Maint
;++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
^#s::
openProdMnt()
{
Run, H:\`%`%VB_Current151_exe\ProdMaint.exe
Sleep, 500
Send, 4378
Sleep, 200
Send, {Enter}
Sleep, 100
Send, PRODM
Sleep, 100
Send, {Enter}
Send, {Enter}
}

^F8::
switchToProdMaint()
{
 Process, Exist, ProdMaint.exe
 ;msgbox errorLevel `n%errorLevel%
	 If errorLevel = 0
		 openProdMnt()
	 else
	 {
	GroupAdd, prdmnt, ahk_exe ProdMaint.exe
	if WinActive("ahk_exe ProdMaint.exe")
		GroupActivate, prdmnt, r
	else
		WinActivate ahk_exe ProdMaint.exe
	 }
}
;------------------------------------------------------------------------------------------------------------------

^#q::
openVendMnt()
{
Run, H:\`%`%VB_Current151_exe\VendMaint.exe	

Sleep, 1000
Send, 4378
Sleep, 200
Sleep, 200
Send, {Enter}
Sleep, 200
Send, RELYT
Sleep, 200
Send, {tab}
Sleep, 200
Send, {tab}
Send, {tab}
Sleep, 200
Send, {Enter}
}

^F9::
switchToVendorMaint()
{
	Process, Exist, VendMaint.exe
 ;msgbox errorLevel `n%errorLevel%
	 If errorLevel = 0
		 openVendMnt()
	 else
	 {
	GroupAdd, vndrmnt, ahk_exe VendMaint.exe
	if WinActive("ahk_exe VendMaint.exe")
		GroupActivate, vndrmnt, r
	else
		WinActivate ahk_exe VendMaint.exe
	 }
}


;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;												Accuterm SDI Automatic Commands
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

;			Import product number and quantity to PO

+^F1::
CopyPasteNewOrder()
{
InputBox, totalChanges, Enter total changes:

	Loop, %totalChanges%
	{
		copyPastePOAccuterm()
	}
}

copyPastePOAccuterm()
{
;Copies a Cell and pastes both item and quantity in accuterm SDI PO menu, 
;then moves to the next cell to repeat.

	Send, ^c
	Sleep, 200
	switchAccutermWindow()
	
	Sleep, 200
	Send, ^v
	Send, {Enter}
	Send, {Enter}    ;send first copy paste from excel to Accuterm
	Sleep, 200
	
	switchExcelWindow()
	
	Sleep, 200
	Send, {tab}   ;move to next cell and copy qnty
	Send, ^c
	Sleep, 200
	
	switchAccutermWindow()  ;move back to accuterm paste qnty enter through to next line
	
	Sleep, 200
	Send, ^v
	Sleep, 200
	Send, {Enter}   ;Clear though Price and ticket type
	Send, {Enter}
	Send, {Enter}
	Sleep, 200
	
	switchExcelWindow()  ;move back to excel and set up for next copy
	
	Sleep, 200
	Send, {Down}
	Send, {Left}
}

;--------------------------------------------------------------------------------------------------------------------------------

;			Update New costs to existing PO

+^F2::
updateNewCost()
{
InputBox, totalChanges, Enter total changes:

	Loop, %totalChanges%
	{
		CopyPasteNewCost()
	}
}

CopyPasteNewCost()
{
;Copies line # from cell and pastes into Accuterm. Moves back to excel and
; copies new price into accuterm
	
	Send, ^c
	Sleep, 200
	
	switchAccutermWindow()
	
	Sleep, 200
	Send, ^v
	Sleep, 200
	Send, {Enter}
	Send, {Enter}
	Sleep, 50
	Send, {Enter}
	Sleep, 200
	
	switchExcelWindow()
	
	Sleep, 200
	Send, {tab}
	Sleep, 200
	Send, ^c
	Sleep, 100
	
	switchAccutermWindow()
	
	Sleep, 200
	Send, ^v
	Sleep, 200
		
	Send, {Enter}
	Send, {Enter}
	Sleep, 50
	Send, {Enter}
	Sleep, 100
	
	switchExcelWindow()
	
	Sleep, 100
	Send, {Down}
	Send, {Left}	

}

;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
;										Accuterm log in options ***messing around***
;+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

yesterdayDate()
{
	today = %a_now%
	today += -1, days
	FormatTime, today, %today%, MM/dd/yy
	sendInput %today%
}

;Logs onto Dailysls from the prviouse day.
+^a::
dailySLS()
{
	Send, TJC
	Send, {Enter}
	Sleep, 500
	Send, Dailysls
	Sleep, 200
	Send, {Enter}
	Sleep, 500
	Send, Y
	Sleep, 200
	Send, {Enter}
	Sleep, 200
	Send, T
	Send, {Enter}
	Sleep, 200
	
	yesterdayDate()
	
	Sleep, 500
	Send, {Enter}
	Sleep, 200
	Send, ALL
	Send, {Enter}
	Sleep, 200
	Send, y
	Send, {Enter}	
}















