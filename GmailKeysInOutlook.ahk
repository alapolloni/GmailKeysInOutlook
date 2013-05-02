;*******************************************************************************
;				                          Information					
;*******************************************************************************
; AutoHotkey Version: 	2.x
; Language:       		English
; Platform:       		XP/Vista/7
; Updated by: 				Ty Myrick 
; Author: 					Lowell Heddings (How-To Geek)
; URL: 						http://lifehacker.com/5175724/add-gmail-shortcuts-to-outlook-with-gmail-keys
; Original script by: 	Jayp 
; Original URL: 			http://www.ocellated.com/2009/03/18/pimping-microsoft-outlook/
;
; Script Function: Gmail Keys adds Gmail Shortcut Keys to Outlook 
; Version 2.x updated for Outlook 2010 
;
;*******************************************************************************
;				                          Version History					
;*******************************************************************************
; Version 2.2 - updated to use Outlook 2013 and groups so that it works on some 
;               other windows as well
; Version 2.1 - updated to use ClearContext v5
; Version 2.0 - updated by Ty Myrick to work with Outlook 2010 
; Version 1.0 - updated by Lowell Heddings 
; Version 0.1 - initial set of hotkeys by Jayp
;*******************************************************************************
; # Win (Windows logo key) 
; ! Alt 
; ^ Control 
; + Shift


#NoEnv ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetTitleMatchMode 2 ;allow partial match to window titles

GroupAdd, GroupOutlook, - Outlook ahk_class rctrl_renwnd32  	; Office 2013 Main Mail window
GroupAdd, GroupOutlook, - Message ahk_class rctrl_renwnd32  	; Office 2013 Open Message window
GroupAdd, GroupOutlook, ahk_class #32770 			; Office 2013 - reminders popup 
return

;next line was to test the GroupAdd
;Numpad5::GroupActivate, GroupOutlook ; Assign a hotkey to visit each Outlook window, one at a time.

;************************
;Hotkeys for Outlook 2013
;************************

;As best I can tell, the window text 'NUIDocumentWindow' is not present on any other items except the main window. Also, I look for the phrase ' - Microsoft Outlook' in the title, which will not appear in the title (unless a user types this string into the subject of a message or task).
	#IfWinActive, ahk_group GroupOutlook ; Office 2013
	;#IfWinActive, - Outlook ahk_class rctrl_renwnd32  ; Office 2013
	;#IfWinActive, - Outlook ahk_class rctrl_renwnd32, NUIDocumentWindow ; Office 2013
	;#IfWinActive, - Microsoft Outlook ahk_class rctrl_renwnd32, NUIDocumentWindow ; Office 2010
	;#IfWinActive, - Microsoft Outlook ahk_class rctrl_renwnd32, NUIDocumentWindow or #IfWinActive,  ahk_class #32770
        ;IfWinActive("ahk_class rctrl_renwnd32") or IfWinActive("ahk_class #32770")
        ;IfWinActive,ahk_class rctrl_renwnd32

;               y::HandleOutlookKeys("^+1", "y") 		;archive message using Quick Steps hotkey  
		e::HandleOutlookKeys("!YY5", "e") 		;archive using ClearContext , send thread to pre-selected Project
		+e::HandleOutlookKeys("!YY8", "+e") 		;using ClearContext , pick new Project and send message
		#::HandleOutlookKeys("^d", "#") 		;delete message using regular Control D
		f::HandleOutlookKeys("^f", "f") 		;forwards message 
		r::HandleOutlookKeys("^r", "r") 		;replies to message 
		a::HandleOutlookKeys("^+r", "a") 		;reply all 
;		v::HandleOutlookKeys("^+v", "v") 		;move message box 
		+u::HandleOutlookKeys("^u", "+u") 		;marks messages as unread 
		+i::HandleOutlookKeys("^q", "+i") 		;marks messages as read 
		j::HandleOutlookKeys("{Down}", "j") 		;move down in list 
		+j::HandleOutlookKeys("+{Down}", "+j") 		;move down and select next item 
		k::HandleOutlookKeys("{Up}", "k") 		;move up 
		+k::HandleOutlookKeys("+{Up}", "+k") 		;move up and select next item 
		o::HandleOutlookKeys("^o", "o") 		;open message
;		s::HandleOutlookKeys("{Insert}", "s") 		;toggle flag (star) 
		s::HandleOutlookKeys("^+g", "s") 		;set follow up options (star) 
		c::HandleOutlookKeys("^n", "c") 		;new message 
		/::HandleOutlookKeys("^e", "/") 		;focus search box 
		.::HandleOutlookKeys("+{F10}", ".") 		;Display context menu 
		l::HandleOutlookKeys("!3", "l") 		;categorize message by calling All Categories hotkey from Quick Access Toolbar 
	#IfWinActive


;Passes Outlook a special key combination for custom keystrokes or normal key value, depending on context

	HandleOutlookKeys( specialKey, normalKey ) 
	{
		OutputDebug, DEBUG:HandleOutlookKeys
		;Activates key only on main outlook window, not messages, tasks, contacts, etc. 
		;IfWinActive, - Microsoft Outlook ahk_class rctrl_renwnd32, NUIDocumentWindow, ,  ;NUIDocumentWindow, what is that?
		;IfWinActive, - Microsoft Outlook ahk_class rctrl_renwnd32, , ,   ;Office2010
		;IfWinActive, - Outlook ahk_class rctrl_renwnd32, , ,   ;Office2013
		IfWinActive, ahk_group GroupOutlook ; Office 2013
      		{

			;Find out which control in Outlook has focus
			ControlGetFocus currentCtrl, A 
			;MsgBox, Control with focus = %currentCtrl%, 
			OutputDebug, DEBUG:currentCtrl: %currentCtrl%

			;Set list of controls that should respond to specialKey. Controls are the list of emails and the main (and minor) controls of the reading pane, including controls when viewing certain attachments.
			;Currently I handle archiving when viewing attachments of Word, Excel, Powerpoint, Text, jpgs, pdfs
			;The control 'RichEdit20WPT1' (email subject line) is used extensively for inline editing. Thus it had to be removed. If an email's subject has focus, it won't archive...
			;   also: RichEdit20WPT2 RichEdit20WPT4 _WwG1
			;OutlookGrid1,OutlookGrid2, = Main Message Window
			;SysListView321 = Reminders
			ctrlList = Acrobat Preview Window1,AfxWndW5,AfxWndW6,EXCEL71,MsoCommandBar1,OlkPicturePreviewer1,paneClassDC1,RichEdit20WPT5,RICHEDIT50W1,SUPERGRID2,SUPERGRID1,OutlookGrid1,OutlookGrid2,SysListView321

			if currentCtrl in %ctrlList%
				{
				;OutputDebug, DEBUG:Control in list.  Sending specialKey: %specialKey%
				;MsgBox, Control in list.
				Send %specialKey%
         			} 
			;Allow typing normalKey somewhere else in the main Outlook window. (Like the search field or the folder pane.)
			else 
				{
;					MsgBox, Control not in list.
					Send %normalKey%
				}

      		}
		;Allow typing normalKey in another window type within Outlook, like a mail message, task, appointment, etc.
		else 
			{
				;OutputDebug, DEBUG:normalkey
				Send %normalKey%
			}
	} ;End HandleOutlookKeys

