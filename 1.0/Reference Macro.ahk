#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force ; allows only once version of this program to be open
DetectHiddenText,on
;DetectHiddenWindows, on

/*
*********************************************************************************************************************************************************
***************************************************************Declare Variables******************************************************************
*********************************************************************************************************************************************************
*/

Tempclip = 0
exityes = 1
counter = 0
skip = 0
Version_number = 1.0

Menu, Tray, NoStandard
Menu, Tray, add, Quit,exitmacro
Menu, Tray, add, Check Box for Update,Boxsite

;******************************************************************
;**************** Creates Manual Names ****************************
;******************************************************************
Global Omm_Name , Sysop_Name , Spec_Name , DA_Name , TA_Name , Tblshoot_Name, Other
Omm_Name := "Operation and Maintenance Manual"
Sysop_Name := "Systems Operation"
Spec_Name := "Specifications"
DA_Name := "Disassembly and Assembly"
TA_Name := "Testing and Adjusting "
Tblshoot_Name := "Troubleshooting"
Other := ""

/*
*********************************************************************************************************************************************************
*************************************************************** Autorun section ******************************************************************
*********************************************************************************************************************************************************
*/
CmdLine := DllCall("GetCommandLine", "Str") ; checks if the program was reloaded or restarted
if !RegExMatch(CmdLine, "\/Restart") ; If it was not restarted
{
   skip = 1 ; set Skip variable to 1
   ; set the rest of these cariables to 0
   omm=0 
   spec =0
   da =0
   ta =0
   tbl =0
   sysop =0
   other =0
}
Else,
Restart = 1


If skip != 1 ; of the program was restarted or reloaded
{
   if  (FileExist("Tempsettingslinking.ini")) ; checks for temp file that was created and if it there, it reads the settings
   {
      IniRead, medianumber, Tempsettingslinking.ini, settings, medianumber
      IniRead, omm,  Tempsettingslinking.ini, settings, omm
      IniRead, spec,  Tempsettingslinking.ini, settings, spec
      IniRead, da,  Tempsettingslinking.ini, settings, da
      IniRead, ta,  Tempsettingslinking.ini, settings, ta
      IniRead, tbl,  Tempsettingslinking.ini, settings, tbl
      IniRead, sysop,  Tempsettingslinking.ini, settings, sysop
      IniRead, other,  Tempsettingslinking.ini, settings, other
      IniRead, editfieldtest, Tempsettingslinking.ini, settings, editfieldtest
      IniRead, editIenumber, Tempsettingslinking.ini, settings, editIenumber
      IniRead,Exityes, Tempsettingslinking.ini, settings, Exityes
   }}



;SetTitleMatchMode, slow
activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
;below does math to make the gui screen show up on roughly in the center of the screen 
Amonh /=2
amonw /=2
amonx := amonx + (amonw/2)
amony := amony + (amonh/2)
;Below is the main GUI window before you start the macro
Gui, add, groupbox, w250 h190,
gui, font, cred
Gui, add, text, yp+10 xp+5 vnotvalidbox ,Must have one of the below boxes checked
gui, font
gui, add, Checkbox, Checked%omm% xp yp+20 vomm gcheckeing, %Omm_Name%
gui, add, Checkbox, Checked%spec% vspec gcheckeing, %Spec_Name%
gui, add, Checkbox, Checked%da% vda gcheckeing, %DA_Name%
gui, add, Checkbox, Checked%ta% vta gcheckeing, %TA_Name% 
gui, add, Checkbox, Checked%tbl% vtbl gcheckeing, %Tblshoot_Name%
gui, add, Checkbox, Checked%sysop% vsysop gcheckeing, %Sysop_Name%
gui, add, Checkbox, Checked%other% vother gcheckeing,Other:
gui, add, edit, xp+50 yp-3 w190 h20 veditfieldtest,  %editfieldtest%
Gui, add, text,xp-50 yp+25 h20, (Optional) Media Number:
Gui, add, Edit, xp+125 yp-3 w100 h20 vmedianumber, %medianumber%
Gui, add, text, xp-125 yp+35 h20, IE Number:
Gui, add, Edit, xp+55 yp-3  w100 h20 veditIenumber gIEnumberguichcek, %editIenumber%
Gui, font, cred
Gui, add, Text, xp+105  vnotvalid, IE NUMBER NOT VALID
Guicontrol,hide, notvalid ; NOTE, this hides the red text 
Guicontrol,hide, editfieldtest ; NOTE this hides the "Other" text edit field 
;Guicontrol, hide, notvalidbox
gui, add, button, x5 w300 H30 Default gstart, Start macro
GUIrun = 1
Gosub, checkeing
gui,show, x%amonx% y%amony% w310, Reference Macro V%Version_number%
Gui, +alwaysontop
If other = 1
Guicontrol,show, editfieldtest 
Gosub, IEnumberguichcek ; goes to the subroutine to set the check boxes and edit fields in case program was reloaded

If Exityes = 1
{
   Sleep 100
   WinActivate, Arbortext
   Sleep 100
   Exityes = 0
}
return ; stops the autoload part of the script

/*
*********************************************************************************************************************************************************
*************************************************************** Subroutines section ******************************************************************
*********************************************************************************************************************************************************
*/

IEnumberguichcek:
{
   GuiControlGet,editIenumber ; gets the text from the IENumber box
   
   If editIenumber =  ; checks to see if the edit box is empty and do the line below if it is
   {
      IEnotvalid = 1
      Guicontrol,hide, notvalid ; ensures the red text for Invalid IT number does not show in window
      Return ; stops the subroutine from here so the rest will not run
   }
   ;if the IE number box is not empty
   If (RegExMatch(editIenumber,"^I[0-9]{8}$")) ; finds if the number is a I0 number with an upper case I
   {
      editIenumber := RegexReplace(editIenumber,"I","i") ; replaces the Uppercase I to a Lower case i so it can be found in ACM
   }
   
   if RegExMatch(editIenumber, "^i[0-9]{8}$") ; searches the clipboard for an item that has a i in the front and then 8 numbers after it (a graphics # for ACM) If found, it pastes, if not then it does nothing
   {
      IEnumber = %editIenumber% ; stores the text box text into theGUI window
      Guicontrol,,EDitIEnumber, %editIenumber% ; Pastes the lower case i number to the text box
      If skipend !=1
      {
         Send {end}
      }
      IEnotvalid = 0 ;sets this var to 0
      Guicontrol,hide, notvalid		
      Return
   }
   Else ; If the number in the EditIEnumber text box is not the letter i with 8 numbers after it, run the code below
   {
      Guicontrol,show, notvalid ; show the red text in the gui screen
      IEnotvalid = 1 ; set this var to 1
      Return
   }}

;if the user hits the X on the main gui window
GuiClose:
{
   gui, -alwaysontop ; lets the main gui screen have a screen on top of it
   activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Amonh /=2
   amonw /=2
   
   amonx := amonx + (amonw/2)
   amony := amony + (amonh/2)
   ; Below is a gui screen to ask if you want to quit
   gui,2:add, Text,,Are you sure you want to quit? 
   gui, 2:add, button, gexitmacro, Yes ; if yes is pressed, it goes to exitmacro subroutine
   Gui, 2:add,button, xp+50 gcancelmacro, No ; if no is pressed, program goes to cancelmacro subroutine
   gui, 2:Show,x%amonx% y%amony%,Are you sure? ; show the gui screen and call the window "Are you sure?"
   return
}


;Exitmacro subroutine
exitmacro:
{
   Exityes = 0 ; set var to 0
   GOsub, OnExitMethod ; goes to Onexitmethod subroutine
   return
}


Cancelmacro:
{
   gui,2:destroy
   gui, +alwaysontop
   winactivate Reference macro
   Return
}

OnExitMethod:
{
   If exityes = 1
   {
      GuiControlGet,medianumber
      GuiControlGet,editfieldtest
      GuiControlGet,editIenumber
      
      IfNotExist, Tempsettingslinking.ini ; of ini file isnt there it makes one
      {
         ;msgbox, not found
         FileAppend,, Tempsettingslinking.ini
         sleep 1000
      }
      IniWrite, %medianumber%, Tempsettingslinking.ini, settings, medianumber
      IniWrite, %omm%, Tempsettingslinking.ini, settings, omm
      IniWrite, %spec%, Tempsettingslinking.ini, settings, spec
      IniWrite, %da%, Tempsettingslinking.ini, settings, da
      IniWrite, %ta%, Tempsettingslinking.ini, settings, ta
      IniWrite, %tbl%, Tempsettingslinking.ini, settings, tbl
      IniWrite, %sysop%, Tempsettingslinking.ini, settings, sysop
      IniWrite, %other%, Tempsettingslinking.ini, settings, other
      IniWrite, %editfieldtest%, Tempsettingslinking.ini, settings, editfieldtest
      IniWrite, %editIenumber%, Tempsettingslinking.ini, settings, editIenumber
      IniWrite,%Exityes%, Tempsettingslinking.ini, settings, Exityes
      
      ;pause,on
      If A_IsCompiled
      {
         Run, %A_ScriptFullPath% /Restart
         Exitapp
      }else  {
         Run, %A_AhkPath% "%A_ScriptFullPath%" /Restart
         exitapp
      }
      Return
   }
   
   
   
   ;pause, on
   IfExist, Tempsettingslinking.ini 
      FileDelete, Tempsettingslinking.ini
   If Exityes = reload
   Reload
   
   Exitapp
   Return
}

Boxsite:
{
   Run, https://cat.box.com/s/c0d95xwih40ia11lhmmvjzm9idzal4q7
   Return
}

macroend:
{
   If pubtype = other
   {
      GuiControlGet,editfieldtest
      pubtype := editfieldtest
   }
   
   GuicontrolGet, Medianumber
   Pubtype := Getpubtype()
   
   activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Amonh /=2
   amonw /=2
   
   amonx := amonx + (amonw/2)
   amony := amony + (amonh/2)
   frontend :=  "<pubtype>" pubtype "</pubtype><media><formno>" Medianumber "</formno></media>"
   gui,3: add, text,, Press Copy text to copy text in field and close window `, then paste the text in front of the ie-topic tag
   gui,3:add,edit,vcopyme, %frontend%
   gui, 3:add, button, Default gcopytext, copy text
   gui,3:show, x%amonx% y%amony%, Reference Macro
   return
}

copytext:
{
   Guicontrol, Focus, copyme
   Send ^a
   sleep 200
   Send ^c
   sleep 200
   Gui, 3:destroy
   mousemove, 0,0
   Exityes = reload
   GOsub, OnExitMethod
   return
}

3guiclose:
{
   Guicontrol, Focus, copyme
   Send ^a
   sleep 200
   Send ^c
   sleep 200
   Gui, 3:destroy
   Exityes = reload
   GOsub, OnExitMethod
}

start:
{
   skipend = 1
   GOsub, IEnumberguichcek
   ;msgbox, ienumber is %IEnotvalid%
   boxcheck = 0
   GuicontrolGet, omm
   GuicontrolGet, spec
   GuicontrolGet, da
   GuicontrolGet, ta
   GuicontrolGet, tbl
   GuicontrolGet, sysop
   GuicontrolGet, other
   
   If omm = 1
   boxcheck = 1
   
   If spec = 1
   boxcheck = 1
   
   if da = 1
   boxcheck = 1
   
   if ta = 1
   boxcheck = 1
   
   if tbl = 1
   boxcheck = 1
   
   if sysop = 1
   boxcheck = 1
   
   if other = 1
   boxcheck = 1
   
   
   
   
   If (IEnotvalid = 1) and (boxcheck = 0)
   {
      ;msgbox, go
      Loop,5
      {
         Winactivate, Reference Macro
         Guicontrol,hide, notvalid
         Guicontrol,hide, notvalidbox
         sleep 100
         Guicontrol,Show, notvalid	
         Guicontrol,Show, notvalidbox
         Sleep 300
      }	
      Return
   }
   
   If boxcheck = 0
   {
      ;msgbox, boxgo
      Loop,5
      {
         Winactivate, Reference Macro
         Guicontrol,hide, notvalidbox
         sleep 100
         Guicontrol,Show, notvalidbox
         Sleep 300
      }
      Return
   }
   Else if boxcheck = 1
   {
      Guicontrol,hide, notvalidbox
   }
   
   
   If IEnotvalid = 1
   {
      Loop,5
      {
         Winactivate, Reference Macro
         Guicontrol,hide, notvalid	
         sleep 100
         Guicontrol,Show, notvalid	
         Sleep 300
      }	
      Return
   }
   
   
   
   Gui, -AlwaysOnTop
   gosub, macrostart 
   return
}


Macrostart:
{
   tempvar := clipboardall
   Gosub, Searchclose
   Gosub, Broswerclose
   
   
   SetTitleMatchMode,2
   IfWinNotExist, Resource Manager
   {
      Gosub, NoresourceGUI
      exit
      SetTitleMatchMode 2
      Gosub, resourceexitst
      Gui, %guinumber%:Destroy
      sleep 3000
   }
   ActivateResourcemanager()
   
   Resourcemove()
   sleep 200
   ;Resourcesizecheck()
   sleep 200
   ClickLocations(Inc1,Inc2,xref1,xref2)
   CoordMode, Mouse, Relative	
   /*
   WinActivate, Resource Manager
   sleep 500
   WinWaitActive, Resource Manager,,10
   */
   
   ActivateResourcemanager()
   
   sleep 250
   CoordMode, Mouse, Relative	
   click 10,0
   sleep 200
   Send {Alt Down}
   Sleep 100
   Send {L}{Alt UP}
   Sleep 200
   Send CMS
   Sleep 250
   Mousemove 200,200
   sleep 250
   Click Right
   
   sleep 500
   Loop,2
   {
      Send {Up}
      sleep 250
   }
   Send {Enter}
   sleep 150
   Send {up 10}	
   mousemove 0,0
   
   Gosub, Searchclose
   Gosub, Broswerclose
   
   SetTitleMatchMode,2
   
   WinActivate Arbortext ; activate arbortext window
   sleep 100
   WinWaitActive, Arbortext,,10
   sleep 750
   
   gosub, CreateIEsearch
   
   Return
}


SetTitleMatchMode, 2
winmovemsgbox:
{
   SetTimer, WinMoveMsgBox, OFF 
   WinMove, %titletext% , Amonx, Amony 
   return
}



CreateIEsearch:
{
   SetTitleMatchMode,2
   WinActivate Arbortext
   sleep 100
   WinWaitActive, Arbortext,,30
   sleep 1000
   Send {Alt Down}{o}{Alt Up}{e} ; send keystrokes
   Sleep 750
   
   While !WinExist("Search")
   {
      Sleep 300
   }
   
   WinActivate, search
   sleep 100
   WinWaitActive, Search,,30 ; wait for search window to be the active window, times out in 11 seconds
   If ErrorLevel ; if times out, the subroutine will stop
   {
      msgbox, Search window did not load within 30 seconds. Macro is exiting
      exityes = 1
      gosub, OnExitMethod   ; stops the subroutine
   }else  {
      
   }	
   Sleep 250
   Searchemove()
   sleep 500	;pauses script for .5 seconds
   Send {Tab}{A}{l}{l}{c}{o}{n} ; sends keystrokes to window
   sleep 300 ; pauses script for 200 milliseconds
   Send {tab}{i}{n}{f}{o}{r} ; sends keystrokes to window
   Sleep 300 ; pauses script for 200 milliseconds
   Send {tab 2}{right}; sends keystrokes to window
   Sleep 300 ; pauses script for 400 milliseconds
   Send {tab 8}{down 13}; sends keystrokes to window
   Sleep 300 ; pauses script for 400 milliseconds
   Send {tab}{enter}
   Sleep 300 ; pauses script for 200 milliseconds
   Send {tab 7} ; sends keystrokes to window
   sleep 250
   GuiControlGet,editIenumber
   Send %IEnumber%
   Sleep 500
   Send {enter}
   sleep 250 ; pauses script for 100 milliseconds
   SetTimer, Checkforfail, 100
   
   +++SetTitleMatchMode,3
   WinActivate, Browser AHK_class TTAFrameXClass
   sleep 100
   WinWaitActive, Browser,, 30
   If ErrorLevel
   {
      msgbox, Browser did not open after 10 seconds. Reloading macro
      exityes = 1
      gosub, OnExitMethod
   }
   
   WinMinimize, Search AHK_class TTAFrameXClass
   sleep 100
   Winactivate, Arbortext AHK_class TTAFrameXClass
   Sleep 10
   WinWaitActive,Arbortext,,5
   Sleep 150
   Send !{Escape}
   
   WinActivate, Browser AHK_class TTAFrameXClass
   Sleep 100
   Browseemove()
   sleep 250
   ; new test scrpit
   Send {Alt Down}{p}{Alt up}
   sleep 100
   sleep 250
   Checkwait := Waitforobjectwindow()
   If checkwait = Timedout
   sleep 500
   Checkwait := Waitforobjectwindow()
   If checkwait = Timedout
   {
      msgbox, Cannot get the object window, Please manually enter link from this point forward or restart the macro by pressing ESC key in Reference Macro window
      Exit
   }
   
   Sleep 500
   Send {Shift Down}
   Sleep 250
   Loop,4
   {
      Send {Shift down}{Tab}{shift up}
      Sleep 100
   }
   
   Loop,5 
   {
      Send {Ctrl down}{c}{Ctrl up}
      sleep 250
      FoundPos := (RegExMatch(Clipboard, "^i[0-9]{8}$"))
      ;msgbox, %FoundPos%
      If FoundPos != 0
      {
         Send {Shift down}{Tab}{shift up}
         Break
      }else  {
         clipboard = 		
         Send {Shift down}{Tab}{shift up}
         Sleep 200
      }}
	  
   If FoundPos = 0
	msgbox, Could not Find Title `n Please either hit Esc and restart the macro, or continue maually
	
   Send {Ctrl down}
   Sleep 50
   Send {c}{Ctrl up}
   Sleep 300
   
   Send {Alt down}
   SLeep 150
   SEnd {c}{Alt up}
   Sleep 150
   
   ;end new test script
   
   CoordMode, Mouse, Relative
   Click 300,200
   
   
   sleep 500
   Send {up 3}
   sleep 1000
   loop,3
   {
      Send {tab}
      sleep 300
   }
   sleep 1000
   Send {Alt down}{g}{Alt up}
   Sleep 300
   WinMinimize, Browser AHK_class TTAFrameXClass
   sleep 200
   SetTitleMatchMode,2
   Winactivate, Resource Manager
   sleep 100
   ActivateResourcemanager()
   Sleep 250
   CoordMode, mouse, Screen
   SetTimer, tooltipfollow, off
   Tooltip, 
   CoordMode, Mouse, Relative	
   Winactivate, Resource Manager
   sleep 100
   click  200,200
   mousemove, 0,0
   sleep 250
   send {down}
   sleep 250
   send {right}
   Sleep 250
   Send {Down}
   
   sleep 200
   CoordMode, Mouse, Relative
   
   Send %clipboard%
   Sleep 750
   
   Send {Alt down}
   sleep 200
   SEnd {I}
   sleep 200
   Send {Alt up}
   Sleep 200
   WinActivate Arbortext
   Sleep 10
   WinWaitActive, Arbortext
   Sleep 1000
   Send {Shift down}
   sleep 350
   send {right}
   sleep 400
   send {Shift up}
   SLEEP 400
   SEND {right}
   sleep 500
   ActivateResourcemanager()
   
   sleep 250
   GetOverlappingWindows()
   sleep 250
   CoordMode, Mouse, Screen	
   Click %Xref1%, %xref2%
   ;CoordMode, Mouse, screen
   sleep 500
   CoordMode, Mouse, Relative
   click 10,0
   sleep 500
   Send {Alt Down}
   Sleep 100
   Send {L}{Alt UP}
   Sleep 500
   Send CMS	
   sleep 200
   click  200,200
   sleep 250
   mousemove, 0,0
   sleep 250
   Send {up 9}
   sleep 100
   Send {down}
   sleep 200
   SEnd {right}
   
   Sleep 175	
   Send {Alt down}
   Sleep 100
   Send {I}
   Sleep 100
   Send {Alt up}
   +++SetTitleMatchMode, 2
   sleep 1000	
   WinActivate, Arbortext
   sleep 10
   WinWaitActive, Arbortext
   Sleep 1000
   Click 300,0
   Send {Shift DOwn}{LEft}
   Sleep 500
   SEnd {LEft}
   Sleep 500
   SEnd {Shift up}
   Sleep 500
   Send {LEft}
   sleep 500
   Send {right}
   
   Sleep 500
   Gosub, Searchclose
   Gosub, Broswerclose
   sleep 500
   gosub, macroend
   Return
}

Checkforfail:
{
   IfWinactive Browser
   {
      SetTimer, Checkforfail, off
   }
   IfWinExist CMS Diagnostic
   {
      SetTimer, Checkforfail, off
      activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
      
      Amonh /=2
      amonw /=2
      
      amonx := amonx + (amonw/2)
      amony := amony + (amonh/2)
      gui, add, Text, xp+5 yp+20 w500 h20 BackgroundTrans, I0# incorrect Now reloading macro
      Gui, Show, x%amonx% y%amony% , Error!!
      sleep 3000
      exityes = 1
      gosub, OnExitMethod
   }
   Return
}

Broswerclose:
{
   IfWinExist Browser
   {
      WinActivate, Browser AHK_class TTAFrameXClass
      Sleep 500
      WinWaitActive, Browser,,10
      sleep 500
      send {Escape down}
      Sleep 100
      SEnd {Escape up}
      Sleep 200
   }
   return
}

Searchclose:
{
   SetTitleMatchMode 3
   IfWinExist Search
   {
      WinActivate, Search AHK_class TTAFrameXClass
      Sleep 500
      WinWaitActive, Search,,10
      sleep 500
      send {Escape down}
      Sleep 100
      SEnd {Escape up}
      Sleep 200
   }
   return
}

SetTitleMatchMode, 2
NoresourceGUI:
{
   guinumber=7
   activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Amonh /=2
   amonw /=2
   
   amonx := amonx + (amonw/2)
   amony := amony + (amonh/2)
   gui %guinumber%: +alwaysontop
   ;msgbox, Prefix is %Guitextlocation%`n Serial is %Clippy%-%Clippy2%`n Mod is %Modifier%
   gui, 7:add, Text, xp+5 yp+20  h20 BackgroundTrans, Cannot find Resource Manager.
   gui, 7:add, Text, xp yp+20  h20 BackgroundTrans, Please Undock the resource manager and rerun macro. 
   gui, 7:add, Text, xp yp+20 h40 BackgroundTrans, If the Resource Manager is not open, Go to View -> Resource Manager. Then undock the resource manager by dragging it outside of the Arbortext Screen
   Gui, 7:Add,button, xp yp+40 Default gclosegui, OK
   Gui, 7:Show, x%amonx% y%amony% , Not Found
   Gui 1:+Alwaysontop
   return
}

noio:
{
   guinumber=6
   activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
   
   Amonh /=2
   amonw /=2
   
   amonx := amonx + (amonw/2)
   amony := amony + (amonh/2)
   gui %guinumber%: +alwaysontop
   ;msgbox, Prefix is %Guitextlocation%`n Serial is %Clippy%-%Clippy2%`n Mod is %Modifier%
   gui, 6:add, Text, xp+5 yp+5 w500 h40 BackgroundTrans, Please copy/paste or enter an i0 Number in the field below and press OK.
   gui, 6:add, Edit, xp yp+20 w100 BackgroundTrans vIEnumber, i0
   Gui, 6:Add,button, Default xp yp+40 gcloseguiie, OK
   Gui, 6:Show, x%amonx% y%amony% , No i0 Number
   Gui, 6:submit, nohide
   Pause, on 
   return
}


tooltipfollow:
{
   ;CoordMode, mouse, screen
   Tooltip, "Double click  on Inclusion tab"
}

closegui:
{
   If guinumber = 5
   SetTimer, tooltipfollow, 50
   
   Pause, off
   Gui, %guinumber%:Destroy
   Return
}

closeguiie:
{
   Pause, off
   GuiControlget,IEnumber
   Stringreplace, IEnumber,IEnumber,%a_space%,,all
   If (RegExMatch(IEnumber,"^I[0-9]{8}$")) ; finds if the number is a Gnumber with an upper case G
   {
      IEnumber := RegexReplace(IEnumber,"I","i") ; replaces the Uppercase G to a Lower case g so it can be found in ACM
   }
   if RegExMatch(IEnumber, "^i[0-9]{8}$") ; searches the clipboard for an item that has a g in the front and then 8 numbers after it (a graphics # for ACM) If found, it pastes, if not then it does nothing
   {
      Gui, %guinumber%:Destroy
   }else  {
      Gui, %guinumber%:Destroy
      guinumber=6
      activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
      
      Amonh /=2
      amonw /=2
      
      amonx := amonx + (amonw/2)
      amony := amony + (amonh/2)
      gui %guinumber%: +alwaysontop
      ;msgbox, Prefix is %Guitextlocation%`n Serial is %Clippy%-%Clippy2%`n Mod is %Modifier%
      gui, 6:add, Text, xp+5 yp+5 w500 h40 BackgroundTrans, Please copy/paste or enter an i0 number in the field below and press OK.
      GUI,6:font, cred
      gui, 6:add, Text, xp yp+45 w500 h40 BackgroundTrans, Please enter a valid i0 number below				
      gui, 6:add, Edit, xp yp+20 w100 BackgroundTrans vIEnumber, %IEnumber%
      Gui, 6:Add,button, xp yp+40 Default gcloseguiie, OK
      Gui, 6:Show, x%amonx% y%amony% , No i0 Number
      Gui, 6:submit, nohide
      gui, 6: font,
      Pause, on
   }
   
   
   Sleep 500
   Return
}


resourceexitst:
{
   Countwinloops = 0
   Loop
   {
      Countwinloops++
      IF Countwinloops = 20
      {
         Msgbox, Macro couldn't find the resource manager window`n Marco will now exit 
         Exit
         Break
      }
      
      IfWinNotExist, Resource Manager 
      {
         Sleep 500
      }
      
      Else  IfWinExist, Resource Manager
      {
         Break
      }}
      Return
   }
   
   CoordMode, Mouse, screen
   ClickLocations(Byref Inc1, Byref Inc2,Byref xref1,Byref xref2)
   {
      CoordMode, Mouse, screen
      guinumber=4
      activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
      
      Amonh /=2
      amonw /=2
      
      amonx := amonx + (amonw)
      amony := amony + (amonh)
      gui %guinumber%: +alwaysontop
      ;msgbox, Prefix is %Guitextlocation%`n Serial is %Clippy%-%Clippy2%`n Mod is %Modifier%
      gui, %guinumber%:add, Text, xp+5 yp+20  h20 BackgroundTrans, Double click the Link/Xref tab.
      gui, %guinumber%:add, Text, xp yp+20  h20 BackgroundTrans, Macro will continue after the Double click.
      ;Gui, %guinumber%:Add,button, Default xp yp+40 gclosegui, OK
      Gui, %guinumber%:Show, x%amonx% y%amony% , Click Xref tab
      
      gosub, doublecheck
      MouseGetPos xref1x,xref2x
      
      Gui, %guinumber%:Destroy
      Sleep 500	
      
      
      
      guinumber=5
      
      gui %guinumber%: +alwaysontop
      ;msgbox, Prefix is %Guitextlocation%`n Serial is %Clippy%-%Clippy2%`n Mod is %Modifier%
      gui, %guinumber%:add, Text, xp+5 yp+20  h20 BackgroundTrans, Double click the inclusion tab.
      gui, %guinumber%:add, Text, xp yp+20  h20 BackgroundTrans, Macro will continue after the double click
      ;Gui, %guinumber%:Add,button, xp yp+40 Default gclosegui, OK
      Gui, %guinumber%:Show, x%amonx% y%amony% , Click inclusion tab
      
      gosub, doublecheck
      
      
      MouseGetPos Inc1x,Inc2x
      
      
      Gui, %guinumber%:Destroy
      Sleep 500
      inc1 := inc1x
      inc2 := inc2x
      xref1 := xref1x
      xref2 := xref2x
      Return
   }
   
   doublecheck:
   {
      sleep 100
      If doubleclick = 1
      Return
      
      Else if Doubleclick = 0
      {
         GOsub doublecheck	
         Return
      }
      Return
   }
   
   
   listlines off
   ~LButton:: 
   {
      Settimer, Resetvar, 500
      If (A_TimeSincePriorHotkey<400) and (A_PriorHotkey="~LButton")
      {
         Doubleclick = 1
         Settimer, Resetvar, 500
      }
      
      ELse Doubleclick = 0
      Return
   }
   
   checkeing:
   {
      boxcheck = 0
      GuicontrolGet, omm
      GuicontrolGet, spec
      GuicontrolGet, da
      GuicontrolGet, ta
      GuicontrolGet, tbl
      GuicontrolGet, sysop
      GuicontrolGet, other
      
      If omm = 1
      boxcheck = 1
      
      If spec = 1
      boxcheck = 1
      
      if da = 1
      boxcheck = 1
      
      if ta = 1
      boxcheck = 1
      
      if tbl = 1
      boxcheck = 1
      
      if sysop = 1
      boxcheck = 1
      
      if other = 1
      boxcheck = 1
      
      if boxcheck = 1
      Guicontrol,hide, notvalidbox
      
      If guirun != 1	
      {
         If boxcheck != 1
         {
            Loop,5
            {
               Winactivate, Reference Macro
               Guicontrol,hide, notvalidbox
               sleep 100
               Guicontrol,Show, notvalidbox
               Sleep 300
            }
         }else  {
            Guicontrol,hide, notvalidbox
         }
         
         If GUIrun = 1
         {
            If boxcheck != 1
            Guicontrol,Show, notvalidbox
            guirun = 0
         }}
      If (A_GuiControl="omm")
      {
         Spec = 0
         da = 0
         ta=0
         tbl=0
         sysop=0
         other=0
         guicontrol,,spec,0
         guicontrol,,da,0
         guicontrol,,ta,0
         guicontrol,,tbl,0
         guicontrol,,sysop,0
         guicontrol,,Other,0
         Guicontrol,hide, editfieldtest
         gui, submit,nohide
         Pubtype := Omm_Name
         Tabs = 6
      }
      If (A_GuiControl="spec")
      {
         omm = 0
         da = 0
         ta=0
         tbl=0
         sysop=0
         other=0
         guicontrol,,omm,0
         guicontrol,,da,0
         guicontrol,,ta,0
         guicontrol,,tbl,0
         guicontrol,,sysop,0
         guicontrol,,Other,0
         Guicontrol,hide, editfieldtest
         gui, submit,nohide
         Pubtype := Spec_Name
         Tabs = 6
      }
      If (A_GuiControl="da")
      {
         Spec = 0
         omm = 0
         ta=0
         tbl=0
         sysop=0
         other=0
         guicontrol,,spec,0
         guicontrol,,omm,0
         guicontrol,,ta,0
         guicontrol,,tbl,0
         guicontrol,,sysop,0
         guicontrol,,Other,0
         Guicontrol,hide, editfieldtest
         gui, submit,nohide
         Pubtype :=DA_Name
         Tabs = 6
      }
      If (A_GuiControl="ta")
      {
         Spec = 0
         da = 0
         omm=0
         tbl=0
         sysop=0
         other=0
         guicontrol,,spec,0
         guicontrol,,da,0
         guicontrol,,omm,0
         guicontrol,,tbl,0
         guicontrol,,sysop,0
         guicontrol,,Other,0
         Guicontrol,hide, editfieldtest
         gui, submit,nohide
         Pubtype := TA_Name
         Tabs = 7
      }
      If (A_GuiControl="tbl")
      {
         Spec = 0
         da = 0
         ta=0
         omm=0
         sysop=0
         other=0
         guicontrol,,spec,0
         guicontrol,,da,0
         guicontrol,,ta,0
         guicontrol,,omm,0
         guicontrol,,sysop,0
         guicontrol,,Other,0
         Guicontrol,hide, editfieldtest
         gui, submit,nohide
         Pubtype := Tblshoot_Name
         Tabs = 6
      }
      If (A_GuiControl="sysop")
      {
         Spec = 0
         da = 0
         ta=0
         tbl=0
         omm=0
         other=0
         guicontrol,,spec,0
         guicontrol,,da,0
         guicontrol,,ta,0
         guicontrol,,tbl,0
         guicontrol,,omm,0
         guicontrol,,Other,0
         Guicontrol,hide, editfieldtest
         gui, submit,nohide
         Pubtype := Sysop_Name
         Tabs = 6
      }
      If (A_GuiControl="other")
      {
         Spec = 0
         da = 0
         ta=0
         tbl=0
         sysop=0
         omm=0
         guicontrol,,spec,0
         guicontrol,,da,0
         guicontrol,,ta,0
         guicontrol,,tbl,0
         guicontrol,,sysop,0
         guicontrol,,omm,0
         Guicontrol,show, editfieldtest
         gui, submit,nohide
         GuiControlGet,editfieldtest
         Pubtype := "Other"
         Tabs = 6
      }
      
      
      Return
   }
   
   
   listlines, off
   Resetvar:
   {
      If Doubleclick = 1
      {
         Doubleclick = 0
         SetTimer,resetvar, off
      }
      Return
   }
   Listlines on
   activeMonitorInfo( ByRef aX, ByRef aY, ByRef aWidth,  ByRef  aHeight, ByRef mouseX, ByRef mouseY  )
   {
      CoordMode, Mouse, Screen
      MouseGetPos, mouseX , mouseY
      SysGet, monCount, MonitorCount
      Loop %monCount%
      {
         SysGet, curMon, Monitor, %a_index%
         if ( mouseX >= curMonLeft and mouseX <= curMonRight and mouseY >= curMonTop and mouseY <= curMonBottom )
         {
            aX      := curMonTop
            ay      := curMonLeft
            aHeight := curMonBottom - curMonTop
            aWidth  := curMonRight  - curMonLeft
            return
         }}}
         
         
         ~esc:: 
         {
            exityes = 1
            gosub, OnExitMethod  
            Return
         }
         
         Waitforobjectwindow()
         {
            WinWaitActive, Object Properties,, 30
            if errorlevel = 1
            {
               Return Timedout
            }
            Else
               Return Success
         }
         
         +++Settitlematchmode,2
         
         GetOverlappingWindows() {
            Winget,HWND,ID, Resource Manager
            VarSetCapacity(RECT1, 16, 0)
            
            VarSetCapacity(RECT2, 16, 0)
            
            PtrType := A_PtrSize ? "Ptr" : "UInt" ; AHK_L : AHK_Basic
            
            Overlapping := ""
            
            DllCall("User32.dll\GetWindowRect", PtrType, HWND, PtrType, &RECT1)
            
            WinGet, HWIN, List
            
            Loop, %HWIN% {
               If (HWIN%A_Index% = HWND)
               
                  Break
               
               DllCall("User32.dll\GetWindowRect", PtrType, HWIN%A_Index%, PtrType, &RECT2)
               
               ; http://msdn.microsoft.com/en-us/library/dd145001(VS.85).aspx
               
               If DllCall("User32.dll\IntersectRect", PtrType, &RECT2, PtrType, &RECT1, PtrType, &RECT2, "UInt")
               
               Overlapping .= (Overlapping ? "," : "") . HWIN%A_Index%
               ; WinGetTitle, Overlapping, ahk_id %Overlapping%
            }
            
            Loop, Parse, Overlapping,`,
            {
               WinGetTitle, Overlapping, AHK_id %a_loopField%
               
               WinMinimize,%Overlapping% AHK_id %a_loopField%
               ;Overlappingn = %overlappingn%%overlapping%`n
            }
            If overlappingn !=
            {
               /*
               activeMonitorInfo( amony,Amonx,AmonW,AmonH,mx,my ) ;gets the coordinates of the screen where the mouse is located.
               
               Amonh /=2
               amonw /=2
               
               amonx := amonx + (amonw)
               amony := amony + (amonh)
               Gui 10:add, text, x5 y5, Please move the following window(s) away from the Resource Manager window to prevent macro faliure:
               GUi, 10:add, edit, xp+6 yp+20 ,%Overlappingn%
               GUi, 10:add, button, gok,OK
               gui 10:show,x%amonx% y%amony%,Move Windows
               pause, on
               
               
               
               Ok:
               {
                  Pause, off
                  Gui,10:Destroy
                  GetOverlappingWindows()
                  Return
               }
               */
               
               Return	
               ; msgbox, Please move the following windows away from the Resource Manager screen to prevent macro faliure:`n`n%Overlappingn%
            }
            Else
               ActivateResourcemanager()
            /* 
            WinActivate, Resource Manager
            sleep 100
            WinWaitActive, Resource Manager,,10
            */
            Sleep 250
            return
         }
         
         listlines, on
         
         ActivateResourcemanager()
         {
            CoordMode, Mouse, Relative 
            WinActivate, Resource Manager
            sleep 100
            WinWaitActive, Resource Manager,,10
            If errorlever = 1
            {
               Msgbox, Cannot Find Resource Manager. Macro will stop.
               Exit
            }
            sleep 100
            Click 50,10
            sleep 100
            Return
         }
         
         Resourcemove()
         {
            CoordMode, Mouse, Relative 
            WinActivate, Resource Manager AHK_class TTAFrameXClass
            Sleep 10
            WinWaitActive,Resource Manager,,3
            sleep 500
            ;Click 50,10
            ;MouseMove 50,10
            SendEvent {Click 50,10,Down}
            CoordMode, Mouse, Screen
            Wingetpos,X,Y,W,H,Resource Manager 
            W /=2
            H /=2
            Scrwidth := (A_ScreenWidth / 2 - w)
            Scrheight := (A_ScreenHeight / 2 - h)
            SEndEvent {Click %Scrwidth%,%Scrheight%, up} 
            Return
            CoordMode, Mouse, Relative 
            Return
         }
         
         Searchemove()
         {
            WinActivate, Search AHK_class TTAFrameXClass
            Sleep 10
            WinWaitActive,Search,,3
            sleep 500
            ;Click 50,10
            ;MouseMove 50,10
            SendEvent {Click 50,10,Down}
            CoordMode, Mouse, Screen
            SEndEvent {Click 50,500, up} 
            Return
         }
         
         Browseemove()
         {
            sleep 750
            CoordMode, Mouse, Relative 
            
            WinActivate, Browser  AHK_class TTAFrameXClass
            Sleep 10
            WinWaitActive,Browser,,3
            sleep 500
            ;Click 50,10
            ;MouseMove 50,10
            WinActivate, Browser AHK_class TTAFrameXClass
            Sleep 10
            WinWaitActive,Browser,,3
            sleep 300
            IfWinNotActive, Browser AHK_class TTAFrameXClass
            {
               Sleep 500
               WinActivate, Browser  AHK_class TTAFrameXClass
               Sleep 10
               WinWaitActive,Browser,,3
               sleep 500
            }
            SendEvent {Click 50,10,Down}
            CoordMode, Mouse, Screen
            SEndEvent {Click 50,500, up} 
            Return
         }
         
Getpubtype()
   {
            GuicontrolGet, omm
            GuicontrolGet, spec
            GuicontrolGet, da
            GuicontrolGet, ta
            GuicontrolGet, tbl
            GuicontrolGet, sysop
            GuicontrolGet, other
            
            If omm = 1
            Pubtype := Omm_Name
            
            If spec = 1
            Pubtype := Spec_Name
            
            if da = 1
            Pubtype := DA_Name
            
            if ta = 1
            Pubtype := TA_Name
            
            if tbl = 1
            Pubtype := Tblshoot_Name
            
            if sysop = 1
            Pubtype := Sysop_Name
            
            if other = 1
			{
            GuicontrolGet, editfieldtest
			Pubtype := editfieldtest
			
			}
            Return Pubtype
         }
         
         
         Resourcesizecheck()
         {
            Wingetpos,X,Y,W,H,Resource Manager 
            Coordmode, mouse, Relative
            Winactivate, Resource Manager ahk_class TTAFrameXClass
            Sleep 300
            W := (w - 4)
            H := (H - 4)
            MouseMove,%w%,%h%
            ;Pause, on
            ;Msgbox, height is %h% `n Width is %w%
            
            If W < 350
            {
               Newwidth := (350 - w )
               Newwidth := (Newwidth - w)
               ;msgbox, Width was %w%`n Width is now %Newwidth%
            }
            
            If W => 350
            {
               Newwidth := w
               ;msgbox, Width was %w%`n Width is now %Newwidth%
            }
            
            If H < 500
            {
               Newneight := (500 - h )
               Newneight := (Newneight - h)
               ;msgbox, Width was %w%`n Width is now %Newwidth%
			}
         
         If H => 500
         {
            Newneight := h
            ;msgbox, Width was %w%`n Width is now %Newwidth%
         }
      
      SendEvent {Click %w%,%h%,Down}
      Sleep 200
      
      SendEvent {Click %Newwidth%,%Newneight%, up}
      sleep 250
      
      return
		}
   
^`::
   {
      Listlines
      Pause, on
	}

Pause::
Pause, Toggle              
