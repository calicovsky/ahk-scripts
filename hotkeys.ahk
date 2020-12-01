/*

=============================================
Some basic rules to remember
=============================================

http://www.autohotkey.com/board/topic/554-probably-dumb-newbie-question-re-multi-line-scripts/
---
...Also, you can't put the first line of a multi-line subroutine on the same
line as its hotkey definition. You can only do that if the subroutine consists of
exactly one line. This is because there is an implicit "return" immediately
after a single-line hotkey

*/

/*
	http://www.autohotkey.com/board/topic/60985-get-paths-of-selected-items-in-an-explorer-window/

	Library for getting info from a specific explorer window (if window handle not specified, the currently active
	window will be used).  Requires AHK_L or similar.  Works with the desktop.  Does not currently work with save
	dialogs and such.

	Explorer_GetSelected(hwnd="")   - paths of target window's selected items
	Explorer_GetAll(hwnd="")        - paths of all items in the target window's folder
	Explorer_GetPath(hwnd="")       - path of target window's folder
	
	example:
		F1::
			path := Explorer_GetPath()
			all := Explorer_GetAll()
			sel := Explorer_GetSelected()
			MsgBox % path
			MsgBox % all
			MsgBox % sel
		return
	
	Joshua A. Kinnison
	2011-04-27, 16:12
*/

Explorer_GetPath(hwnd="")
{
	if !(window := Explorer_GetWindow(hwnd))
		return ErrorLevel := "ERROR"
	if (window="desktop")
		return A_Desktop
	path := window.LocationURL
	path := RegExReplace(path, "ftp://.*@","ftp://")
	StringReplace, path, path, file:///
	StringReplace, path, path, /, \, All 
	
	; thanks to polyethene
	Loop
		If RegExMatch(path, "i)(?<=%)[\da-f]{1,2}", hex)
			StringReplace, path, path, `%%hex%, % Chr("0x" . hex), All
		Else Break
	return path
}
Explorer_GetAll(hwnd="")
{
	return Explorer_Get(hwnd)
}
Explorer_GetSelected(hwnd="")
{
	return Explorer_Get(hwnd,true)
}

Explorer_GetWindow(hwnd="")
{
	; thanks to jethrow for some pointers here
    WinGet, process, processName, % "ahk_id" hwnd := hwnd? hwnd:WinExist("A")
    WinGetClass class, ahk_id %hwnd%
	
	if (process!="explorer.exe")
		return
	if (class ~= "(Cabinet|Explore)WClass")
	{
		for window in ComObjCreate("Shell.Application").Windows
			if (window.hwnd==hwnd)
				return window
	}
	else if (class ~= "Progman|WorkerW") 
		return "desktop" ; desktop found
}
Explorer_Get(hwnd="",selection=false)
{
	if !(window := Explorer_GetWindow(hwnd))
		return ErrorLevel := "ERROR"
	if (window="desktop")
	{
		ControlGet, hwWindow, HWND,, SysListView321, ahk_class Progman
		if !hwWindow ; #D mode
			ControlGet, hwWindow, HWND,, SysListView321, A
		ControlGet, files, List, % ( selection ? "Selected":"") "Col1",,ahk_id %hwWindow%
		base := SubStr(A_Desktop,0,1)=="\" ? SubStr(A_Desktop,1,-1) : A_Desktop
		Loop, Parse, files, `n, `r
		{
			path := base "\" A_LoopField
			IfExist %path% ; ignore special icons like Computer (at least for now)
				ret .= path "`n"
		}
	}
	else
	{
		if selection
			collection := window.document.SelectedItems
		else
			collection := window.document.Folder.Items
		for item in collection
			ret .= item.path "`n"
	}
	return Trim(ret,"`n")
}

SetTitleMatchMode 2

#InstallKeybdHook
FormatTime, TimeString,, yyyyMMdd
; -------------- key navigations --------------

;==========================================================
; Emacs/nano like shortcuts (left/right/up/down, etc.)
;==========================================================

SC07B & e::
	If GetKeyState("LShift","P")
		Send +{End}
	Else
		Send {End}
Return
SC07B & a::
	If GetKeyState("LShift","P")
		Send +{Home}
	Else
		Send {Home}
Return
SC07B & f::
	If GetKeyState("LShift","P")
		Send +{Right}
	Else
		Send {Right}
Return
SC07B & b::
	If GetKeyState("LShift","P")
		Send +{Left}
	Else
		Send {Left}
Return
SC07B & p::
	If GetKeyState("LShift","P")
		Send +{Up}
	Else
		Send {Up}
Return
SC07B & n::
	If GetKeyState("LShift","P")
		Send +{Down}
	Else
		Send {Down}
Return
SC07B & d::
	If GetKeyState("LShift","P")
		Send +{Delete}
	Else
		Send {Delete}
Return

; Old ones without shift key; switch back to these
; if new ones above caused any problems
;SC07B & e::Send {End}
;SC07B & a::Send {Home}
;SC07B & f::Send {Right}
;SC07B & b::Send {Left}
;SC07B & p::Send {Up}
;SC07B & n::Send {Down}
;SC07B & d::Send {Delete}

;==========================================================
; Other shortcuts for text editing and others utilities
;==========================================================

; CapsLock + ;: one word left
SC07B &  `;::
	If GetKeyState("LShift","P")
		Send +^{Left}
	Else
		Send ^{Left}
Return
; CapsLock + ': one word right
SC07B & '::
	If GetKeyState("LShift","P")
		Send +^{Right}
	Else
		Send ^{Right}
Return

; Old ones without shift key; switch back to these
; if new ones above caused any problems
;SC07B & `;::Send ^{Left}
;SC07B & '::Send ^{Right}

; CapsLock + H: backspace
SC07B & h::Send {Backspace}

; Alt + H: Ctrl + backspace
!h::Send ^{Backspace}

; Alt + D: delete word to the right
!d::Send ^{Delete}

; Alt + I: Cycle input methods
!i::Send !{LShift}

; Caps Lock + 0: Insert the current date (yyyy-MM-dd)
; Caps Lock + Shift + 0: Insert the current date & time (yyyy-MM-dd hh:mm:ss)
SC07B & 0::
	If GetKeyState("LShift","P")
		FormatTime, TimeString,, yyyy-MM-dd hh:mm:ss
	Else
		FormatTime, TimeString,, yyyy-MM-dd
	Send %TimeString%
Return

SC07B & -::
FormatTime, TimeString,, yyyy-MM-dd
Send %TimeString%
return

SC07B & =::
FormatTime, TimeString,, HH:mm:ss
Send %TimeString%
return

; Ctrl + Alt + /: Mute
Ctrl & /::
	If GetKeyState("LAlt","P")
		SoundSet, +1, , mute  ; Toggle the master mute (set it to the opposite state)
return

Ctrl & -::
	If GetKeyState("LAlt","P")
		SoundSet -2  ; Decrease master volume by 2%
return

Ctrl & =::
	If GetKeyState("LAlt","P")
		SoundSet +2  ; Increase master volume by 2%
return

; Unfortunately, these hotkeys didn't catch on
; SC07B & [::Send {PgUp}
; SC07B & ]::Send {PgDn}
; ![::Send ^{Home}
; !]::Send ^{End}
; SC07B & l::Send ^{Delete}
; SC07B & WheelDown::Send {PgDn}
; SC07B & WheelUp::Send {PgUp}
; !+n::Send !+{Down}        ; Start rectangular selection (down)
; !+p::Send !+{Up}          ; Start rectangular selection (up)

; 
;+SC07B&WheelDown::Send ^{End} ; Compiled, but does not work. Shift + Caps fires ^{End}
;+SC07B & +WheelUp::Send ^{Home}


; Uncomment the following lines replace the mouse middle button click with Ctrl
; (use this when the mouse buttons cannot be customized).
;MButton::Send {Ctrl down}
;MButton Up::Send {Ctrl up}

; 2:41 PM 9/6/2013 Assigned Ctrl + / to "SelectCurrentWord" in VC.
;SC07B & /::Send ^/

;==========================================================
; Enable CapsLock for some "Ctrl + *" Windows shortcuts
;==========================================================

SC07B & z::Send ^z ; CapsLock + Z: undo
SC07B & w::Send ^w ; CapsLock + W: close current window
SC07B & r::Send ^r ; CapsLock + R: reload/refresh
;SC07B & Tab::Send {Ctrl down}{Tab down}
;SC07B & Tab UP::Send {Ctrl up}{Tab up}
SC07B & Tab::Send ^{Tab down}
SC07B & Tab UP::Send ^{Tab up}
;SC07B & +Tab::Send ^{Tab}


;==========================================================
; Code Templates for C++
;==========================================================

!#i::Send if(  ){Enter}{{}{Enter}{}}{Up}{Up}{Right}{Right}{Right}
!#f::Send for( int i=0; i<num_elements; i{+}{+} ){Enter}{{}{Enter}{}}{Enter}{Up}{Up}{Right}{Enter}
!#c::Send class MyClass{Enter}{{}{Enter}public:{Enter}{Tab}MyClass(){{}{}}{Enter}{Tab}~MyClass(){{}{}}{Enter}void MyMethod(){{}{}}{Enter}void Accept( MyVisitor& visitor ){{} visitor.Visit( *this ) {}}{Enter}{}};{Enter}
!#b::Send class BaseClass{Enter}{{}{Enter}public:{Enter}{Tab}BaseClass(){{}{}}{Enter}{Tab}virtual ~BaseClass(){{}{}}{Enter}virtual void VirtualFunction(){{}{}}{Enter}{}};{Enter}
!#d::Send class DerivedClass : public BaseClass{Enter}{{}{Enter}public:{Enter}{Tab}DerivedClass(){{}{}}{Enter}{Tab}~DerivedClass(){{}{}}{Enter}virtual void VirtualFunction(){{}{}}{Enter}{}};{Enter}
!#s::Send switch( variable ){Enter}{{}{Enter}case CASE_0:{Enter}break;{Enter}case CASE_1:{Enter}break;{Enter}default:{Enter}break;{Enter}{}}{Enter}
!#h::Send {#}ifndef __HeaderName_HPP__{Enter}{#}define __HeaderName_HPP__{Enter}{Enter}{Enter}{Enter}{Enter}{#}endif /* __HeaderName_HPP__ */{Enter}
!#g::Send const %clipboard%& Get%clipboard%() const {{} return m_%clipboard%; {}}
!#a::Send std::string name;
!#v::Send std::vector<Type> vec;
;!#t::Send list<Type>::iterator itr;
!#t::Send {#}ifndef __%clipboard%_HPP__{Enter}{#}define __%clipboard%_HPP__{Enter}{Enter}namespace my_namespace{Enter}{{}{Enter}{Enter}class %clipboard%{Enter}{{}{Enter}public:{Enter}{Tab}%clipboard%(){{}{}}{Enter}{Tab}~%clipboard%(){{}{}}{Enter}{}};{Enter}{Enter}{}} // namespace my_namespace{Enter}{Enter}{#}endif /* __%clipboard%_HPP__ */{Enter}
!#p::Send boost::shared_ptr<ClassName> ptr;^{Left}^{Left}^{Left}+^{Left}
!#o::Send shared_ptr<ClassName> ptr( new ClassName );^{Left}^{Left}^{Left}^{Left}^{Left}^{Left}^{Left}+^{Left}
!#z::Send const std::string& name
!#y::Send for( itr = container.begin();{Enter}itr {!}= container.end();{Enter}itr{+}{+} ){Enter}{{}{Enter}{}}{Enter}{Up}{Up}{Right}{Enter}
!#6::Send for( list<Type>::iterator itr = container.begin(); itr {!}= container.end(); itr{+}{+} ){Enter}{{}{Enter}{}}{Enter}{Up}{Up}{Right}{Enter}
!#m::Send int main( int argc, char *argv[] ){Enter}{{}{Enter}return 0;{Enter}{}}{Enter}
!#n::Send namespace my_namespace{Enter}{{}{Enter}{Enter}{}} // my_namespace{Enter}
!#e::Send enum EnumName{Enter}{{}{Enter}ENUM_0,{Enter}ENUM_1,{Enter}ENUM_2,{Enter}NUM_ENUMS{Enter}{}};{Enter}
!#0::Send {#}ifndef __ClassName_HPP__{Enter}{#}define __ClassName_HPP__{Enter}{Enter}{Enter}class ClassName{Enter}{{}{Enter}public:{Enter}{Tab}ClassName(){{}{}}{Enter}{Tab}~ClassName(){{}{}}{Enter}{}};{Enter}{Enter}{Enter}{Enter}{#}endif /* __ClassName_HPP__ */{Enter}

; -------------- lines --------------
;!#=::Send ====================  ====================^{Left}{Left}


;==========================================================
; Open Directory
;==========================================================

; CapsLock + ~: open the home directory of the current local machine
SC07B & `::
	EnvGet, home_dir, HOME
    Run Explorer %home_dir%
Return

; Win + ~: home directory (Explorer)
#`::
	EnvGet, home_dir, HOME
    Run Explorer %home_dir%
Return

SC07B & SC029::
	EnvGet, home_dir, HOME
    Run Explorer %home_dir%
Return


;==========================================================
; Application Launchers
;==========================================================

; Win + N: Notepad++
#n::Run "C:\Program Files\Notepad++\notepad++.exe"

; Win + W: WinMerge
;#w::Run "C:\awww\WinMergeU.exe"

; Win + C: Google Chrome
;#c::Run "C:\Program Files\awww\chrome.exe"

; Win + X: eject removable media
;#x::Run "C:\Program Files\awww\FreeEject" G:

; Win 
;SC07B & LWin::
;	If GetKeyState("b","P")
;		Run "C:\Program Files\WinMerge\WinMergeU.exe"
;Return


; Win + 1: Appcatlion 1
;#1::Run "C:\Program Files\app1\app1.exe"

; Win + 2: Appcatlion 2
;#2::Run "C:\Program Files\app2\app2.exe"

; Win + 3: Appcatlion 3
;#3::Run "C:\Program Files\app3\app3.exe"

; Win + 4: Appcatlion 4
;#4::Run "C:\Program Files\app4\app4.exe"

;Ctrl & e::End
Numpad0 & Numpad2::Run Notepad
SetCapsLockState, AlwaysOff


;==========================================================
; Application-specific Shortcuts
;==========================================================

; Notepad++
; #IfWinActive ahk_class Notepad++
; ...
; #IfWinActive

; File Explorer
#IfWinActive ahk_class CabinetWClass
;!SC07B & p::Send !{Up}
!u::Send !{Up}             ; Alt + U: Go to the parent directory
!WheelUp::Send !{Up}       ; Alt + Mouse Wheel Up: Go to the parent directory
SC07B & q::Send !{Enter}   ; Caps Lock + Q: Properties
SC07B & 4::Send ^+n        ; Caps Lock + 4: Create a new folder
!2::Send {F2}              ; Alt + 2: Rename files/directories
!3::Send ^+n               ; Alt + 3: Create a new folder
+WheelUp::Send !{Left}     ; Navigate backward
+WheelDown::Send !{Right}  ; Navigate forward

; Enter: do specific action based on the extension or fall back to enter
;Enter::
;	sel := Explorer_GetSelected()
;
;	StringGetPos, pos, sel, .jpg
;	if pos >= 0
;	{
;		Run C:\Program Files\awww\awww.exe "%sel%"
;		return
;	}
;
;Send {Enter}
;return

; Caps Lock + G: launch application (as configured here)
SC07B & g::
	sel := Explorer_GetSelected()

	if InStr(FileExist(sel), "D")
	{
		Run cmd /K cd "%sel%"
		return
	}

	StringGetPos, pos, sel, .7z
	if pos >= 0
	{
		Run D:\programs\archivers\7-Zip\7z.exe x "%sel%"
		return
	}
	StringGetPos, pos, sel, .mp4
	if pos >= 0
	{
		; Convert the mp4 to an mp3 file (extract the audio track and save it as an mp3 file).
		Run convert_mp4_to_mp3.bat "%sel%"
		return
	}
	Run D:\programs_x86\text_editors\Notepad++\notepad++.exe %sel%
return

; CapsLock + T: Git diff
; SC07B & t::
; 	sel := Explorer_GetSelected()
; 	Run "awww.exe" /command:diff /path:"%sel%"
; return

!c::
	clipboard := Explorer_GetSelected()
return
#IfWinActive

; WinMerge
#IfWinActive ahk_class WinMergeWindowClassW
SC07B & r::Send {F5}         ; Refresh
!WheelDown::Send !{Down}
!WheelUp::Send !{Up}
!n::Send !{Down}
!p::Send !{Up}
#IfWinActive

; Jasc Paint Shop Pro 7
#IfWinActive , Jasc Paint Shop Pro
^W::Send ^{F4}               ; Ctrl + W: close the current window
SC07B & w::Send ^{F4}        ; CapsLock + W: close the current window
SC07B & i::Send ^+i          ; CapsLock + I: invert the selection
#IfWinActive

; Word
#IfWinActive ahk_class OpusApp
!WheelUp::Send ^>            ; Alt + mouse wheel up: increase the font size
!WheelDown::Send ^<          ; Alt + mouse wheel down: decrease the font size
#IfWinActive

; PowerPoint
#IfWinActive ahk_class PP12FrameClass
!WheelUp::Send ^+>           ; Alt + mouse wheel up: increase the font size
!WheelDown::Send ^+<         ; Alt + mouse wheel down: decrease the font size
SC07B & q::Send {AppsKey}o   ; CapsLock + Q: ???
!#k::Send {AppsKey}ia        ; Alt + Win + K: Insert a row above
!#j::Send {AppsKey}ib        ; Alt + Win + J: Insert a row below
!#c::Send {Alt}jdaac         ; Alt + Win + C: Align objects vertically at the center
!#h::Send {Alt}jdaav         ; Alt + Win + H: Equal spacing (horizontal)
!#v::Send {Alt}jdaav         ; Alt + Win + V: Equal spacing (vertical)
!#l::Send {Alt}jdaal         ; Alt + Win + L: Align left
!#r::Send {Alt}jdaar         ; Alt + Win + R: Align right
#IfWinActive

; Excel
#IfWinActive ahk_class XLMAIN
SC07B & 2::Send {Alt}hmm
#IfWinActive

; Google Chrome
#IfWinActive ahk_class Chrome_WidgetWin_1
!.::Send {F3}                ; Find next
!,::Send +{F3}               ; Find previous
SC07B & .::Send {F3}         ; Find next
SC07B & ,::Send +{F3}        ; Find previous
!WheelDown::Send {F3}        ; Find next
!WheelUp::Send +{F3}         ; Find previous
SC07B & WheelDown::Send {F3} ; Find next
SC07B & WheelUp::Send +{F3}  ; Find previous
SC07B & s::Send ^f           ; Search
SC07B & t::Send ^t           ; Open a new tab
SC07B & 2::Send {Browser_Back}
SC07B & 3::Send {Browser_Forward}
+WheelUp::Send {Browser_Back}
+WheelDown::Send {Browser_Forward}
#IfWinActive

; WMP
#IfWinActive , Windows Media Player
SC07B & Space::Send ^p                ; Pause
#IfWinActive

; Console Windows
; #IfWinActive ahk_class ConsoleWindowClass
; ...
; #IfWinActive
;The script above did not work with #IfWinActive ahk_class HwndWrapper ... #IfWinActive
;Windows Spy of AutoHotkey shows an id string that follows 'HwndWrapper', the window name of the VC's text editor
;Might have to specify the id as well, but the every window has its own id.
;Worked without #IfWinActive ahk_class HwndWrapper ... #IfWinActive, but it would affect other applications.

;F1::
;	path := Explorer_GetPath()
;	all := Explorer_GetAll()
;	sel := Explorer_GetSelected()
;	MsgBox % path
;	MsgBox % all
;	MsgBox % sel
;return

SetStoreCapslockMode, On




