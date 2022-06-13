
; -----------------------------------------------------------
; Name: Obsidian.ahk
; Description: This file contains the automation process of git for Obsidian App
; Source url: https://github.com/hsayed21/Obsidian.ahk
; Author: Hamada Sayed(x00h)[https://github.com/hsayed21]
; License: Copyright (c) 2022 hsayed21
; -----------------------------------------------------------

#singleInstance, force
#Persistent
DetectHiddenWindows, On
SetWorkingDir, F:\7S\IS\ZNotes\My_Plan_2020\Courses_Note_Offline
isPressed := False
freeze := 0
App = Obsidian.exe ; Change the name to application you are checking.
DetectAppStartOrClose()
return


; ###### Detect App [ Start | Close ]
; Action When Start App
StartApp()
{
	global freeze
	If (freeze = 0)
	{
		SetTimer, noShutdown, 40
		freeze = 1
	}

	Gosub, CheckInternetConnection
	Try
	{
		RunWaitOne("git pull",1)
	}
	Catch e
	{
		errorMsg := "Error on line " . e.line . ": " . e.message . "`n"
		FileAppend, errorMsg, %A_ScriptDir%\errorlog.txt
		Gosub, CheckInternetConnection
		StartApp()
	}	
}

; Action When Close App
CloseApp()
{
	StartApp()
	Gosub, CheckInternetConnection
	
	RunWaitOne("git add .")
	Sleep 200
	FormatTime, TimeString,,yyyy-MM-dd HH:mm:ss
	commit = Last Sync: %TimeString%
	commit := """" commit """"
	RunWaitOne("git commit -m " commit)
	Sleep 200
	Try
	{
		RunWaitOne("git push",1)
	}
	Catch e
	{
		errorMsg := "Error on line " . e.line . ": " . e.message . "`n"
		FileAppend, errorMsg, %A_ScriptDir%\errorlog.txt
		Gosub, CheckInternetConnection
		CloseApp()
	}	

	Sleep, 5000
	SetTimer, noShutdown, Off
	isPressed := False
	freeze := 0
}

; Listen On Messages Closing , Starting
DetectAppStartOrClose()
{
    ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms644947(v=vs.85).aspx
    Static MsgNumber := DllCall("User32.dll\RegisterWindowMessageW", "Str","SHELLHOOK", "UInt")

    ; https://msdn.microsoft.com/en-us/library/windows/desktop/ms644989(v=vs.85).aspx
    If (!MsgNumber || !DllCall("User32.dll\RegisterShellHookWindow", "Ptr", A_ScriptHwnd))
        Return FALSE
    OnMessage(MsgNumber, "CaptureShellMessage")
    Return TRUE
}

DeregisterShellHookWindow()
{
    Return DllCall("User32.dll\DeregisterShellHookWindow", "Ptr", A_ScriptHwnd)
		; https://msdn.microsoft.com/en-us/library/windows/desktop/ms644979(v=vs.85).aspx
} 

CaptureShellMessage(wParam, lParam, msg, hwnd)
{
	global App
  If ( wParam = 1 ) ;  HSHELL_WINDOWCREATED := 1
  {
    WinGet, proName, ProcessName, ahk_id %lParam%
    If ( proName = App ) 
    {
      ; MsgBox, Obsidian App is Started
			StartApp()
    }
  }Else if ( wParam = 2){ ; HSHELL_WINDOWDESTROYED := 2
    WinGet, proName, ProcessName, ahk_id %lParam%
    If ( proName = App ) 
    {
      ; MsgBox, Obsidian App is Closed
			CloseApp()
    }
  }
}


; ###### Check Internet Connection
CheckInternetConnection:
Loop
{
	if (ping("8.8.8.8") = "Online")
		return
	Sleep, 2000
}
return

ping(host) 
{	
	colPings := ComObjGet( "winmgmts:" ).ExecQuery("Select * From Win32_PingStatus where Address = '" host "'")._NewEnum

	While colPings[objStatus]
	Return ((oS:=(objStatus.StatusCode="" or objStatus.StatusCode<>0)) ? "Offline" : "Online" )
}


; ###### CMD
RunWaitOne(command, t:=0) {
  DetectHiddenWindows On
  Run %ComSpec%,, Hide, pid
  WinWait ahk_pid %pid%
  DllCall("AttachConsole", "UInt", pid)

  shell := ComObjCreate("WScript.Shell")
  exec := shell.Exec(ComSpec " /C " command)
  DllCall( "FreeConsole" )
  if (!t)
		return exec.StdOut.ReadAll()

	res := exec.StdErr.ReadAll()
	If (res = "" || res = "Everything up-to-date`n" || res = "Already up to date.`n")
		return
	else
		throw { what: "error", file: A_LineFile, line: A_LineNumber, message: exec.StdErr.ReadAll() }
  ; return exec.StdErr.ReadAll()
}


; ###### Detect Click On Windows Power Button  
noShutdown:
MouseGetPos, x, y, hwnd
obj:=Acc_GetInfoUnderCursor()
WinGet, ProcessN, ProcessName, ahk_id %hwnd%
If (ProcessN = "StartMenuExperienceHost.exe" && obj.Name = "Power")
    isPressed := True
else
    isPressed := False
return


#If, isPressed
LButton::
RButton::
Space::
Enter::
NumpadEnter::return
return


; ###### ACC Lib
Acc_GetInfoUnderCursor()
{
    Acc:=Acc_ObjectFromPoint(child)
    Try Name:="", Name:=Acc.accName(child)
    Try Value:="", Value:=Acc.accValue(child)
    return {Name:Name, Value:Value}
}

Acc_ObjectFromPoint(ByRef _idChild_ = "", x = "", y = "")
{
    Acc_Init()
    If  DllCall("oleacc\AccessibleObjectFromPoint", "Int64", x==""||y==""
    ? 0*DllCall("GetCursorPos","Int64*",pt)+pt:x&0xFFFFFFFF|y<<32, "Ptr*", pacc
    , "Ptr", VarSetCapacity(varChild,8+2*A_PtrSize,0)*0+&varChild)=0
        return ComObjEnwrap(9,pacc,1), _idChild_:=NumGet(varChild,8,"UInt")
}

Acc_Init()
{
    Static h
    If (!h)
        h:=DllCall("LoadLibrary","Str","oleacc","Ptr")
}

; ###### Hotkey
; Stop Detect Shutdown
^Esc::
SetTimer, noShutdown, Off
isPressed := False
freeze := 0
return

; Exit Script
!Esc::ExitApp
