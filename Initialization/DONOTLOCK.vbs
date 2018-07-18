
'====================================================================================
'To Press NumLock Key ON if it is OFF 

'====================================================================================
 Dim x, oWshShell 
 x = IsNumlocked 
 If x = 0 Then 
   set oWshShell = CreateObject("WScript.Shell") 
  oWshShell.SendKeys "{NUMLOCK}" 
 End If 
 Function IsNumLocked() 
  Dim oWrd 
  Set oWrd = CreateObject("Word.Application") 
  IsNumLocked = oWrd.Numlock 
  oWrd.Application.Quit True 
 End Function 



'====================================================================================
'Code to  make computer keep unlock to avoid failuer due to autolock during batch execution 

'====================================================================================

set wsc = CreateObject("WScript.Shell")

Do
wscript.Sleep (60*1000)
wsc.SendKeys ("{SCROLLLOCK 2}")
Loop
