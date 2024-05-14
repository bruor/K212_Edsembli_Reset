;This script MUST BE run as administrator 
;in order for keystrokes and cursor movement to work
 
#include <Array.au3>
#include <GuiToolBar.au3>
 
;Restart IIS
RunWait("iisreset")
Sleep(2000)
 
;right click systray icon
Local $hSysTray_Handle
$sSearchtext='Running - CMD.NET Process Management Service';
$iButton=Get_SysTray_IconText($sSearchtext)
_GUICtrlToolbar_ClickButton($hSysTray_Handle, $iButton, "right", True, 1)
 
;press down then enter to run discover
Sleep(200)
Send("{DOWN}")
Sleep(200)
Send("{ENTER}")
 
 
Func Get_SysTray_IconText($sSearch)
    For $i = 1 To 99
        ; Find systray handles
        $hSysTray_Handle = ControlGetHandle('[Class:Shell_TrayWnd]', '', '[Class:ToolbarWindow32;Instance:' & $i & ']')
        If @error Then
            ;MsgBox(16, "Error", "System tray not found")
            ExitLoop
        EndIf
 
        ; Get systray item count
        Local $iSysTray_ButCount = _GUICtrlToolbar_ButtonCount($hSysTray_Handle)
        If $iSysTray_ButCount = 0 Then
            ;MsgBox(16, "Error", "No items found in system tray")
            ContinueLoop
        EndIf
 
        Local $aSysTray_ButtonText[$iSysTray_ButCount]
 
        ; Look for wanted tooltip
        For $iSysTray_ButtonNumber = 0 To $iSysTray_ButCount - 1
	    ;$ButtonText = _GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSysTray_ButtonNumber)
	    ;ConsoleWrite(@CR & "ButtonText" & $ButtonText)
            If $sSearch= _GUICtrlToolbar_GetButtonText($hSysTray_Handle, $iSysTray_ButtonNumber) Then _
            Return SetError(0, $i, $iSysTray_ButtonNumber)
        Next
    Next
    Return SetError(1, -1, -1)
 
EndFunc   ;==>Get_SysTray_IconText
 
Func PrintList($List)
    If IsArray($List) Then
        Local $txt = ""
        For $i = 0 to UBound($List) -1
            $txt = $txt & "," & $List[$i]
        Next
        Local $out = StringMid($txt,2)
        Global $Result = "[" & $out & "]"
        ConsoleWrite($Result)
        Return $Result
    Else
        MsgBox(0, "List Error", "Variable is not an array or a list")
    EndIf
EndFunc   ;==>PrintList
