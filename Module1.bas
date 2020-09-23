Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type KBDLLHOOKSTRUCT
    vkCode As Long        'value of the key you pressed
    scanCode As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Public Const WH_KEYBOARD = 2
Public Const WH_KEYBOARD_LL = 13
Public Const HC_ACTION = 0
Public Const VK_DELETE = &H2E
Public KeyBoardHook As Long

'notice that the lparam is passed byref.  this is becuase lparam will be a pointer to a structure _
This is a user defined routine it will recieve messages from the keyboard before they are _
entered into the calling threads que.
Public Function LowLevelKeyboardProc(ByVal nCode As Long, ByVal wParam As Long, lParam As Long) As Long

Dim xpInfo As KBDLLHOOKSTRUCT

    If nCode = HC_ACTION Then
        CopyMemory xpInfo, lParam, Len(xpInfo) 'copy the structure from lParam to xpinfo
        If xpInfo.vkCode = VK_DELETE Then 'the delete key was hit
            If xpInfo.flags = 1 Then 'the delete key is in the down state
                LowLevelKeyboardProc = -1 'In essence this will prevent the delete key from being recognized _
                                           in the calling threads cue
            End If
        End If
    Else
        'this will be called if there are multiple hooks made to the keyboard
        LowLevelKeyboardProc = CallNextHookEx(KeyBoardHook, nCode, wParam, lParam) 'this will be called if there are multiple hooks made to the keyboard
    End If

End Function

