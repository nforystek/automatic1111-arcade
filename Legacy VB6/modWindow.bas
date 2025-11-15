Attribute VB_Name = "modWindow"
#Const modWindow = -1
Option Explicit

Public Const GWL_WNDPROC = (-4)

Private Const DecaultClassName = "Message"

Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Function WindowInitialize(Optional ByVal lpWndProc As Long = -1, Optional ByVal ClassName As String = "", Optional ByVal WindowName As String = "") As Long
    If ClassName = "" Then
        ClassName = DecaultClassName
    End If
    Dim hwnd As Long
    
    hwnd = CreateWindowEx(ByVal 0&, ClassName, WindowName, 0&, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)
    
    If lpWndProc = -1 Then
        SetWindowLong hwnd, GWL_WNDPROC, AddressOf WindowDefaultProc
    ElseIf lpWndProc > 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, lpWndProc
    End If
    WindowInitialize = hwnd
    
End Function

Public Sub WindowTerminate(ByRef hwnd As Long)
    If hwnd <> 0 Then
        DestroyWindow hwnd
        hwnd = 0
    End If

End Sub

Public Function WindowDefaultProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    WindowDefaultProc = DefWindowProc(hwnd, uMsg, wParam, lParam)

End Function






