Attribute VB_Name = "modFOnt"

''OW TO USE THIS MODULE


''system wide install/uninstall (requires admin)
'If InstallFontSystemWide("C:\Temp\MyFont.ttf") Then
'    MsgBox "Installed!"
'Else
'    MsgBox "Failed."
'End If
'Call UninstallFontSystemWide("MyFont.ttf")


''private install/uninstall (no admin required)
'If InstallFontPrivate("C:\Temp\MyFont.ttf") Then
'    MsgBox "Private font loaded."
'End If
'Call UninstallFontPrivate("C:\Temp\MyFont.ttf")






' ================================
'   modFontInstall.bas
'   Reusable Font Installation Module
'   Works on Windows XP ? Windows 11
' ================================

Option Explicit

' --- API DECLARATIONS ---

Private Declare Function AddFontResource Lib "gdi32" _
    Alias "AddFontResourceA" (ByVal lpFileName As String) As Long

Private Declare Function AddFontResourceEx Lib "gdi32" _
    Alias "AddFontResourceExA" (ByVal lpFileName As String, _
    ByVal fl As Long, ByVal pdv As Long) As Long

Private Declare Function RemoveFontResource Lib "gdi32" _
    Alias "RemoveFontResourceA" (ByVal lpFileName As String) As Long

Private Declare Function RemoveFontResourceEx Lib "gdi32" _
    Alias "RemoveFontResourceExA" (ByVal lpFileName As String, _
    ByVal fl As Long, ByVal pdv As Long) As Long

Private Declare Function SendMessage Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const HWND_BROADCAST = &HFFFF&
Private Const WM_FONTCHANGE = &H1D

Private Const FR_PRIVATE = &H10
Private Const FR_NOT_ENUM = &H20


' ================================
'   PUBLIC FUNCTIONS
' ================================

' --------------------------------
' Install a font system-wide
' Requires admin rights
' --------------------------------
Public Function InstallFontSystemWide(ByVal SourcePath As String) As Boolean
    Dim fontsFolder As String
    Dim destPath As String
    Dim result As Long

    fontsFolder = Environ$("WINDIR") & "\Fonts\"
    destPath = fontsFolder & GetFileName(SourcePath)

    On Error Resume Next
    FileCopy SourcePath, destPath
    If Err.Number <> 0 Then
        InstallFontSystemWide = False
        Exit Function
    End If
    On Error GoTo 0

    result = AddFontResource(destPath)

    If result > 0 Then
        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
        InstallFontSystemWide = True
    Else
        InstallFontSystemWide = False
    End If
End Function


' --------------------------------
' Install a private font
' Does NOT require admin rights
' Only visible to your application
' --------------------------------
Public Function InstallFontPrivate(ByVal FontPath As String) As Boolean
    Dim result As Long

    result = AddFontResourceEx(FontPath, FR_PRIVATE Or FR_NOT_ENUM, 0)

    InstallFontPrivate = (result > 0)
End Function


' --------------------------------
' Uninstall a system-wide font
' --------------------------------
Public Function UninstallFontSystemWide(ByVal FontFileName As String) As Boolean
    Dim fontsFolder As String
    Dim FullPath As String
    Dim result As Long

    fontsFolder = Environ$("WINDIR") & "\Fonts\"
    FullPath = fontsFolder & FontFileName

    result = RemoveFontResource(FullPath)

    If result > 0 Then
        SendMessage HWND_BROADCAST, WM_FONTCHANGE, 0, 0
        UninstallFontSystemWide = True
    Else
        UninstallFontSystemWide = False
    End If
End Function


' --------------------------------
' Uninstall a private font
' --------------------------------
Public Function UninstallFontPrivate(ByVal FontPath As String) As Boolean
    Dim result As Long

    result = RemoveFontResourceEx(FontPath, FR_PRIVATE Or FR_NOT_ENUM, 0)

    UninstallFontPrivate = (result > 0)
End Function


' ================================
'   HELPER FUNCTIONS
' ================================

Private Function GetFileName(ByVal FullPath As String) As String
    Dim pos As Long
    pos = InStrRev(FullPath, "\")
    If pos > 0 Then
        GetFileName = Mid$(FullPath, pos + 1)
    Else
        GetFileName = FullPath
    End If
End Function

