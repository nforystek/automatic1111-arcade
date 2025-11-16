Attribute VB_Name = "modKioskApp"
Option Explicit

Private Enum ApiConsts
    WM_POWER = &H48
    PWR_SUSPENDREQUEST = 1
    WM_POWERBROADCAST = &H218
    PBT_APMQUERYSUSPEND = 0
    PWR_FAIL = -1
    DENY_QUERY = &H424D5144

    IDX_WNDPROC = -4
End Enum

Private Type SYSTEM_POWER_STATUS
    ACLineStatus        As Byte
    BatteryFlag         As Byte
    BatteryLifePercent  As Byte
    Reserved1           As Byte
    BatteryLifeTime     As Long
    BatteryFullLifeTime As Long
End Type

' Execution state flags
Private Const ES_CONTINUOUS As Long = &H80000000
Private Const ES_SYSTEM_REQUIRED As Long = &H1
Private Const ES_DISPLAY_REQUIRED As Long = &H2

Private Const SPI_SCREENSAVERRUNNING As Long = &H61

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function SetThreadExecutionState Lib "kernel32.dll" (ByVal esFlags As Long) As Long

'local variables
Private SysPwrStat      As SYSTEM_POWER_STATUS
Private prevProcAddress As Long
Private hWndActive      As Long

Public Sub DoLockOut()
    #If VBIDE = -1 Then
    If (Not frmMain.Visible) And GetSecureState Then
        Shell "rundll32.exe user32.dll, LockWorkStation"
    End If
    #End If
End Sub

Private Function GetSecureState() As Boolean
    GetSecureState = CBool(RegReadString(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaverIsSecure", True) Or _
                            RegReadString(HKEY_CURRENT_USER, "Control Panel\Desktop", "ScreenSaverIsSecure", False))
End Function


' Call this when you want to prevent sleep/screensaver
Public Sub PreventSleep()
    SetThreadExecutionState ES_CONTINUOUS Or ES_SYSTEM_REQUIRED Or ES_DISPLAY_REQUIRED
End Sub

' Call this when you're done (e.g., on form unload or when playback ends)
Public Sub AllowSleep()
    SetThreadExecutionState ES_CONTINUOUS
End Sub


Public Sub ActivatePowerMonitor()

    ShowCursor 0
    #If VBIDE = 0 Then
    
    If prevProcAddress = 0 Then
        prevProcAddress = SetWindowLong(frmMain.hwnd, IDX_WNDPROC, AddressOf MessageHook)
        hWndActive = frmMain.hwnd
    End If

    Dim dTask As Integer
    Dim junk As Boolean
     
    dTask = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 1, junk, 0)
    
    PreventSleep

    #End If
End Sub

Public Sub DeactivatePowerMonitor()
    
    #If VBIDE = 0 Then
    AllowSleep
    
    Dim eTask As Integer
    Dim junk As Boolean
     
    eTask = SystemParametersInfo(SPI_SCREENSAVERRUNNING, 0, junk, 0)
    

    If prevProcAddress Then
        SetWindowLong hWndActive, IDX_WNDPROC, prevProcAddress
        prevProcAddress = 0
    End If
    #End If
    ShowCursor 1
End Sub

Private Function IsOnMainsSupply() As Boolean

    GetSystemPowerStatus SysPwrStat
    IsOnMainsSupply = (SysPwrStat.ACLineStatus = 1)

End Function

Private Function MessageHook(ByVal hwnd As Long, ByVal nMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    MessageHook = CallWindowProc(prevProcAddress, hwnd, nMsg, wParam, lParam)
    If IsOnMainsSupply Then
        Select Case True
          Case nMsg = WM_POWER And wParam = PWR_SUSPENDREQUEST
            MessageHook = PWR_FAIL
          Case nMsg = WM_POWERBROADCAST And wParam = PBT_APMQUERYSUSPEND
            MessageHook = DENY_QUERY
        End Select
    End If

End Function

