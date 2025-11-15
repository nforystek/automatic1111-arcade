Attribute VB_Name = "modProcess"
#Const modProcess = -1
Option Explicit
'TOP DOWN
Option Compare Binary

Option Private Module

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
 
Private Const PROCESS_TERMINATE As Long = &H1

Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwAccess As Long, ByVal fInherit As Integer, ByVal hObject As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Public Sub KillVisibleProcesses()
    EnumWindows AddressOf EnumWindowsProc, 0
End Sub

Private Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Long
    If IsWindowVisible(hwnd) Then

        Dim PID As Long
        GetWindowThreadProcessId hwnd, PID
        If PID <> 0 And PID <> GetCurrentProcessId Then
            Dim hProc As Long
            hProc = OpenProcess(PROCESS_TERMINATE, 0, PID)
            If hProc <> 0 Then
                TerminateProcess hProc, 0
                CloseHandle hProc
            End If
        End If

    End If
    EnumWindowsProc = 1 ' Continue enumeration
End Function

Public Function OpenWebsite(ByVal WebSite As String, Optional ByVal Silent As Boolean) As Boolean
    OpenWebsite = (RunFile(WebSite) <> 0)
End Function
Public Function RunFile(ByVal File As String, Optional ByVal Params As String = "", Optional ByVal FocusPID As Long = 1) As Long

    If Not PathExists(File, True) Then
        RunFile = ShellExecute(0, "open", File, Params, 0&, FocusPID)
    Else
        RunFile = ShellExecute(0, "open", GetFileName(File), Params, GetFilePath(File), FocusPID)
    End If

End Function

Public Function RunProcess(ByVal Path As String, Optional ByVal Params As String = "", Optional ByVal Focus As Integer = vbNormalFocus) As Long

    RunProcess = Shell(Trim(Path & " " & Params), Focus)

End Function

Public Function ProcessRunning(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Long
    On Local Error GoTo catch

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim I As Integer
    Dim cnt As Long
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0
    
    Do While rProcessFound
        I = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, I - 1))

        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
               
                cnt = cnt + 1
        ElseIf IsNumeric(EXEorPID) Then
            If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                cnt = cnt + 1
            End If
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)

    Loop
    Call CloseHandle(hSnapshot)

    ProcessRunning = cnt
    Exit Function
catch:
    Err.Clear
End Function


Public Function RunningProcessCount(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Long

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim I As Integer
    Dim cnt As Long

    Const TH32CS_SNAPPROCESS As Long = 2&
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0

    Do While rProcessFound
        I = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, I - 1))
        
        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
            
            RunningProcessCount = RunningProcessCount + 1


        End If

        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    
    Call CloseHandle(hSnapshot)

End Function


Public Function IsProccessEXERunning(ByVal EXE As Variant, Optional ByVal ExactMatch As Boolean = True) As Boolean

    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim I As Integer
    Dim cnt As Long

    Const TH32CS_SNAPPROCESS As Long = 2&
  
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    cnt = 0

    Do While rProcessFound
        I = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase(Left(uProcess.szexeFile, I - 1))
        
        If ProcessCheck(szExename, EXE, ExactMatch) Then
                
            IsProccessEXERunning = True

        End If

        rProcessFound = ProcessNext(hSnapshot, uProcess)

    Loop
    
    Call CloseHandle(hSnapshot)

End Function

Public Function KillApp(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim I As Integer
    On Local Error GoTo Finish
    appCount = 0
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
    Do While rProcessFound
        I = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, I - 1))

        If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
                
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(1, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
            
        ElseIf IsNumeric(EXEorPID) Then
            If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                KillApp = True
                appCount = appCount + 1
                myProcess = OpenProcess(1, False, uProcess.th32ProcessID)
                AppKill = TerminateProcess(myProcess, exitCode)
                Call CloseHandle(myProcess)
            End If
        End If
       
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop


    Call CloseHandle(hSnapshot)
Finish:
End Function

Private Function ColExists(ByRef col As Collection, ByVal Val As String) As Boolean
    If col.Count > 0 Then
        Dim I As Long
        For I = 1 To col.Count
            If col(I) = Val Then
                ColExists = True
                Exit Function
            End If
        Next
    End If
End Function

Public Function KillSubApps(ByVal EXEorPID As Variant, Optional ByVal ExactMatch As Boolean = True) As Boolean

    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim I As Integer
    On Local Error GoTo Finish
    appCount = 0
    Dim col As New Collection
    
    Const TH32CS_SNAPPROCESS As Long = 2&
    
    Debug.Print "KillSubApps: " & EXEorPID
    Dim foundAdd As Boolean
    
    foundAdd = True
    
    Do While foundAdd
        foundAdd = False
        
        uProcess.dwSize = Len(uProcess)
        hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
        rProcessFound = ProcessFirst(hSnapshot, uProcess)
    
        Do While rProcessFound
            I = InStr(1, uProcess.szexeFile, Chr(0))
            szExename = LCase$(Left$(uProcess.szexeFile, I - 1))
            If ProcessCheck(szExename, EXEorPID, ExactMatch) Then
                If Not ColExists(col, uProcess.th32ProcessID) Then
                    col.Add CStr(uProcess.th32ProcessID), Replace(Replace(CStr(EXEorPID), " ", "_"), ".", "_")
                    foundAdd = True
                End If
            ElseIf IsNumeric(EXEorPID) Then
                If (uProcess.th32ProcessID = CLng(EXEorPID)) Then
                    If Not ColExists(col, uProcess.th32ProcessID) Then
                        col.Add CStr(uProcess.th32ProcessID), Replace(Replace(CStr(EXEorPID), " ", "_"), ".", "_")
                        foundAdd = True
                    End If
                ElseIf (uProcess.th32ParentProcessID = CLng(EXEorPID)) Or _
                    ColExists(col, uProcess.th32ParentProcessID) Then
                    If Not ColExists(col, uProcess.th32ProcessID) Then
                        col.Add CStr(uProcess.th32ProcessID)
                        foundAdd = True
                    End If
                End If
            ElseIf ColExists(col, uProcess.th32ParentProcessID) Then
                If Not ColExists(col, uProcess.th32ProcessID) Then
                    col.Add CStr(uProcess.th32ProcessID)
                    foundAdd = True
                End If
            End If
    
            rProcessFound = ProcessNext(hSnapshot, uProcess)
        Loop
    
        Call CloseHandle(hSnapshot)
    
    Loop

    If col.Count > 0 Then
        For I = 1 To col.Count

            If col(I) <> col(Replace(Replace(CStr(EXEorPID), " ", "_"), ".", "_")) Then
              '  If CLng(col(I)) <> GetCurrentProcessId Then
                    Debug.Print "    SubApp: " & CLng(col(I))
                    KillSubApps = True
                    appCount = appCount + 1
                    myProcess = OpenProcess(1, False, CLng(col(I)))
                    AppKill = TerminateProcess(myProcess, exitCode)
                    Call CloseHandle(myProcess)
               ' End If
            End If
            
        Next
    End If
    
Finish:
End Function


Private Function ProcessCheck(ByVal szExename As String, ByVal EXEorPID As Variant, ByVal ExactMatch As Boolean) As Boolean
    ProcessCheck = ( _
             ( _
               ( _
                 (Right(szExename, Len(EXEorPID)) = LCase(EXEorPID)) Or _
                 (Right(LCase(EXEorPID), Len(szExename)) = szExename) _
                ) _
                Or _
                ( _
                  (Left(szExename, Len(EXEorPID)) = LCase(EXEorPID)) Or _
                  (Left(LCase(EXEorPID), Len(szExename)) = szExename) _
                ) _
              ) _
              And (Not ExactMatch) _
            ) _
           Or _
           ( _
             ( _
                ( _
                  (LCase(szExename) = LCase(EXEorPID)) Or _
                  ((InStr(EXEorPID, "\") = 0 And InStr(EXEorPID, "/") = 0) And (LCase$(GetFileName(szExename)) = LCase(EXEorPID))) Or _
                  ((InStr(szExename, "\") = 0 And InStr(szExename, "/") = 0) And (LCase$(szExename) = LCase(GetFileName(EXEorPID)))) _
                ) _
              ) _
              And ExactMatch _
            )
End Function




