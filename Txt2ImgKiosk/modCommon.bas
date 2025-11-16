Attribute VB_Name = "modCommon"
#Const modCommon = -1
Option Explicit

Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, ByRef bWow64Process As Boolean) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Public Function ArraySize(InArray, Optional ByVal InBytes As Boolean = False) As Long
On Error GoTo dimerror

    Static dimcheck As Long

    If UBound(InArray) = -1 Or LBound(InArray) = -1 Then
        ArraySize = 0
    Else
        ArraySize = (UBound(InArray) + -CInt(Not CBool(-LBound(InArray)))) * IIf(InBytes, LenB(InArray(LBound(InArray))), 1)
    End If
    Exit Function
startover:
    Err.Clear
    On Error GoTo -1
    On Error GoTo 0
    On Error GoTo dimerror
    If UBound(InArray, dimcheck) = -1 Or LBound(InArray, dimcheck) = -1 Then
        ArraySize = 0
    Else
        ArraySize = (UBound(InArray, dimcheck) + -CInt(Not CBool(-LBound(InArray, dimcheck)))) * IIf(InBytes, LenB(InArray(LBound(InArray, dimcheck), LBound(InArray, dimcheck - 1))), 1)
    End If
    
    Exit Function
dimerror:
    If dimcheck = 0 Then
        dimcheck = 2
        Err.Clear
        GoTo startover
    End If
    ArraySize = 0
End Function

Public Function Convert(Info)
    Dim N As Long
    Dim out() As Byte
    Dim Ret As String
    Select Case VBA.TypeName(Info)
        Case "String"
            If Len(Info) > 0 Then
                ReDim out(0 To Len(Info) - 1) As Byte
                For N = 0 To Len(Info) - 1
                    out(N) = Asc(Mid(Info, N + 1, 1))
                Next
            Else
                ReDim out(-1 To -1) As Byte
            End If
            Convert = out
        Case "Byte()"
            If (ArraySize(Info) > 0) Then
                On Error GoTo dimcheck
                For N = LBound(Info) To UBound(Info)
                    Ret = Ret & Chr(Info(N))
                Next
            End If
            Convert = Ret
    End Select
    Exit Function
dimcheck:
    If Err Then Err.Clear
    For N = LBound(Info, 2) To UBound(Info, 2)
        Ret = Ret & Chr(Info(0, N))
    Next
    Convert = Ret
End Function

Public Function System64Bit() As Boolean
    Dim handle As Long
    Dim is64Bit As Boolean
    is64Bit = False
    handle = GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process")
    If handle <> 0 Then
        IsWow64Process GetCurrentProcess(), is64Bit
    End If
    System64Bit = is64Bit
End Function

Public Function AppPath() As String
    Dim lpTemp As String
#If VBIDE Then
    lpTemp = "C:\Stable-diffusion\"
#Else
    lpTemp = IIf((Right(App.Path, 1) = "\"), App.Path, App.Path & "\")
#End If
    AppPath = lpTemp
End Function

Public Function NextArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            NextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
        Else
            NextArg = Trim(TheParams)
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            NextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
        Else
            NextArg = TheParams
        End If
    End If
End Function

Public Function RemoveArg(ByVal TheParams As String, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveArg = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator)))
        Else
            RemoveArg = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveArg = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator), Len(TheParams) - Len(TheSeperator))
        Else
            RemoveArg = ""
        End If
    End If
End Function

Public Function RemoveNextArg(ByRef TheParams As Variant, ByVal TheSeperator As String, Optional ByVal Compare As VbCompareMethod = vbBinaryCompare, Optional ByVal TrimResult As Boolean = True) As String
    If TrimResult Then
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Trim(Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1))
            TheParams = Trim(Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator)))
        Else
            RemoveNextArg = Trim(TheParams)
            TheParams = ""
        End If
    Else
        If InStr(1, TheParams, TheSeperator, Compare) > 0 Then
            RemoveNextArg = Left(TheParams, InStr(1, TheParams, TheSeperator, Compare) - 1)
            TheParams = Mid(TheParams, InStr(1, TheParams, TheSeperator, Compare) + Len(TheSeperator))
        Else
            RemoveNextArg = TheParams
            TheParams = ""
        End If
    End If
End Function

Public Function PathExists(ByVal URL As String, Optional ByVal IsFile As Variant = Empty) As Boolean
    'checks fot the existance of a path, if it is known checking for a file vs. a folder, set infile = true
    'if folder infile = false, left empty it will attempt to check whichever exists and may false positive.
    'it only checks local system and msn network uri loations, in any formatting, or type falgs possible
    
    Dim Ret As Boolean
    
    If Left(URL, 2) = "\\" Then GoTo altcheck
    If Left(LCase(URL), 7) = "file://" Then
        URL = Replace(Mid(URL, 8), "|/", ":/")
    End If
    If (Len(URL) = 2) And (Mid(URL, 2, 1) = ":") Then
        URL = URL & "\"
    End If
        
    On Error GoTo altcheck

    URL = Replace(URL, "/", "\")
    If InStr(Mid(URL, 3), ":") > 0 Or InStr(Mid(URL, 3), "?") > 0 _
        Or InStr(Mid(URL, 3), """") > 0 Or InStr(Mid(URL, 3), "<") > 0 _
         Or InStr(Mid(URL, 3), ">") > 0 Or InStr(Mid(URL, 3), "|") > 0 Then
        PathExists = False
    ElseIf Len(URL) > 2 Then
        If Len(URL) <= 3 And Mid(URL, 2, 1) = ":" Then
            If VBA.TypeName(IsFile) = "Empty" Then
                PathExists = (Dir(URL, vbVolume) <> "") Or (Dir(URL & "\*") <> "")
            Else
                PathExists = ((Dir(URL, vbVolume) <> "") Or (Dir(URL & "\*") <> "")) And (Not IsFile)
            End If
        Else
            Do While Right(URL, 1) = "\"
                URL = Left(URL, Len(URL) - 1)
            Loop
            Dim attr As Long
            Dim chk1 As String
            Do
                If VBA.TypeName(IsFile) = "Empty" Then
                    chk1 = Dir(URL, attr)
                    If chk1 <> "" And Not Ret Then
                        If InStr(URL, "*") > 0 Then
                            Ret = True
                        Else
                            If Len(URL) > Len(chk1) Then
                                Ret = LCase(Right(URL, Len(chk1))) = LCase(chk1)
                            Else
                                Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                            End If
                        End If
                    End If
                    If Not Ret Then
                        chk1 = Dir(URL, attr + vbDirectory)
                        If chk1 <> "" Then
                            If InStr(URL, "*") > 0 Then
                                Ret = True
                            Else
                                If Len(URL) > Len(chk1) Then
                                    Ret = LCase(Right(URL, Len(chk1))) = LCase(chk1)
                                Else
                                    Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                                End If
                                If Ret Then
                                    If Not (GetAttr(URL) And vbDirectory) = vbDirectory Then
                                        Ret = False
                                    End If
                                End If
                            End If
                        End If
                    End If
                Else
                    If Not IsFile Then
                        chk1 = Dir(URL, attr + vbDirectory)
                        If chk1 <> "" And Not Ret Then
                            If InStr(URL, "*") > 0 Then
                                Ret = True
                            Else
                                If Len(URL) > Len(chk1) Then
                                    Ret = LCase(Right(URL, Len(chk1))) = LCase(chk1)
                                Else
                                    Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                                End If
                                If Ret Then
                                    If Not (GetAttr(URL) And vbDirectory) = vbDirectory Then
                                        Ret = False
                                    End If
                                End If
                            End If
                        End If
                    Else
                        chk1 = Dir(URL, attr)
                        If chk1 <> "" And Not Ret Then
                            If InStr(URL, "*") > 0 Then
                                Ret = True
                            Else
                                If Len(URL) > Len(chk1) Then
                                    Ret = (LCase(Right(URL, Len(chk1))) = LCase(chk1))
                                Else
                                    Ret = LCase(Right(chk1, Len(URL))) = LCase(URL)
                                End If
                            End If
                        End If
                    End If
                End If
                Select Case attr
                    Case vbNormal
                        attr = vbSystem
                    Case vbSystem
                        attr = vbHidden
                    Case vbHidden
                        attr = vbReadOnly
                    Case vbReadOnly
                        attr = vbHidden + vbReadOnly
                    Case vbHidden + vbReadOnly
                        attr = vbHidden + vbSystem
                    Case vbHidden + vbSystem
                        attr = vbHidden + vbSystem + vbReadOnly
                    Case vbHidden + vbSystem + vbReadOnly
                        attr = vbSystem + vbReadOnly
                    Case vbSystem + vbReadOnly
                        attr = vbNormal
                End Select
            Loop Until Ret Or attr = vbNormal
            PathExists = Ret
        End If
    End If

    Exit Function
altcheck:

    Select Case Err.Number
        Case 55, 58, 70
            PathExists = True
        Case Else '53, 52
            Err.Clear
    End Select

'55 File already open
'58 File already exists
'70 Permission denied
'52 Bad file name or number
'53 File not found

    On Error GoTo fixthis:

    If (URL = vbNullString) Then
        PathExists = False
        Exit Function
    ElseIf (Not IsEmpty(IsFile)) Then
        If ((GetFilePath(URL) = vbNullString) And IsFile And (Not (URL = vbNullString))) Or ((GetFileName(URL) = vbNullString) And (Not IsFile) And (Not (URL = vbNullString))) Then
            PathExists = False
            Exit Function
        End If
    End If
    
    On Error GoTo 0
    On Error GoTo -1
    On Error Resume Next
    
    Dim Alt As Integer
    Alt = GetAttr(URL)
    If Err.Number = 0 Then
        If (IsEmpty(IsFile)) Then
            PathExists = True
        Else
            PathExists = IIf(IsFile, Not CBool(((Alt And vbDirectory) = vbDirectory)), CBool(((Alt And vbDirectory) = vbDirectory)))
        End If
        Exit Function
    End If
    
fixthis:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 55, 58, 70
                PathExists = True
            Case Else
                PathExists = False
        End Select
        Err.Clear
    End If
End Function

Public Function ReadFile(ByVal Path As String) As String
    'a comprehensive ReadFile that is share/lock aware with network
    'error proofing and a latent epalse attempt timing and timeout
    
    Dim num As Long
    Dim Text As String
    Dim timeout As Single
    
    num = FreeFile
    On Error Resume Next
    On Local Error Resume Next
    If PathExists(Path, True) Then
        Open Path For Append Shared As #num Len = 1 ' LenB(Chr(CByte(0)))
        Close #num
        Select Case Err.Number
            Case 54, 70, 75
                Err.Clear
                On Error GoTo tryagain
                On Local Error GoTo tryagain
                
                Open Path For Binary Access Read Lock Write As num Len = 1
                If timeout <> 0 Then
                    Open Path For Binary Shared As #num Len = 1
                End If
                Text = String(LOF(num), " ")
                Get #num, 1, Text
                Close #num
            Case Else
                On Error GoTo tryagain
                On Local Error GoTo tryagain
                
                Open Path For Binary Access Read As num Len = 1
                If timeout <> 0 Then
                    Open Path For Binary Shared As num Len = 1
                End If
                Text = String(LOF(num), " ")
                Get #num, 1, Text
                Close #num
        End Select
        If Err Then GoTo failit
        On Error GoTo 0
        On Local Error GoTo 0
    End If
    ReadFile = Text
    Exit Function
tryagain:
    On Error GoTo tryagain
    On Local Error GoTo tryagain
    If timeout = 0 Then
        timeout = Timer
        Resume Next
    ElseIf Timer - timeout > 10 Then
        GoTo failit
    Else
        On Error GoTo failit
        Resume
    End If
failit:
    On Error GoTo 0
    On Local Error GoTo 0
    Err.Raise 75, "ReadFile"
End Function

Public Function WriteFile(ByVal Path As String, ByRef Text As String) As Boolean
    'a comprehensive WriteFile that is share/lock aware with network
    'error proofing and a latent epalse attempt timing and timeout
    
    If PathExists(Path, True) Then
        If (GetAttr(Path) And vbReadOnly) <> 0 Then Exit Function
    End If
    
    Dim timeout As Single
    Dim num As Integer
    
    On Error Resume Next
    On Local Error Resume Next
    
    num = FreeFile
    Open Path For Output Shared As #num Len = 1  'Len = LenB(Chr(CByte(0)))
    Close #num
    
    Select Case Err.Number

        Case 54, 70, 75
            Err.Clear
            On Error GoTo tryagain
            On Local Error GoTo tryagain
            
            Open Path For Binary Access Write Lock Read As #num Len = 1
            If timeout <> 0 Then
                Open Path For Binary Shared As #num Len = 1
            End If
            Put #num, 1, Text
            Close #num
            WriteFile = True
        Case 0
            On Error GoTo tryagain
            On Local Error GoTo tryagain
            
            Open Path For Binary Access Write As #num Len = 1
            If timeout <> 0 Then
                Open Path For Binary Shared As #num Len = 1
            End If
            Put #num, 1, Text
            Close #num
            WriteFile = True
    End Select

    If Err Then GoTo failit
    On Error GoTo 0
    On Local Error GoTo 0
    
    Exit Function
tryagain:
    On Error GoTo tryagain
    On Local Error GoTo tryagain
    
    If timeout = 0 Then
        timeout = Timer
        Resume Next
    ElseIf Timer - timeout > 10 Then
        GoTo failit
    Else
        Resume
    End If
failit:
    On Error GoTo 0
    On Local Error GoTo 0
    Err.Raise 75, "WriteFile"
End Function

Public Function GetFilePath(ByVal URL As String) As String
    Dim nFolder As String
    If InStr(URL, "/") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "/") - 1)
        If nFolder = "" Then nFolder = "/"
    ElseIf InStr(URL, "\") > 0 Then
        nFolder = Left(URL, InStrRev(URL, "\") - 1)
        If nFolder = "" Then nFolder = "\"
    Else
        nFolder = ""
    End If
    GetFilePath = nFolder
End Function

Public Function GetFileTitle(ByVal URL As String) As String
    URL = GetFileName(URL)
    If InStrRev(URL, ".") > 0 Then
        URL = Left(URL, InStrRev(URL, ".") - 1)
    End If
    GetFileTitle = URL
End Function

Public Function GetFileName(ByVal URL As String) As String
    If InStr(URL, "/") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "/") + 1)
    ElseIf InStr(URL, "\") > 0 Then
        GetFileName = Mid(URL, InStrRev(URL, "\") + 1)
    Else
        GetFileName = URL
    End If
End Function







