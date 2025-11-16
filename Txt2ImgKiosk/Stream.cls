VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stream"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'TOP DOWN
Option Compare Binary

Public Enum MemoryScopes
    heap = 0
    Locking = 1
    [Local] = 2
    [Global] = 4
End Enum

Private pScope As MemoryScopes
Private Pointer As Long

Private Const LMEM_DISCARDABLE = &HF00
Private Const LMEM_FIXED = &H0
Private Const LMEM_ZEROINIT = &H40

Private Const GMEM_DISCARDABLE = &H100
Private Const GMEM_FIXED = &H0

Private Const HEAP_NO_SERIALIZE = &H1

Private Const LMEM_MOVEABLE = &H2
Private Const lPtr = &H40

Private Const GMEM_MOVEABLE = &H2
Private Const GPTR = &H40

Private Declare Function AryPtr Lib "msvbvm60" Alias "VarPtr" (ary() As Any) As Long
 
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

'&H11000000
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function FlushInstructionCache Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, ByVal dwSize As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long

Private Declare Function LocalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal wBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal wBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function LocalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function LocalSize Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function HeapUnlock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapLock Lib "kernel32" (ByVal hHeap As Long) As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long, ByVal dwBytes As Long) As Long
Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Const READ_CONTROL = &H20000

Private Const SYNCHRONIZE = &H100000
Private Const FILE_READ_ATTRIBUTES = (&H80)              '  all
Private Const FILE_READ_DATA = (&H1)                     '  file pipe
Private Const FILE_READ_EA = (&H8)                       '  file directory
Private Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Private Const FILE_APPEND_DATA = (&H4)                   '  file
Private Const FILE_WRITE_ATTRIBUTES = (&H100)            '  all
Private Const FILE_WRITE_DATA = (&H2)                    '  file pipe
Private Const FILE_WRITE_EA = (&H10)                     '  file directory
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)
Private Const FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)

Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000

Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000

Private Const INVALID_SET_FILE_POINTER = -1

Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const CREATE_NEW = 1

Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        Offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

Private Declare Function VarPtrArray Lib "MSVBVM60.DLL" Alias "VarPtr" (Var() As Any) As Long
Private Declare Function StrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As Long) As Long
Private Declare Function StrCpyReverse Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Long, ByVal lpString2 As String) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

Private Declare Function hWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function hReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Private Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

Private Const MAX_PATH = 260

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long


Public Property Get Address(Optional ByVal Offset As Long = 0) As Long ' _
Retrieves the address numberical value where the memory this object is allocated at.
    RtlMoveMemory ByVal VarPtr(Address), (Pointer + Offset), 4
End Property

Public Property Get Scope() As MemoryScopes ' _
Retrieves the type of memory functionally this stream is utilizing, Heap, Locking and/or Local or Global defined by the Reset method.
    Scope = pScope
End Property

Private Sub Class_Initialize()
    pScope = MemoryScopes.heap
    'pScope = MemoryScopes.[Local] Or MemoryScopes.Locking
End Sub

Private Sub Class_Terminate()
    Reset pScope
End Sub

Private Sub MemoryModify(Optional ByVal Size As Long = 0)

    If (Pointer <> 0) Then
        If ((Scope And MemoryScopes.Locking) = MemoryScopes.Locking) And (Not (Length = 0)) Then
            If ((Scope And MemoryScopes.Local) = MemoryScopes.Local) Then
                If LocalUnlock(Pointer) <> 0 Then Err.Raise 8, App.Title, "Local memory unlock failure."
            ElseIf ((Scope And MemoryScopes.Global) = MemoryScopes.Global) Then
                If GlobalUnlock(Pointer) <> 1 Then Err.Raise 8, App.Title, "Global memory unlock failure."
            ElseIf ((Scope And MemoryScopes.heap) = MemoryScopes.heap) Then
                If HeapUnlock(GetProcessHeap) = 0 Then Err.Raise 8, App.Title, "Heap memory unlock failure."
            End If
        End If
    End If
    
    If (Size = 0) Then
        If (Pointer <> 0) And (Not (Length = 0)) Then
            If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
                LocalFree Pointer
            ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
                GlobalFree Pointer
            ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
                HeapFree GetProcessHeap, 0, ByVal Pointer
            End If
            Pointer = 0
        End If
    Else
        If (Pointer = 0) And (Length = 0) Then
            If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalAlloc(LMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalAlloc(GMEM_MOVEABLE And VarPtr(Pointer), Size)
            ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
                Pointer = HeapAlloc(GetProcessHeap, 0, Size)
            End If
        Else
            If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
                Pointer = LocalReAlloc(Pointer, Size, LMEM_MOVEABLE Or lPtr)
            ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
                Pointer = GlobalReAlloc(Pointer, Size, GMEM_MOVEABLE Or GPTR)
            ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
                Pointer = HeapReAlloc(GetProcessHeap, 0, ByVal Pointer, Size)
            End If
        End If
    End If

    If (Pointer <> 0) Then
        If ((Scope And MemoryScopes.Locking) = MemoryScopes.Locking) And (Not (Length = 0)) Then
            If ((Scope And MemoryScopes.Local) = MemoryScopes.Local) Then
                If Pointer <> LocalLock(Pointer) Then Err.Raise 8, App.Title, "Local memory lock mismatch."
            ElseIf ((Scope And MemoryScopes.Global) = MemoryScopes.Global) Then
                If Pointer <> GlobalLock(Pointer) Then Err.Raise 8, App.Title, "Global memory lock mismatch."
            ElseIf ((Scope And MemoryScopes.heap) = MemoryScopes.heap) Then
                If HeapLock(GetProcessHeap) = 0 Then Err.Raise 8, App.Title, "Heap memory lock mismatch."
            End If
        End If
    End If

End Sub

Public Sub Clone(ByRef Origin As Stream, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) ' _
Duplicates the Origin Stream object truning the calling instances stream data to be exactly the same, or same as a portion of the Origin strean, optionally defined by the Offset and Width.
    Dim lSize As Long
    Reset Origin.Scope
    lSize = Origin.Length
    If Width = -1 Or lSize - Offset < Width Then Width = lSize - Offset
    If Offset < lSize And Offset >= 0 And Offset + Width <= lSize And Width > 0 Then
        Length = Width
        RtlMoveMemory ByVal Pointer, ByVal Origin.Address(Offset), Width
    End If
End Sub

Public Sub Post(ByRef Data As Byte) ' _
Appends Data bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data. Psudonym to Concat()
    Length = Length + 1
    Plot Length - 1, Data
End Sub

Public Sub Concat(ByRef Data() As Byte) ' _
Appends Data bytes onto the stream at the end of the stream increasing it's size by the amount of bytes in supplied by Data. Psudonym to Post()
    Dim nSize As Long
    nSize = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
    If (nSize > 0) Then
        Dim lSize As Long
        lSize = Length
        MemoryModify (nSize + lSize)
        RtlMoveMemory ByVal Pointer + lSize, ByVal VarPtr(Data(LBound(Data))), nSize
'    Else
'        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Property Get Length() As Long ' _
Gets the amount of bytes with in this instance of the stream object.
    If Pointer <> 0 Then
        If ((pScope And MemoryScopes.Local) = MemoryScopes.Local) Then
            Length = LocalSize(Pointer)
        ElseIf ((pScope And MemoryScopes.Global) = MemoryScopes.Global) Then
            Length = GlobalSize(Pointer)
        ElseIf ((pScope And MemoryScopes.heap) = MemoryScopes.heap) Then
            Length = HeapSize(GetProcessHeap, 0, ByVal Pointer)
        End If
    Else
        Length = 0
    End If
End Property

Public Property Let Length(ByVal RHS As Long) ' _
Sets the amount of bytes with in this instance of the stream object, least the current value set truncated on the right, or expands on the right with garbage bytes.
    If (RHS > 0) Then
        MemoryModify RHS
    ElseIf (RHS = 0) Then
        Reset pScope
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Property

Public Sub PlaceAt(ByRef Data() As Byte, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) ' _
Places contents of memory byte array Data into the stream excluding Offset number of bytes, and only up to Width but not over the total allotted.
    Dim lSize As Long
    lSize = Length
    If Width = -1 Then Width = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
    If ((lSize - Offset) <= Width) Then Width = lSize - Offset
    If (Width > 0) And (Pointer <> 0) And lSize >= (Offset + Width) Then
        RtlMoveMemory ByVal Pointer + Offset, ByVal VarPtr(Data(LBound(Data))), Width
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Function Partial(Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Byte() ' _
Returns a sub set array of bytes from the stream excluding Offset from the beginning, and only of optional Width amount of bytes, or to the end of the stream when not supplied Width, to retrieve.

    Dim lSize As Long
    lSize = Length
    If ((Width - Offset) > lSize) Or (Width = -1) Then Width = lSize - Offset
    Dim Ret() As Byte
    If (Width > 0) And (Pointer <> 0) Then
        ReDim Ret(0 To Width - 1) As Byte
        RtlMoveMemory ByVal VarPtr(Ret(LBound(Ret))), ByVal Pointer + Offset, Width
    Else
        ReDim Ret(-1 To -1) As Byte
    End If
    Partial = Ret
    Erase Ret
End Function

Public Sub Prepend(ByRef Data() As Byte) ' _
Places Data bytes at the beginning of the stream adding on to the streams total size by the amount of bytes in Data.

    If (Length = 0) Or (Pointer = 0) Then
        Concat Data
    Else
        Dim nSize As Long
        nSize = UBound(Data) + -CInt(Not CBool(-LBound(Data)))
        Length = Length + (nSize * 2)
        Pull nSize
        PlaceAt Data, 0, nSize
    End If
End Sub

Public Sub Pull(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the left, truncating the contents on the right, and leaving Offset number of garbage bytes on the left.

    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        If Pointer <> 0 Then
            RtlMoveMemory ByVal Pointer + Offset, ByVal Pointer, (-lSize + -Offset + (lSize * 2))
            MemoryModify (lSize - Offset)
        End If
    ElseIf (Offset = lSize) Then
        Reset pScope
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub


Public Sub Push(ByVal Offset As Long) ' _
Shifts the entire container of the stream by Offset to the right, truncating the contents on the left, and leaving Offset number of garbage bytes on the right.
    Dim lSize As Long
    lSize = Length
    If (Offset < lSize) Then
        If Pointer <> 0 Then
            RtlMoveMemory ByVal Pointer, ByVal Pointer + Offset, (-lSize + -Offset + (lSize * 2))
            MemoryModify (lSize - Offset)
        End If
    ElseIf (Offset = lSize) Then
        Reset pScope
'    Else
'        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub


Public Sub Reset(Optional ByRef ScopeAs As MemoryScopes) ' _
Resets the stream object discarding the stored data.  Optionally setting Scopes, otherwise it keeps the current scope, for any further use.
    MemoryModify
End Sub

Public Function Pass(ByVal Data As Byte, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Long ' _
Preforms a pass for Data returning the count of it in the stream, optionally confined to Offset and Width.
'i.e. a stream with "hello" as the only thing in it's buffer, then calling Pass(asc("l")) returns a 2.
    Dim lPos As Long
    Dim bPart() As Byte
    bPart = Partial(Offset, Width)
    For lPos = LBound(bPart) To UBound(bPart)
        If bPart(lPos) = Data Then Pass = Pass + 1
    Next
End Function

Public Function Poll(ByVal Data As Byte, ByVal Ballot As Long, Optional ByVal Offset As Long = 0, Optional ByRef Width As Long = -1) As Long ' _
Preforms a poll for Data returning the offset up to the numerical Ballot's occurance of it, optionally confined to Offset and Width.
'i.e. Poll(vbCr, 3) returns the offset of the third occurance of vbCr with in the stream, if non exist, the length of the stream is returned.
    If Ballot > 0 And Width <> 0 And Length > 0 Then
        Dim lPos As Long
        Dim bPart() As Byte
        If Ballot = 0 Then Exit Function
        bPart = Partial(Offset, Width)
        For lPos = LBound(bPart) To UBound(bPart)
            If bPart(lPos) = Data Then
                Poll = Poll + 1
                If Poll = Ballot Then
                    Poll = lPos + LBound(bPart)
                    Exit Function
                End If
                
            End If
        Next
        If Width > 0 Then
            Poll = Width
        Else
            Poll = Length
        End If
 '   Else
'        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Function

Public Sub Pinch(ByVal Offset As Long, ByVal Width As Long) ' _
Removes a portion of the stream defined by Offset and Width.
    Dim lSize As Long
    lSize = Length
    If Offset < lSize And Offset > 0 And Offset + Width < lSize And Width > 0 Then
        Dim N As New Stream
        
        N.Concat Partial(0, Offset)
        N.Concat Partial(Offset + Width, lSize - (Offset + Width))
        Clone N
        
        Set N = Nothing
        
    ElseIf Offset = 0 And Width = lSize Then
        Reset
    ElseIf Offset = 0 Then
        Push Width
        'Length = Length - Width
    ElseIf Offset + Width = lSize Then
        Length = Length - Width
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Sub Pyramid(ByRef SubStream As Stream, ByVal Offset As Long, ByVal Width As Long) ' _
Places the full SubStream data into the calling stream overwriting the segment Offset and Wdith define.
'the size of Offset to Width and SubStream make no difference, only and all from OFffset to Width get
'overwritten by the size of SubStream put into the Calling stream, potentially changing its size.
    Dim lSize As Long
    lSize = Length
    If Offset >= 0 And Offset + Width <= lSize Then
        Dim lSize2 As Long
        lSize2 = SubStream.Length
        If lSize2 > 0 Then
            Dim N As New Stream
            
            If Offset > 0 Then N.Concat Partial(0, Offset)
            N.Concat SubStream.Partial
            If lSize - (Offset + Width) > 0 Then N.Concat Partial(Offset + Width, lSize - (Offset + Width))
            Clone N
            
            Set N = Nothing
        Else
            Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
        End If
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Sub

Public Function Peek(Optional ByVal Offset As Long = 0) As Byte ' _
Gets the first byte in the data stream, optionally excluding Offset number of bytes retruning the immediate byte after Offset.
    If Offset >= 0 And Offset < Length Then
        RtlMoveMemory ByVal VarPtr(Peek), ByVal Pointer + Offset, 1
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If
End Function

Public Sub Plot(ByVal Offset As Long, ByVal Data As Byte)   ' _
Sets a byte in the data stream immediatly after Offset and does not alter the size of the stream.

    If Offset >= 0 And Offset < Length Then
        RtlMoveMemory ByVal Pointer + Offset, ByVal VarPtr(Data), 1
    Else
        Err.Raise 8, App.Title, "Invalid parameter, array, or object length requirements."
    End If

End Sub

