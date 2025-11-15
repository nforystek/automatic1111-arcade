Attribute VB_Name = "modFolders"
#Const modFolders = -1
Option Explicit

Private Type SHITEMID
    cb As Long
    abID() As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type


Private Const NOERROR = 0
Private Const CSIDL_MYMUSIC = &HD  ' My Music folder
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const MAX_PATH = 260
Private Const TOKEN_QUERY = (&H8)

Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pIdl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hwndOwner As Long, ByVal nFolder As Long, pIdl As ITEMIDLIST) As Long

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal Path As String, ByVal cbBytes As Long) As Long

Private Declare Function GetUserProfileDirectory Lib "userenv" Alias "GetUserProfileDirectoryA" (ByVal hToken As Long, ByVal lpProfileDir As String, lpcchSize As Long) As Boolean

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, ByRef TokenHandle As Long) As Long

Public Function GetMyMusicFolder() As String
    Dim BI As BROWSEINFO
    Dim nFolder As Long
    Dim IDL As ITEMIDLIST
    Dim sPath As String
    With BI
        nFolder = CSIDL_MYMUSIC
        If SHGetSpecialFolderLocation(ByVal 0&, ByVal nFolder, IDL) = NOERROR Then
            .pidlRoot = IDL.mkid.cb
        End If
        .pszDisplayName = String$(MAX_PATH, 0)

        .ulFlags = BIF_RETURNONLYFSDIRS

    End With

    sPath = String$(MAX_PATH, 0)
    SHGetPathFromIDList ByVal IDL.mkid.cb, ByVal sPath

    sPath = Left(sPath, InStr(sPath, vbNullChar) - 1)

    GetMyMusicFolder = sPath
End Function

Public Function GetSystem32Folder() As String
    Static winDir As String
    If winDir = "" Then
        Dim Ret As Long
        winDir = String(45, Chr(0))
        Ret = GetSystemDirectory(winDir, 45)
        winDir = Trim(Replace(winDir, Chr(0), ""))
        If Right(winDir, 1) <> "\" Then winDir = winDir + "\"
    End If
    GetSystem32Folder = winDir
End Function

Public Function GetCurrentUserProfileFolder() As String

    Dim hToken As Long
    Dim sLibrary As String
    sLibrary = String(255, Chr(0))
    OpenProcessToken GetCurrentProcess, TOKEN_QUERY, hToken
    GetUserProfileDirectory hToken, sLibrary, 255
    GetCurrentUserProfileFolder = Replace(sLibrary, Chr(0), "")
    
End Function

Public Function MakeFolder(ByRef Path As String)
    On Error Resume Next
    If InStr(Path, "\") > 0 Then
        GetAttr Left(Path, InStrRev(Path, "\") - 1)
        If Err.Number = 76 Or Err.Number = 53 Then
            Err.Clear
            MakeFolder = Path
            Path = MakeFolder(Left(Path, InStrRev(Path, "\") - 1))
        Else
            MakeFolder = Path
        End If
    End If
    If Err.Number = 0 Then
        GetAttr MakeFolder
        If Err.Number = 76 Or Err.Number = 53 Then
            Err.Clear
            On Error GoTo -1
            MkDir MakeFolder
        End If
    End If
End Function








