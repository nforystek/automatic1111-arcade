Attribute VB_Name = "modReg64Bit"
Option Explicit

' === Constants ===
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005

' Registry value types
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

' Registry options
Public Const REG_OPTION_NON_VOLATILE = 0

' Access rights
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_WOW64_64KEY = &H100
Public Const KEY_WOW64_32KEY = &H200
Public Const KEY_READ = &H20019
Public Const KEY_WRITE = &H20006
Public Const KEY_ALL_ACCESS = &HF003F

' Disposition values
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

' Error codes
Public Const ERROR_SUCCESS = 0&

' === Registry API Declarations ===
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal ulOptions As Long, _
    ByVal samDesired As Long, _
    phkResult As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal Reserved As Long, _
    ByVal lpClass As String, _
    ByVal dwOptions As Long, _
    ByVal samDesired As Long, _
    ByVal lpSecurityAttributes As Long, _
    phkResult As Long, _
    lpdwDisposition As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long
    
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteKeyEx Lib "advapi32.dll" Alias "RegDeleteKeyExA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    ByVal samDesired As Long, _
    ByVal Reserved As Long) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" ( _
    ByVal hKey As Long) As Long

Public Function RegWriteString(hRoot As Long, sSubKey As String, sValueName As String, sValue As String, Optional b64Bit As Boolean = False) As Boolean
    If b64Bit And Not System64Bit Then Exit Function
    Dim hKey As Long, lResult As Long, lDisp As Long
    Dim flags As Long
    flags = KEY_ALL_ACCESS
    If b64Bit Then flags = flags Or KEY_WOW64_64KEY Else flags = flags Or KEY_WOW64_32KEY
    
    lResult = RegCreateKeyEx(hRoot, sSubKey, 0, vbNullString, 0, flags, 0, hKey, lDisp)
    If lResult = ERROR_SUCCESS Then
        sValue = sValue & vbNullChar
        lResult = RegSetValueEx(hKey, sValueName, 0, REG_SZ, ByVal sValue, Len(sValue))
        RegCloseKey hKey
        RegWriteString = (lResult = ERROR_SUCCESS)
    End If
End Function

Public Function RegReadString(hRoot As Long, sSubKey As String, sValueName As String, Optional b64Bit As Boolean = False) As String
    If b64Bit And Not System64Bit Then Exit Function
    Dim hKey As Long, lResult As Long, lType As Long, lSize As Long
    Dim sBuffer As String
    Dim flags As Long
    flags = KEY_QUERY_VALUE
    If b64Bit Then flags = flags Or KEY_WOW64_64KEY Else flags = flags Or KEY_WOW64_32KEY
    
    lResult = RegOpenKeyEx(hRoot, sSubKey, 0, flags, hKey)
    If lResult = ERROR_SUCCESS Then
        lResult = RegQueryValueEx(hKey, sValueName, 0, lType, ByVal 0&, lSize)
        If lResult = ERROR_SUCCESS And lSize > 0 Then
            sBuffer = String$(lSize, vbNullChar)
            lResult = RegQueryValueEx(hKey, sValueName, 0, lType, ByVal sBuffer, lSize)
            If lResult = ERROR_SUCCESS Then
                RegReadString = Left$(sBuffer, InStr(sBuffer, vbNullChar) - 1)
            End If
        End If
        RegCloseKey hKey
    End If
End Function

' Write a REG_DWORD value
Public Function RegWriteDWORD(hRoot As Long, sSubKey As String, sValueName As String, lValue As Long, Optional b64Bit As Boolean = False) As Boolean
    If b64Bit And Not System64Bit Then Exit Function
    Dim hKey As Long
    Dim lResult As Long
    Dim lDisposition As Long
    Dim samDesired As Long
    
    ' Handle 64-bit registry view if requested
    samDesired = KEY_WRITE
    If b64Bit Then samDesired = samDesired Or KEY_WOW64_64KEY
    
    ' Open or create the key
    lResult = RegCreateKeyEx(hRoot, sSubKey, 0&, vbNullString, REG_OPTION_NON_VOLATILE, samDesired, 0&, hKey, lDisposition)
    If lResult = ERROR_SUCCESS Then
        ' Write the DWORD value
        lResult = RegSetValueEx(hKey, sValueName, 0&, REG_DWORD, lValue, 4)
        RegCloseKey hKey
        RegWriteDWORD = (lResult = ERROR_SUCCESS)
    Else
        RegWriteDWORD = False
    End If
End Function

' Read a REG_DWORD value
Public Function RegReadDWORD(hRoot As Long, sSubKey As String, sValueName As String, Optional b64Bit As Boolean = False) As Long
    If b64Bit And Not System64Bit Then Exit Function
    Dim hKey As Long
    Dim lResult As Long
    Dim lType As Long
    Dim lData As Long
    Dim lDataSize As Long
    Dim samDesired As Long
    
    ' Handle 64-bit registry view if requested
    samDesired = KEY_READ
    If b64Bit Then samDesired = samDesired Or KEY_WOW64_64KEY
    
    ' Open the key
    lResult = RegOpenKeyEx(hRoot, sSubKey, 0&, samDesired, hKey)
    If lResult = ERROR_SUCCESS Then
        lDataSize = 4
        lResult = RegQueryValueEx(hKey, sValueName, 0&, lType, lData, lDataSize)
        If lResult = ERROR_SUCCESS And lType = REG_DWORD Then
            RegReadDWORD = lData
        Else
            RegReadDWORD = 0 ' Default if not found or wrong type
        End If
        RegCloseKey hKey
    Else
        RegReadDWORD = 0
    End If
End Function

' Delete a registry value
Public Function RegDeleteValueName(hRoot As Long, sSubKey As String, sValueName As String, Optional b64Bit As Boolean = False) As Boolean
    If b64Bit And Not System64Bit Then Exit Function
    Dim hKey As Long
    Dim lResult As Long
    Dim samDesired As Long
    
    ' Handle 64-bit registry view if requested
    samDesired = KEY_SET_VALUE
    If b64Bit Then samDesired = samDesired Or KEY_WOW64_64KEY
    
    ' Open the key
    lResult = RegOpenKeyEx(hRoot, sSubKey, 0&, samDesired, hKey)
    If lResult = ERROR_SUCCESS Then
        ' Delete the value
        lResult = RegDeleteValue(hKey, sValueName)
        RegCloseKey hKey
        RegDeleteValueName = (lResult = ERROR_SUCCESS)
    Else
        RegDeleteValueName = False
    End If
End Function

' Delete an entire registry key (and all its values, but not subkeys)
Public Function RegDeleteKeyPath(hRoot As Long, sSubKey As String, Optional b64Bit As Boolean = False) As Boolean
    If b64Bit And Not System64Bit Then Exit Function
    Dim lResult As Long
    
    ' Basic delete (no WOW64 handling)
    If Not b64Bit Then
        lResult = RegDeleteKey(hRoot, sSubKey)
        RegDeleteKeyPath = (lResult = ERROR_SUCCESS)
    Else
        ' Use RegDeleteKeyEx if you want 64-bit view support
        lResult = RegDeleteKeyEx(hRoot, sSubKey, KEY_WOW64_64KEY, 0&)
        RegDeleteKeyPath = (lResult = ERROR_SUCCESS)
    End If
End Function


