Attribute VB_Name = "modSettings"
#Const modSettings = -1
Option Explicit

Private Settings As New Collection

Public Sub SetSetting(ByVal sName As String, ByVal sValue As Variant)

    Dim s As New Setting
    With s
        .Name = sName
        .Value = sValue
    End With
    On Error Resume Next
    Settings.Add s, sName
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Settings.Remove sName
        Settings.Add s, sName
    Else
        On Error GoTo 0
    End If
End Sub

Public Function LoadSettings() As Boolean
    Dim rs As New ADODB.Recordset
    
    rsQuery rs, "SELECT * FROM Settings;"
    
    If rs.EOF Then
            
'        SetSetting "WinTop", CStr((Screen.Height / 2) - (5940 / 2))
'        SetSetting "WinLeft", CStr((Screen.Width / 2) - (8265 / 2))
'        SetSetting "WinWidth", "9270"
'        SetSetting "WinHeight", "7275"
'        SetSetting "WinState", "0"

        SetSetting "Version", App.Major & "." & App.Minor & "." & App.Revision

    
        LoadSettings = True
    Else
        
        Dim F As ADODB.Field
        
        For Each F In rs.Fields
            SetSetting F.Name, IIf(IsNull(rs(F.Name)), "", rs(F.Name))
        Next
        
        
        If ExistsSetting("Version") Then
            If GetSetting("Version") <> App.Major & "." & App.Minor & "." & App.Revision Then
                MyDebugPrint "LoadSettings() Error1: The database selected is for Locket " & GetSetting("Version") & ", the running version is" & vbCrLf & _
                       "Locket is " & App.Major & "." & App.Minor & "." & App.Revision & " and can't open databases from other Locket versions."
            Else
                LoadSettings = True
            End If
        Else
            MyDebugPrint "LoadSettings() Error2: The database selected is for 2.0.0, the running version" & vbCrLf & _
                   "is " & App.Major & "." & App.Minor & "." & App.Revision & " and is not compatible"
        End If
    End If
    
    dbClose rs
End Function

Public Sub SaveSettings()
    
    Dim s As Setting
    Dim sNames As String
    Dim sValues As String
    
    For Each s In Settings
        sNames = sNames & s.Name & ","
        If IsNumeric(s.Value) Then
            sValues = sValues & s.Value & ","
        Else
            sValues = sValues & "'" & Replace(s.Value, "'", "''") & "',"
        End If
    Next
    sNames = Left(sNames, Len(sNames) - 1)
    sValues = Left(sValues, Len(sValues) - 1)
    dbQuery "DELETE * FROM Settings;"
    On Error GoTo 0

    dbQuery "INSERT INTO Settings (" & sNames & ") VALUES (" & sValues & ");"


End Sub
Public Function ExistsSetting(ByVal sName As String) As Boolean
    On Error Resume Next
    Dim test As String
    test = Settings(sName).Name
    ExistsSetting = (Err.Number = 0)
    If Err Then Err.Clear
    On Error GoTo 0
End Function

Public Function GetSetting(ByVal sName As String) As Variant
    
    GetSetting = Settings.Item(sName).Value
End Function

