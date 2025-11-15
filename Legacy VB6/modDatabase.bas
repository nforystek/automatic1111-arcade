Attribute VB_Name = "modDatabase"
#Const modDatabase = -1
Option Explicit

Public dbName As String

Private dbConn As Object
Public Function dbState() As Boolean
    dbState = (Not (dbConn.State = 0))
End Function
Public Function dbFile(Optional ByVal sFileName As String = "") As String
    If sFileName = "" Then sFileName = dbName
    dbFile = AppPath & sFileName
End Function

Public Function dbOpen(Optional ByVal keepopen As Boolean = True) As Boolean
    On Error Resume Next
    Dim tmp As ADODB.Connection
 
    If (dbConn Is Nothing) Then
        Set dbConn = CreateObject("ADODB.Connection")
    ElseIf Not (dbConn.State = 0) Then
        dbConn.Close
    End If
    
    dbConn.ConnectionTimeout = 0
    dbConn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & dbFile("Txt2ImgKiosk.mdb") & ";"
    
    dbOpen = (Not (dbConn.State = 0)) And (Err.Number = 0)
    
    If Not keepopen Then
        dbClose
    Else
        dbOpen = True
    End If
   
End Function

Public Function rsQuery(ByRef rs As Object, ByVal SQLStr) As Boolean
    
    If (rs Is Nothing) Then
        Set rs = CreateObject("ADODB.RecordSet")
    ElseIf Not (rs.State = 0) Then
        rs.Close
    End If
    
    rs.Open SQLStr, dbConn, adOpenKeyset, adLockOptimistic  ', adOpenDynamic, adLockOptimistic ', , 3
    rsQuery = (rs.State = 1)

End Function

Public Function dbQuery(ByVal SQLStr) As Boolean
    Dim rs As Object
    Set rs = CreateObject("ADODB.RecordSet")
    
    rs.Open SQLStr, dbConn, adOpenKeyset, adLockOptimistic ', , 3
    dbQuery = (rs.State = 1)

    dbClose rs
End Function

Public Sub dbClose(Optional ByRef obj As Object = Nothing)

    If (obj Is Nothing) Then Set obj = dbConn

    If Not (obj Is Nothing) Then
        If Not (obj.State = 0) Then obj.Close
        Set obj = Nothing
    End If
End Sub
