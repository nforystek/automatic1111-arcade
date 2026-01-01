Attribute VB_Name = "modDBFiles"
#Const modFiles = -1
Option Explicit

Public rsFile As New ADODB.Recordset

Public Sub FileGetTopVotes()
    On Error Resume Next
    
    rsQuery rsFile, "SELECT * FROM Files ORDER BY FileVote DESC;"
    
    If Not Err.Number = 0 Then
        MyDebugPrint "FileGetTopVotes() Error: " & Err.Description

        Err.Clear
        
    End If
End Sub

Public Function FilePutArray(ByVal Filename As String, ByRef b() As Byte, ByVal Prompt As String, ByVal Negate As String) As Long
    On Error Resume Next
    
    
    
    Dim TempID As String
    Dim FileID As Long
    Dim filePos As Long
    Dim strdata As String
    Dim recCount As Long
    
    rsQuery rsFile, "SELECT * FROM Files ORDER BY FileVote DESC, FileName DESC;"
    If Not rsFile.EOF Then
        rsFile.MoveFirst
        Do Until rsFile.EOF
            recCount = recCount + 1
            rsFile.MoveNext
        Loop
    End If

    Do While recCount > (TotalImages - 1)
        rsFile.MoveLast
        TempID = TempID & rsFile("ID") & ","
        recCount = recCount - 1
    Loop
    
    Do Until TempID = ""
        rsQuery rsFile, "DELETE FROM Files WHERE ID=" & RemoveNextArg(TempID, ",") & ";"
    Loop
     
    TempID = GUID

    rsQuery rsFile, "INSERT INTO Files (FileName) " & "VALUES ('" & TempID & "');"
    rsQuery rsFile, "SELECT * FROM Files WHERE FileName='" & TempID & "';"
    
    FileID = rsFile("ID")
    FilePutArray = rsFile("ID")
    
    rsQuery rsFile, "UPDATE Files SET FileName='" & Filename & "', Prompt='" & Replace(Prompt, "'", "''") & "', Negate='" & Replace(Negate, "'", "''") & "' WHERE ID=" & FileID & ";"
    rsQuery rsFile, "SELECT * FROM Files WHERE ID=" & FileID & ";"

    strdata = StrConv(b(), vbUnicode)

    If Len(strdata) > 0 Then

        rsFile("FileData").AppendChunk strdata
        rsFile.Update

    End If
    
    dbClose rsFile
    
    If Not Err.Number = 0 Then
        MyDebugPrint "FilePutArray() Error: " & Err.Description
        Err.Clear
        
    End If

End Function

Public Sub FileGetArray(ByVal FileID As String, ByRef b() As Byte, ByRef votes As Long) 'Returns the name of the local file created
    On Error Resume Next
    
    Dim lngLen As Long
    Dim strdata As String
    
    rsQuery rsFile, "SELECT * FROM Files WHERE ID=" & FileID & ";"
    
    lngLen = rsFile.Fields("FileData").ActualSize

    If lngLen > 0 Then

        ReDim b(0 To lngLen - 1) As Byte

         b = StrConv(rsFile("FileData").GetChunk(lngLen), vbFromUnicode)

    End If
    votes = rsFile("FileVote")

    dbClose rsFile
    
    If Not Err.Number = 0 Then
        MyDebugPrint "FileGetArray() Error: " & Err.Description
        Err.Clear
        
    End If

End Sub

Public Function FileVoteFor(ByVal FileID As String) As Long 'Returns the name of the local file created
    On Error Resume Next
    
    rsQuery rsFile, "SELECT * FROM Files WHERE ID=" & FileID & ";"
    
    Dim votes As Long
    votes = rsFile("FileVote") + 1
    
    rsQuery rsFile, "UPDATE Files SET FileVote=" & votes & " WHERE ID=" & FileID & ";"

    FileVoteFor = votes

    dbClose rsFile
    
    If Not Err.Number = 0 Then
        MyDebugPrint "FileVoteFor() Error: " & Err.Description
        Err.Clear
        
    End If

End Function


Public Function FileRemove(ByVal FileID As String)
    On Error Resume Next
    
    dbQuery "DELETE FROM Files WHERE ID=" & FileID & ";"

    If Not Err.Number = 0 Then
        MyDebugPrint "FileRemove() Error: " & Err.Description
        Err.Clear
        
    End If
End Function
