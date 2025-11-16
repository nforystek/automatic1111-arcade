VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Talk to DOS"
   ClientHeight    =   4530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   0
      Width           =   1695
   End
   Begin VB.TextBox txtSend 
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Text            =   "PING localhost"
      Top             =   0
      Width           =   4815
   End
   Begin VB.TextBox txtDos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   4095
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   480
      Width           =   6615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 'Lets specify the use of the dll right here.
 Dim WithEvents oLaunch As RedirectLib.Application
Attribute oLaunch.VB_VarHelpID = -1
 Dim bExecute As Boolean

Private Sub btnSend_Click()
#If UNICODE = -1 Then
  oLaunch.Write StrConv(txtSend.Text + vbCrLf, vbFromUnicode)   'this sends the writing to the cmd.exe then a return
#Else
 oLaunch.Write txtSend.Text + vbCrLf 'this sends the writing to the cmd.exe then a return
#End If
End Sub

Private Sub Form_Load()
 Set oLaunch = New RedirectLib.Application
  oLaunch.BufferSize = 8192
  oLaunch.Wait = 1000
  bExecute = False 'This means it is not running yet
  'Lets now start the connection to the cmd.exe
  If bExecute Then 'Checks to see if already connected
   oLaunch.Stop 'if it is, then it stops the connection.
  End If
  oLaunch.Name = "c:\windows\system32\cmd.exe" 'say where the cmd.exe is or the command.exe
   Select Case oLaunch.Start 'Starts the connection to the command prompt
    Case laAlreadyRunning
     'Already going
    Case laWindowsError
     MsgBox "Error: " & CStr(oLaunch.LastErrorNumber) & "!" 'if there was a problem.
    Case laOk
     bExecute = True 'Everything went smooth, we are now connected to cmd.exe
   End Select
  Me.Show 'Show the form
  txtSend.SetFocus 'set focus to the typing box for immediate availablitity.
End Sub

Private Sub Form_Terminate()
 Set oLaunch = Nothing 'Sets everything to nothing for closing the app.
End Sub

Private Sub Form_Unload(Cancel As Integer)
  oLaunch.Stop 'this closes the onnection to the cmd.exe
End Sub

Private Sub oLaunch_DataReceived(ByVal sData As String)
 'This is when the dll gets information back from the cmd.exe
 Debug.Print "oLaunch_DataReceived"
#If UNICODE = -1 Then
    sData = StrConv(sData, vbUnicode)
#End If
 txtDos.Text = txtDos.Text + sData 'Write the info to the text box.
 Debug.Print "[" & sData & "]"
 Debug.Print InStr(sData, vbCrLf) & " " & (Right(sData, 1) = vbLf)
 txtDos.SelStart = Len(txtDos.Text) 'Just places the cursor at the bottom of the text _
 box just like a real dos prompt
End Sub

Private Sub oLaunch_ProcessEnded()
 bExecute = False 'if the connection is closed, it just resets the variable.
End Sub

Private Sub txtSend_KeyUp(KeyCode As Integer, Shift As Integer)
'This is just to watch for the Enter key to be pressed.
'If it is pressed then it calls for the send button to be pressed.
'Which inturn sends the data.
 If KeyCode = 13 Then
  btnSend_Click
  txtSend.Text = ""
 End If
End Sub
