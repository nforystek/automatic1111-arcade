VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10095
   ClientLeft      =   -45
   ClientTop       =   -105
   ClientWidth     =   16155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Transformers Movie"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   Moveable        =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   16155
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3645
      Top             =   2205
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1920
      Left            =   360
      ScaleHeight     =   1920
      ScaleWidth      =   2865
      TabIndex        =   14
      Top             =   5880
      Visible         =   0   'False
      Width           =   2865
      Begin VB.Image Image5 
         Height          =   1185
         Index           =   0
         Left            =   255
         Stretch         =   -1  'True
         Top             =   195
         Width           =   1440
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4875
      Left            =   5040
      ScaleHeight     =   4875
      ScaleWidth      =   4725
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   4725
      Begin VB.Image Image3 
         Height          =   6615
         Index           =   1
         Left            =   2280
         Picture         =   "frmMain.frx":0442
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   27375
      End
      Begin VB.Image Image3 
         Height          =   6615
         Index           =   0
         Left            =   480
         Picture         =   "frmMain.frx":1207C
         Stretch         =   -1  'True
         Top             =   1920
         Width           =   27375
      End
      Begin VB.Image Image3 
         Height          =   6615
         Index           =   2
         Left            =   1200
         Picture         =   "frmMain.frx":282EA
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   27375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   16000
      Left            =   4230
      Top             =   2460
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Vote (ALT)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5625
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   8805
      Visible         =   0   'False
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate (ENTER)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7905
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Next (PAGE DOWN)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   285
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8805
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2925
      Left            =   360
      ScaleHeight     =   2925
      ScaleWidth      =   4680
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   4680
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Content not suitable for all viewers was detected and it has been censored. Your credits have been refunded."
         ForeColor       =   &H00FFFFFF&
         Height          =   1515
         Left            =   240
         TabIndex        =   15
         Top             =   180
         Width           =   4185
      End
      Begin VB.Image Image1 
         Height          =   1365
         Left            =   0
         Top             =   0
         Width           =   1965
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Prompt"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   540
      TabIndex        =   4
      Top             =   150
      Visible         =   0   'False
      Width           =   4365
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   465
         Width           =   3780
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Negate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   60
      TabIndex        =   3
      Top             =   1245
      Visible         =   0   'False
      Width           =   4095
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   480
         Width           =   3510
      End
   End
   Begin VB.CommandButton Command0 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Back (PAGE UP)"
      BeginProperty Font 
         Name            =   "Transformers Movie"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7920
      Width           =   4815
   End
   Begin VB.CommandButton Command3 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Elections (CTRL)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   7800
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Leader Board"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   10560
      TabIndex        =   7
      Top             =   1440
      Visible         =   0   'False
      Width           =   4980
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   720
         ScaleHeight     =   945
         ScaleWidth      =   10545
         TabIndex        =   16
         Top             =   2640
         Width           =   10545
         Begin VB.CommandButton Command5 
            Caption         =   "View (F1)"
            Height          =   735
            Index           =   0
            Left            =   8400
            TabIndex        =   17
            Top             =   120
            Width           =   1935
         End
         Begin VB.Image Image2 
            Height          =   720
            Index           =   0
            Left            =   240
            Stretch         =   -1  'True
            Top             =   105
            Width           =   735
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "1st Place With # VOTES"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   1290
            TabIndex        =   18
            Top             =   225
            Width           =   7020
         End
      End
      Begin VB.Image Image4 
         Height          =   2250
         Left            =   840
         Picture         =   "frmMain.frx":4116A
         Stretch         =   -1  'True
         Top             =   930
         Width           =   26085
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left Arrow (BACK)     Right Arrow (NEXT)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4080
      TabIndex        =   12
      Top             =   6960
      Width           =   7665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   6000
      TabIndex        =   11
      Top             =   360
      Visible         =   0   'False
      Width           =   7665
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'screens don't move, unless these
Private Const TIMER_NOCREDITS = 16000
Private Const TIMER_TEMPINFO = 4000

'states of the game operation
Private Const STATE_STARTUP = 0
Private Const STATE_NOCREDIT = 3
Private Const STATE_READY = 1
Private Const STATE_WORKING = 2
Private Const STATE_ERROR = 3

'screen gui constants
Private Const TAB_STARTUP = 0
Private Const TAB_NOCREDIT = 1
Private Const TAB_LEADERBOARD = 2
Private Const TAB_GENERATE = 3
Private Const TAB_VIEWIMAGE = 4
Private Const TAB_VOTEONIMAGE = 5

'amount of credits used for tasks
Private Const CREDIT_TOGENERATE = 2
Private Const CREDIT_TOVOTE = 1

'for the incomming UUCode Image
Private ImageText As String
Private ImageName As String


'used for recieving the
'image from python stdout
Private uu As UUCode
Private ss As Stream


'about the interface
Private State As Integer
Private OnTab As Integer
Private Credit As Integer
Private ImgCount As Long

'textbox constants
Private Const Text1GreyText = "(Enter your idea in text here, to generate into an image)"
Private Const Text2GreyText = "(Optionally, enter ideas you don't want in the iamge here)"

'for when sending to the python script
Private Const ImageHeight = 512
Private Const ImageWidth = 512
Private Const Steps = 20
Private Const Seed = -1

'external cmd.exe app redirects
Private WithEvents oImager As RedirectLib.Application
Attribute oImager.VB_VarHelpID = -1
Private WithEvents oLaunch As RedirectLib.Application
Attribute oLaunch.VB_VarHelpID = -1
Private bExecute As Boolean
Private bExecute2 As Boolean

'sound collections
Private CoinDrop As New Collection
Private Generate As New Collection
Private Ambient As New Collection

'the next three functions help stop flickering
Private Sub SetEnabled(ByRef ctrl As Control, ByVal Ena As Boolean)
    If ctrl.Enabled <> Ena Then ctrl.Enabled = Ena
End Sub
Private Sub SetVisible(ByRef ctrl As Control, ByVal Vis As Boolean)
    If ctrl.Visible <> Vis Then ctrl.Visible = Vis
End Sub
Private Sub SetCaption(ByRef ctrl As Control, ByVal Cap As String)
    If ctrl.Caption <> Cap Then ctrl.Caption = Cap
End Sub

'all control (buttons etc...) is bipassed to this function, there is no mouse
Private Function KeyHandler(KeyCode As Integer, Shift As Integer) As Boolean

    'If Shift = 0 Then
        KeyHandler = True
        Select Case KeyCode
        #If USEESC = -1 Then
            Case 27 'ESCAPE

                SaveSettings
                Uninstall
                Unload frmMain
    
        #End If
            Case 8
                KeyCode = 0
            Case 13 'ENDER
                If OnTab = TAB_GENERATE Then Command2_Click
                KeyCode = 0
            Case 17 'CTRL
                If OnTab = TAB_VIEWIMAGE Then Command3_Click
            Case 18 'ALT
                If OnTab = TAB_VOTEONIMAGE Then Command4_Click
            Case 33 'PAGEUP
                Command0_Click
            Case 34 'PAGEDOWN
                Command1_Click
            Case 37 'LEFT
                If OnTab = TAB_VOTEONIMAGE Then MoveImageLeft
                
            Case 38 'UP
            Case 40 'DOWN
            Case 39 'RIGHT
                If OnTab = TAB_VOTEONIMAGE Then MoveImageRight
                
            Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123 'F1-F12
                ViewTopChats (KeyCode - 111)
                
        #If USECOIN = 0 Then
            Case 45 'INSERT
                CoinIn
        #End If
           
            Case Else
                KeyHandler = False
    
        End Select

    'End If
End Function

'lockdown requirements
Private Sub EnableUI()

    SetEnabled Text1, ((WebUIURL <> "") And (State = STATE_READY)) And (OnTab = TAB_GENERATE)
    SetEnabled Text2, ((WebUIURL <> "") And (State = STATE_READY)) And (OnTab = TAB_GENERATE)

    SetEnabled Command2, ((WebUIURL <> "") And (State = STATE_READY)) And (OnTab = TAB_GENERATE)
    SetEnabled Command3, ((WebUIURL <> "") And (State = STATE_READY)) And (OnTab = TAB_VIEWIMAGE)
    SetEnabled Command4, ((WebUIURL <> "") And (State = STATE_READY)) And (OnTab = TAB_VOTEONIMAGE)
    
    SetEnabled Command0, ((WebUIURL <> "") And (State = STATE_READY))
    SetEnabled Command1, ((WebUIURL <> "") And (State = STATE_READY))
    
    SetCaption Command2, "Generate (ENTER) [Requires 2 Credits, You Have " & Credit & " Credits]"
    SetCaption Command4, "Vote (ALT) [Requires 1 Credits, You Have " & Credit & " Credits]"
    
End Sub

'state changer with some restrictive measure
Public Sub SetState(ByVal NewState As Integer)
    If State = STATE_STARTUP Then Me.Cls
    
    State = NewState
    
    If (State = STATE_NOCREDIT) Or (State = STATE_READY And Credit = 0) Then
        State = STATE_NOCREDIT
        SetEnabled Timer1, True
        
    ElseIf (State <> STATE_NOCREDIT) And (Credit > 0) Then
        SetEnabled Timer1, False
        
    End If
    
    EnableUI
End Sub

'similar to a tab control, this changes the screen
Public Sub ShowTab(ByVal cIndex As Integer)

    OnTab = cIndex
    
    SetVisible Picture3, (OnTab = TAB_NOCREDIT)
    
    
    SetVisible Frame3, (OnTab = TAB_LEADERBOARD)
    
    SetVisible Frame1, (OnTab = TAB_GENERATE)
    SetVisible Frame2, (OnTab = TAB_GENERATE)

    SetVisible Label1, (OnTab = TAB_VIEWIMAGE) Or (OnTab = TAB_VOTEONIMAGE)
    SetVisible Picture1, (OnTab = TAB_VIEWIMAGE)

    SetVisible Picture4, (OnTab = TAB_VOTEONIMAGE)
    SetVisible Label2, (OnTab = TAB_VOTEONIMAGE)
    
    SetVisible Command0, True
    SetVisible Command1, True
    SetVisible Command2, (OnTab = TAB_GENERATE) And (Credit > 0)
    SetVisible Command3, (OnTab = TAB_VIEWIMAGE) And (Credit > 0)
    SetVisible Command4, (OnTab = TAB_VOTEONIMAGE) And (Credit > 0)
    
    EnableUI

    If OnTab = TAB_GENERATE Then
        Text1_Change
        Text2_Change
        On Error Resume Next
        Text1.SetFocus
    ElseIf OnTab = TAB_VOTEONIMAGE Then
        RefreshGallery
        On Error Resume Next
        Picture4.SetFocus
    ElseIf OnTab = TAB_VIEWIMAGE Then
    
        If (Picture1.Tag = 0) Then
            SetCaption Label1, "No generated image yet!"
        ElseIf (Picture1.Tag = 1) Then
            SetCaption Label1, "Your Genreated Image:"
        ElseIf (Picture1.Tag = 2) Then
            SetCaption Label1, "Last Genreated Image:"
        End If
        Set Image1.Picture = Nothing
        On Error Resume Next
        Picture1.SetFocus
    ElseIf OnTab = TAB_LEADERBOARD Then
        SetCaption Label1, "Last Genreated Image:"
        ViewWinner
    ElseIf OnTab = TAB_NOCREDIT Then
        If Picture1.Tag = 1 Then
            Picture1.Tag = 2
        End If
    End If

End Sub

'loads the generated image cache
'into a image control collection
Private Sub RefreshGallery()
    On Error GoTo errout:
    
    Dim rs As New ADODB.Recordset
    Dim b() As Byte
    Dim votes As Long
    
    Dim ids As String
    Dim fi As Long
   
    Dim cnt As Long
    For cnt = Image5.LBound To Image5.UBound
        If Image5(cnt).Visible Then
            fi = cnt
            Exit For
        End If
    Next

    rsQuery rsFile, "SELECT * FROM Files ORDER BY FileVote DESC, FileName ASC;"

    For cnt = Image5.LBound + 1 To Image5.UBound
        Unload Image5(cnt)
    Next
    
    If Not rsFile.EOF Then
        rsFile.MoveFirst
        Do Until rsFile.EOF
            
            ids = ids & rsFile("ID") & ","
    
            rsFile.MoveNext
        Loop
    End If
    dbClose rsFile
    
    Dim pic As StdPicture
    Dim pic2 As StdPicture
    
    cnt = 0
    
    Do Until ids = ""
        
        If cnt > 0 Then
            Load Image5(cnt)
        End If
        
        FileGetArray CLng(NextArg(ids, ",")), b, votes
         
        Set pic = PictureFromByteStream(b)
        
        Set Image5(cnt).Picture = pic
        Image5(cnt).Tag = RemoveNextArg(ids, ",")
        
        cnt = cnt + 1
        
    Loop
    
    If cnt > 0 Then
        SetCaption Label2, "Left Arrow (BACK)     Right Arrow (NEXT)"
        
        If fi >= Image5.LBound And fi <= Image5.UBound Then
        
            SetCaption Label1, "Image Election Gallery: (" & Trim(CStr(fi + 1)) & " of " & Trim(CStr(Image5.Count)) & ")"
            For cnt = Image5.LBound To Image5.UBound
                SetVisible Image5(cnt), (fi = cnt)
            Next
    
        End If
    Else
        SetCaption Label2, ""
        SetCaption Label1, "No images in the gallery!"
    End If
errout:
    If Err Then
        MyDebugPrint "RefreshGallery() Error: " & Err.Description
        Err.Clear
    End If
End Sub


Private Sub Command0_Click()
    If Credit = 0 Then
        Select Case OnTab
            Case TAB_NOCREDIT
                ShowTab TAB_LEADERBOARD
            Case TAB_LEADERBOARD
                ShowTab TAB_NOCREDIT
            Case TAB_VIEWIMAGE
                ShowTab TAB_LEADERBOARD
        End Select
    ElseIf State = STATE_READY Then
        Select Case OnTab
            Case TAB_LEADERBOARD
                ShowTab TAB_VOTEONIMAGE
            Case TAB_VOTEONIMAGE
                ShowTab TAB_VIEWIMAGE
            Case TAB_VIEWIMAGE
                ShowTab TAB_GENERATE
            Case TAB_GENERATE
                ShowTab TAB_LEADERBOARD
        End Select
    End If
    If Timer1.Enabled Then
        Timer1.Enabled = False
        Timer1.Interval = TIMER_NOCREDITS
        Timer1.Enabled = True
    End If
End Sub
Private Sub Command1_Click()
    If Credit = 0 Then
        Select Case OnTab
            Case TAB_NOCREDIT
                ShowTab TAB_LEADERBOARD
            Case TAB_LEADERBOARD
                ShowTab TAB_NOCREDIT
            Case TAB_VIEWIMAGE
                ShowTab TAB_LEADERBOARD
        End Select
    ElseIf State = STATE_READY Then
        Select Case OnTab
            Case TAB_LEADERBOARD
                ShowTab TAB_GENERATE
            Case TAB_GENERATE
                ShowTab TAB_VIEWIMAGE
            Case TAB_VIEWIMAGE
                ShowTab TAB_VOTEONIMAGE
            Case TAB_VOTEONIMAGE
                ShowTab TAB_LEADERBOARD
        End Select
    End If
    If Timer1.Enabled Then
        Timer1.Enabled = False
        Timer1.Interval = TIMER_NOCREDITS
        Timer1.Enabled = True
    End If
End Sub

Private Sub Command0_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub
Private Sub Command1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub
Private Sub Command3_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub
Private Sub Command4_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub

Private Sub Command0_GotFocus()
    CommandFocus
End Sub
Private Sub Command1_GotFocus()
    CommandFocus
End Sub
Private Sub Command2_GotFocus()
    CommandFocus
End Sub
Private Sub Command3_GotFocus()
    CommandFocus
End Sub
Private Sub Command4_GotFocus()
    CommandFocus
End Sub
Private Sub CommandFocus()
    On Error Resume Next
    Me.SetFocus
End Sub

Private Sub Command2_Click()
    CreateImage
End Sub

Private Sub Command3_Click()
    ShowTab TAB_VOTEONIMAGE
End Sub

Private Sub Command4_Click()
    VoteOnImage
End Sub

Private Sub SetAllFonts()
    Me.Font = "Transformers Movie"
    If Me.Font = "Transformers Movie" Then
        Dim ctrl As Control
        For Each ctrl In Me.Controls
            Select Case TypeName(ctrl)
                Case "Label", "CommandButton", "Frame", "TextBox"
                    ctrl.Font = Me.Font
                Case "Image", "PictureBox", "Timer"
                Case Else
            End Select
        Next
    End If
End Sub

Private Sub Command5_Click(Index As Integer)
    ViewTopChats (Index + 1)
End Sub

Private Sub Command5_GotFocus(Index As Integer)
    CommandFocus
End Sub

Private Sub Command5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub

Private Sub Form_Initialize()
    
    On Error Resume Next
    
    Set oLaunch = New RedirectLib.Application
    oLaunch.BufferSize = 8192
    oLaunch.Wait = 1
    Set oImager = New RedirectLib.Application
    oImager.BufferSize = 32767
    oImager.Wait = 1
    
    If Err Then
        Debug.Print "Form_Initialize() Error: " & Err.Number & " " & Err.Description
        Err.Clear
        End
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub


Private Sub Form_Load()

    On Error Resume Next
    Picture1.Tag = 0
    
    SetAllFonts
    
    LoadMusicFiles CoinDrop, "coindrop"
    LoadMusicFiles Generate, "generate"
    LoadMusicFiles Ambient, "ambient"
    
    If bExecute Then
        oLaunch.Stop
        bExecute = False
    End If

#If VBIDE = 0 Or USESD = -1 Then
    oLaunch.Name = GetSystem32Folder & "\cmd.exe" 'say where the cmd.exe is or the command.exe
    Select Case oLaunch.Start 'Starts the connection to the command prompt
        Case laAlreadyRunning
            'Already going
            bExecute = True
        Case laWindowsError
            MyDebugPrint "Form_Load() Error1: " & CStr(oLaunch.LastErrorNumber) & "!" 'if there was a problem.
        Case laOk
            bExecute = True 'Everything went smooth, we are now connected to cmd.exe
    End Select

#End If

    If Err Then
        MyDebugPrint "Form_Load() Error2: " & Err.Number & " " & Err.Description
        Err.Clear
        End
    End If
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then Cancel = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Picture3.Left = 0
    Picture3.Top = 0
    Picture3.Width = Me.ScaleWidth
    Picture3.Height = Me.ScaleHeight - Command0.Height

    Image3(0).Left = 0
    Image3(0).Top = 0
    Image3(0).Width = Picture3.Width
    Image3(0).Height = (Picture3.Height \ 3)
    
    Image3(1).Left = 0
    Image3(1).Top = Image3(0).Height
    Image3(1).Width = Picture3.Width
    Image3(1).Height = (Picture3.Height \ 3)
    
    Image3(2).Left = 0
    Image3(2).Top = (Image3(0).Height * 2)
    Image3(2).Width = Picture3.Width
    Image3(2).Height = (Picture3.Height \ 3)
        
    Command0.Top = Me.ScaleHeight - Command0.Height
    Command1.Top = Command0.Top
    Command2.Top = Command0.Top - Command2.Height
    Command3.Top = Command0.Top - Command3.Height
    Command4.Top = Command0.Top - Command4.Height
    
    Command0.Left = 0
    Command0.Width = Me.ScaleWidth / 2
    Command1.Left = Me.ScaleWidth / 2
    Command1.Width = Me.ScaleWidth / 2
    
    Command2.Left = 0
    Command2.Left = 0
    Command3.Left = 0
    Command4.Left = 0
    
    Command2.Width = Me.ScaleWidth
    Command3.Width = Me.ScaleWidth
    Command4.Width = Me.ScaleWidth

    Frame3.Top = 0
    Frame3.Left = 0
    Frame3.Width = Me.ScaleWidth
    Frame3.Height = (Me.ScaleHeight - Command1.Height)
    
    Image4.Top = 0
    Image4.Left = 0
    Image4.Width = Me.ScaleWidth
    
    Picture2.Left = (Frame3.Width / 2) - (Picture2.Width / 2)
    Picture2.Top = ((Frame3.Height / 2) + (Image4.Height / 2)) - (Picture2.Height / 2)
    
    Frame1.Height = ((Frame3.Height - Command1.Height) / 2)
    Frame2.Height = Frame1.Height
    Frame1.Width = Me.ScaleWidth
    Frame2.Width = Me.ScaleWidth
    Frame1.Top = 0
    Frame1.Left = 0
    Frame2.Top = Frame1.Height
    Frame2.Left = 0
    
    Text1.Width = Frame1.Width - (Text1.Left * 2)
    Text2.Width = Frame2.Width - (Text2.Left * 2)
    Text1.Height = Frame1.Height - Text1.Top - Text1.Left
    Text2.Height = Frame2.Height - Text2.Top - Text2.Left
        
    Picture1.Width = ImageWidth * Screen.TwipsPerPixelX
    Picture1.Height = ImageHeight * Screen.TwipsPerPixelY
    
    Picture1.Top = ((Me.ScaleHeight - Command1.Height) / 2) - (Picture1.Height / 2)
    Picture1.Left = (Me.ScaleWidth / 2) - (Picture1.Width / 2)
    
    Image1.Top = 0
    Image1.Left = 0
    Image1.Width = Picture1.Width
    Image1.Height = Picture1.Height
    
    Label1.Left = Picture1.Left
    Label1.Width = Picture1.Width
    Label1.Top = Picture1.Top - Label1.Height

    Label2.Left = Picture1.Left
    Label2.Width = Picture1.Width
    Label2.Top = Picture1.Top + Picture1.Height
    
    Picture4.Top = Picture1.Top
    Picture4.Left = Picture1.Left
    Picture4.Width = Picture1.Width
    Picture4.Height = Picture1.Height
    
    Label4.Top = (Picture1.Height / 2) - (Label4.Height / 2)
    Label4.Left = (Picture1.Width / 2) - (Label4.Width / 2)
    
    Dim cnt As Long
    For cnt = Image5.LBound To Image5.UBound
        Image5(cnt).Top = 0
        Image5(cnt).Left = 0
        Image5(cnt).Width = Picture4.Width
        Image5(cnt).Height = Picture4.Height
    Next
       
    If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    #If VBIDE = -1 And USESD = 0 Then

    #Else
        KillSubApps GetCurrentProcessId
    #End If
    
    If bExecute Then
        oLaunch.Stop
        bExecute = False
    End If

    If bExecute2 Then
        oImager.Stop
        bExecute2 = False
    End If

End Sub


Private Sub Form_Terminate()
                
    ClearMusicCollection CoinDrop
    ClearMusicCollection Generate
    ClearMusicCollection Ambient
    
    Set oImager = Nothing
    Set oLaunch = Nothing
End Sub


'the following two functions move the image viewed
'in the gallery to the image left or right of it
Private Function MoveImageRight() As Long
    If Image5.Count > 1 Then
        Dim cnt As Long
        For cnt = Image5.LBound To Image5.UBound
            If Image5(cnt).Visible Then
                If cnt = Image5.UBound Then
                    Image5(Image5.LBound).Visible = True
                    MoveImageRight = (Image5.LBound + 1)
                Else
                    Image5(cnt + 1).Visible = True
                    MoveImageRight = ((cnt + 1) + 1)
                End If
                Image5(cnt).Visible = False
                Exit For
            End If
        Next
        Label1.Caption = "Image Election Gallery: (" & Trim(CStr(MoveImageRight)) & " of " & Trim(CStr(Image5.Count)) & ")"
    End If
End Function

Private Function MoveImageLeft() As Long
    If Image5.Count > 1 Then
        Dim cnt As Long
        For cnt = Image5.LBound To Image5.UBound
            If Image5(cnt).Visible Then
                If cnt = Image5.LBound Then
                    Image5(Image5.UBound).Visible = True
                    MoveImageLeft = (Image5.UBound + 1)
                Else
                    Image5(cnt - 1).Visible = True
                    MoveImageLeft = cnt
                End If
                Image5(cnt).Visible = False
                Exit For
            End If
        Next
            Label1.Caption = "Image Election Gallery: (" & Trim(CStr(MoveImageLeft)) & " of " & Trim(CStr(Image5.Count)) & ")"
    End If
End Function

Private Function GetImageID() As Long
    Dim cnt As Long
    For cnt = Image5.LBound To Image5.UBound
        If Image5(cnt).Visible Then
            GetImageID = Image5(cnt).Tag
            Exit For
        End If
    Next
End Function

'this function should execute when the
'coin goes in the machine, for credits
Public Sub CoinIn()
    Timer1.Enabled = False
    If CoinDrop.Count > 0 Then
        CoinDrop(RandomPositive(1, CoinDrop.Count)).PlaySound
    End If
    
    Credit = Credit + (CREDIT_TOGENERATE + CREDIT_TOVOTE)
    SetState STATE_READY
    ShowTab TAB_GENERATE
End Sub


'the next two functions are the game
'and the functions that take credits
Private Sub VoteOnImage()
    If Credit >= 1 Then
        Dim ID As Long
        ID = GetImageID
        If ID <> 0 Then
            If Label2.Caption <> "You've added 1 vote for this image!" Then
                Credit = Credit - CREDIT_TOVOTE
                SetState STATE_WORKING
                FileVoteFor ID
                SetState STATE_READY
                Timer1.Enabled = False
                Timer1.Interval = TIMER_TEMPINFO
                Label2.Caption = "You've added 1 vote for this image!"
                Timer1.Enabled = True
            End If
        End If
    ElseIf Credit = 0 Then 'show a red error
        Command4.BackColor = &HFF&
        Timer1.Enabled = False
        Timer1.Interval = TIMER_TEMPINFO
        Timer1.Enabled = True
    End If
End Sub

Private Sub ViewTopChats(ByVal Place As Long)
    If Place <= TopNumberOf Then

        ShowTab TAB_VIEWIMAGE
        ViewWinner Place
        
        If Timer1.Enabled Then
            Timer1.Enabled = False
            Timer1.Interval = TIMER_NOCREDITS
            Timer1.Enabled = True
        End If
    End If
End Sub

Private Sub CreateImage()

    If Credit >= 2 Then
        Credit = Credit - CREDIT_TOGENERATE
    
        SetState STATE_WORKING
    
        If Generate.Count > 0 Then
            Generate(RandomPositive(1, Generate.Count)).PlaySound
        End If
        
        ImageText = ""
        Set uu = New UUCode
        Set ss = New Stream
        
        If bExecute2 Then
            oImager.Stop
            bExecute2 = False
        End If
    
        oImager.Name = "python.exe """ & SDPath & "Txt2Img2Txt" & IIf(InStr(WebUIURL, "https") > 0, "_SSL", "") & ".py"" " & Seed & " " & Steps & " " & ImageWidth & " " & ImageHeight & " """ & IIf(Text1.Text <> "", Replace(Text1.Text, """", "\"""), " ") & """ """ & IIf(Text2.Text <> "", Replace(Text2.Text, """", "\"""), " ") & """" & vbCrLf
            
        #If VBIDE = 0 Or USESD = -1 Then
            Select Case oImager.Start 'Starts the connection to the command prompt
                Case laAlreadyRunning
                    'Already going
                    bExecute2 = True
                Case laWindowsError
                    MyDebugPrint "CreateImage() Error: " & CStr(oImager.LastErrorNumber) & "!" 'if there was a problem.
                Case laOk
                    bExecute2 = True 'Everything went smooth, we are now connected to cmd.exe
            End Select
        #Else
            oImager_DataReceived "begin 0 " & Replace(modGuid.GUID, "-", "") & ".bmp" & vbCrLf
            oImager_DataReceived "end" & vbCrLf
        #End If
    
    ElseIf Credit = 0 Then 'show a red error
        Command2.BackColor = &HFF&
        Timer1.Enabled = False
        Timer1.Interval = TIMER_TEMPINFO
        Timer1.Enabled = True
    End If
End Sub

'the following two functions are for the music content collections
'you can add as many variant sounds as you like by the naming scheme _#
Private Sub LoadMusicFiles(ByRef col As Collection, ByVal Prefix As String)
    Dim cnt As Long
    cnt = 1
    Dim plr As Player
     
    Do While PathExists(GetMyMusicFolder & "\SoundFX\" & Prefix & "_" & Trim(CStr(cnt)) & ".mp3", True)
        Set plr = New Player
        plr.Filename = GetMyMusicFolder & "\SoundFX\" & Prefix & "_" & Trim(CStr(cnt)) & ".mp3"
        col.Add plr
        Set plr = Nothing
        cnt = cnt + 1
    Loop
    
End Sub

Private Sub ClearMusicCollection(ByRef col As Collection)
    Dim plr As Player
    Do While col.Count > 0
        Set plr = col(1)
        col.Remove 1
        If plr.IsPlaying Then plr.StopSound
        Set plr = Nothing
    Loop
    Set col = Nothing
End Sub



Public Sub StartUp()
    On Error Resume Next
                
    SetState STATE_STARTUP

    ShowTab TAB_STARTUP
   
    Me.Cls
    Me.Print "Starting up..."
    
    Form_Resize
    
    #If VBIDE = 0 Or USESD = -1 Then
        
        oLaunch.Write """" & SDPath & "mywebui.bat""" & vbCrLf
    #Else
        'fake the startup of stable-diffusion
        oLaunch_DataReceived "Running on local URL: https://127.0.0.1:7680"
    #End If

    If Err Then
        MyDebugPrint "StartUp() Error: " & Err.Number & " " & Err.Description
        Err.Clear
    End If
End Sub


'Private Sub ResetSD()
'
'    KillSubApps GetCurrentProcessId
'
'    If bExecute Then
'        oLaunch.Stop
'        bExecute = False
'    End If   'if it is, then it stops the connection.
'
'    WebUIURL = ""
'    LastError = ""
'
'    Select Case oLaunch.Start 'Starts the connection to the command prompt
'        Case laAlreadyRunning
'            'Already going
'            bExecute = True
'        Case laWindowsError
'            MyDebugPrint "ResetSD() Error: " & CStr(oLaunch.LastErrorNumber) & "!" 'if there was a problem.
'        Case laOk
'            bExecute = True 'Everything went smooth, we are now connected to cmd.exe
'    End Select
'
'    StartUp
'End Sub

'Private Sub WebUI()
'    OpenWebsite WebUIURL
'End Sub

Private Sub oImager_DataReceived(ByVal sData As String)
    Static process As Boolean
      
    ImageText = ImageText & sData

    Dim inline As String
    
    Do While InStr(ImageText, vbCrLf) > 0
    
        inline = RemoveNextArg(ImageText, vbCrLf, , False)
        If Left(inline, 5) = "begin" And process = False Then
            process = True
            
            ImageName = RemoveArg(RemoveArg(inline, " "), " ")
            
            SetState STATE_WORKING
             
        ElseIf Left(inline, 3) = "end" And process = True Then
            process = False
           
            Dim pic As StdPicture
            If ss.Length = 0 Then ss.Concat LoadResData(1, "BMP")
            
            Set pic = PictureFromByteStream(ss.Partial)
            
            Dim ID As Long
            SetVisible Label4, False
            
            ID = FilePutArray(Replace(ImageName, ".bmp", ""), ss.Partial, IIf(Text1.Text = Text1GreyText, "", Text1.Text), IIf(Text2.Text = Text2GreyText, "", Text2.Text))
            
            Picture1.Tag = 1
            
            ShowTab TAB_VIEWIMAGE
            
            Set Image1.Picture = pic
            Set Picture1.Picture = pic
    
            Set pic = Nothing
            
            Set ss = Nothing
            Set uu = Nothing
            
            bExecute2 = False
            
            Dim X As Single
            Dim Y As Single
            Dim C As Boolean
            For X = 0 To (Picture1.ScaleWidth / Screen.TwipsPerPixelX) - 1
                For Y = 0 To (Picture1.ScaleHeight / Screen.TwipsPerPixelY) - 1
                    If Not (Picture1.Point(X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY) = vbBlack) Then
                        C = True
                        Exit For
                    End If
                Next
                If C Then Exit For
            Next
            SetState STATE_READY
            If Not C Then
                FileRemove ID
                Credit = Credit + CREDIT_TOGENERATE
                SetVisible Label4, True
            Else
                SetVisible Label4, False
            End If
            
            
            If Timer1.Enabled Then
                SetEnabled Timer1, False
                Timer1.Interval = TIMER_NOCREDITS
                SetEnabled Timer1, True
            End If

        ElseIf process Then
            
            ss.Concat StrConv(uu.Decode(Right(inline, Len(inline) - 1)), vbFromUnicode)
        
        End If
        
    Loop
        
End Sub

Private Sub oImager_ProcessEnded()
    bExecute2 = False
End Sub

Private Sub oLaunch_ProcessEnded()
    bExecute = False
End Sub

Private Sub oLaunch_DataReceived(ByVal sData As String)
    Dim intext As String

    intext = sData
    
    Dim inline As String

    Do While intext <> ""

        inline = RemoveNextArg(intext, vbCrLf)
        
        If inline Like "*Running on local URL:*" And State = STATE_STARTUP Then
            If Right(Trim(NextArg(inline, "://")), 5) = "https" Then
                WebUIURL = "https://" & NextArg(NextArg(Trim(RemoveArg(inline, "://")), " "), vbCrLf)
            ElseIf Right(Trim(NextArg(inline, "://")), 4) = "http" Then
                WebUIURL = "http://" & NextArg(NextArg(Trim(RemoveArg(inline, "://")), " "), vbCrLf)
            End If
            SetState STATE_NOCREDIT
            ShowTab TAB_NOCREDIT
        ElseIf inline Like "*http://*" And State = STATE_STARTUP Then
            WebUIURL = "http://" & NextArg(NextArg(Trim(RemoveArg(inline, "http://")), vbCrLf), "/")
            SetState STATE_NOCREDIT
            ShowTab TAB_NOCREDIT
        ElseIf inline Like "*https://*" And State = STATE_STARTUP Then
            WebUIURL = "https://" & NextArg(NextArg(Trim(RemoveArg(inline, "https://")), vbCrLf), "/")
            SetState STATE_NOCREDIT
            ShowTab TAB_NOCREDIT
            
        ElseIf inline Like "*Traceback*" And State = STATE_WORKING Then
            LastError = "The program has thrown an exception"
             SetState STATE_ERROR
        ElseIf inline Like "*Press any key to continue . . .*" And State = STATE_WORKING Then
            LastError = "The program has thrown an exception"
             SetState STATE_ERROR
        ElseIf inline Like "*Error:*" And State = STATE_WORKING Then
            LastError = "The program has thrown an exception"
             SetState STATE_ERROR
'        ElseIf inline Like "*100%|*" And State = STATE_WORKING Then
'             SetState STATE_READY
        ElseIf inline Like "*%|*" Then
             SetState STATE_WORKING
        End If

    Loop

End Sub

Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub

Private Sub Picture2_GotFocus()
    CommandFocus
End Sub

Private Sub Picture2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub

Private Sub Picture3_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub

Private Sub Picture4_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyHandler KeyCode, Shift
End Sub




Private Sub TextClickEntry(ByRef Text As TextBox, ByVal GreyText As String)
    If Text.Text = GreyText Then Text.SelStart = 0
End Sub

Private Sub TextChangeEntry(ByRef Text As TextBox, ByVal GreyText As String)
    If Text.Visible Then
        If Text.ForeColor = &H80000008 Then
            If Text.Text = "" Then
                Text.ForeColor = &HC0C0C0
                Text.Text = GreyText
                Text.SelStart = 0
            End If
        End If
    End If
End Sub

Private Sub TextKeyDownEntry(ByRef Text As TextBox, ByVal GreyText As String, KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case 46, 8
            If Text.Text = GreyText Then KeyCode = 0
    End Select
    
    If Replace(Text.Text, GreyText, "") <> Text.Text And Replace(Text.Text, GreyText, "") <> "" And KeyCode <> 8 Then
        Dim backup As Integer
        backup = Text1.SelStart
        Text.Text = Replace(Text.Text, GreyText, "")
        If backup <= Len(Text.Text) Then Text.SelStart = backup
        Text.ForeColor = &H80000008
    End If
    
    If Text.Text = GreyText Then Text.SelStart = 0
End Sub

Private Sub TextKeyPressEntry(ByRef Text As TextBox, ByVal GreyText As String, KeyAscii As Integer)
    If Text.Visible Then
        If Text.Text = GreyText Then
            Text.ForeColor = &HC0C0C0
        ElseIf Not Text.Text = GreyText Then
            Text.ForeColor = &H80000008
        End If
        If Replace(Text.Text, GreyText, "") <> Text.Text And KeyAscii <> 8 Then
            Text.Text = Replace(Text.Text, GreyText, "")
            Text.ForeColor = &H80000008
        End If
    End If
End Sub

Private Sub TextKeyUpEntry(ByRef Text As TextBox, ByVal GreyText As String, KeyCode As Integer, Shift As Integer)
    If Text.Visible Then
        If Replace(Text.Text, GreyText, "") <> Text.Text And Replace(Text.Text, GreyText, "") <> "" And KeyCode <> 8 Then
            Dim backup As Integer
            backup = Text1.SelStart
            Text.Text = Replace(Text.Text, GreyText, "")
            If backup <= Len(Text.Text) Then Text.SelStart = backup
            Text.ForeColor = &H80000008
        End If
    End If
End Sub

Private Sub Text1_Change()
    TextChangeEntry Text1, Text1GreyText
End Sub

Private Sub Text1_Click()
    TextClickEntry Text1, Text1GreyText
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not KeyHandler(KeyCode, Shift) Then
        TextKeyDownEntry Text1, Text1GreyText, KeyCode, Shift
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    TextKeyPressEntry Text1, Text1GreyText, KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    TextKeyUpEntry Text1, Text1GreyText, KeyCode, Shift
End Sub

Private Sub Text2_Change()
    TextChangeEntry Text2, Text2GreyText
End Sub

Private Sub Text2_Click()
    TextClickEntry Text2, Text2GreyText
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not KeyHandler(KeyCode, Shift) Then
        TextKeyDownEntry Text2, Text2GreyText, KeyCode, Shift
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    TextKeyPressEntry Text2, Text2GreyText, KeyAscii
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
    TextKeyUpEntry Text2, Text2GreyText, KeyCode, Shift
End Sub

'this function loads the image at place in the top whatever
'and it refreshes the top whatever of images when place = -1
Private Sub ViewWinner(Optional ByVal Place As Long = -1)
    On Error GoTo errout:

    FileGetTopVotes
    
    Dim tmp As Long
    Dim votes As Long
    Dim ids As String
    
    Dim pic As StdPicture
    Dim b() As Byte
    
    If Place = -1 Then
        For tmp = Image2.LBound + 1 To Image2.UBound
            Unload Image2(tmp)
        Next
        For tmp = Label3.LBound + 1 To Label3.UBound
            Unload Label3(tmp)
        Next
        For tmp = Command5.LBound + 1 To Command5.UBound
            Unload Command5(tmp)
        Next
        Picture2.Height = ((Image2(0).Top * 2) + Image2(0).Height)
        Label3(0).Top = (Image2(0).Top + (Image2(0).Height / 2)) - (Label3(0).Height / 2)
        Command5(0).Top = Image2(0).Top
        
        
    End If

    ImgCount = 0
    tmp = 0
    If Not rsFile.EOF Then
        rsFile.MoveFirst
        Do Until tmp = TopNumberOf Or rsFile.EOF
            tmp = tmp + 1
            ids = ids & rsFile("ID") & ","
            ImgCount = ImgCount + 1
            rsFile.MoveNext
        Loop
    End If
    dbClose rsFile
    
    If Place = -1 Then
        If tmp > 0 Then
            Command5(0).Enabled = True
            Image2(Image2.UBound).Visible = True
            Label3(Label3.UBound).Visible = True
            Command5(Command5.UBound).Visible = True
                
            Picture2.Height = ((Image2(0).Top * 2) + Image2(0).Height)
            
            Do Until tmp = 1
                Load Image2(Image2.UBound + 1)
                Load Label3(Label3.UBound + 1)
                Load Command5(Command5.UBound + 1)
                tmp = tmp - 1
                Picture2.Height = Picture2.Height + (Image2(0).Top + Image2(0).Height)
                Image2(Image2.UBound).Top = (Image2(Image2.UBound - 1).Top + Image2(Image2.UBound - 1).Height) + Image2(0).Top
                Label3(Image2.UBound).Top = (Image2(Image2.UBound).Top + (Image2(Image2.UBound).Height / 2)) - (Label3(Image2.UBound).Height / 2)
                Command5(Command5.UBound).Top = Image2(Image2.UBound).Top
                Image2(Image2.UBound).Visible = True
                Label3(Label3.UBound).Visible = True
                Command5(Command5.UBound).Visible = True
                Command5(Command5.UBound).Enabled = True
            Loop
            
            Form_Resize
        Else
            Set Image2(0).Picture = LoadPicture("")
            Label3(0).Caption = "No image votes yet!"
            Command5(0).Enabled = False
        End If
   End If

    tmp = 0
    
    Do Until ids = ""
        tmp = tmp + 1
        If tmp = Place Then
            FileGetArray NextArg(ids, ","), b, votes
            Set pic = PictureFromByteStream(b)
            Set Image1.Picture = pic
            Label1.Caption = "Number " & Place & " in the poles with " & votes & " votes."
        ElseIf Place = -1 Then
            FileGetArray NextArg(ids, ","), b, votes
            
            Select Case tmp
                Case 1
                    Label3(tmp - 1).Caption = "1st place with " & votes & " votes (F1)"
                    
                Case 2
                    Label3(tmp - 1).Caption = "2nd place with " & votes & " votes (F2)"
                Case 3
                    Label3(tmp - 1).Caption = "3rd place with " & votes & " votes (F3)"
                Case Else
                    Label3(tmp - 1).Caption = Trim(CStr(tmp)) & "th place with " & votes & " votes (F" & Trim(CStr(tmp)) & ")"
            End Select
            
            Command5(tmp - 1).Caption = "View (F" & Trim(CStr(tmp)) & ")"
            Set pic = PictureFromByteStream(b)
            Set Image2(tmp - 1).Picture = pic
        End If
        RemoveNextArg ids, ","
    Loop
    Set pic = Nothing
    Erase b
         
errout:
    If Err Then
        MyDebugPrint "ViewWinner() Error: " & Err.Description
        Err.Clear
    End If
End Sub

'gui timer
Private Sub Timer1_Timer()
    If Me.Visible Then
    
        Static TopOf As Integer
        
        If Timer1.Interval <> TIMER_TEMPINFO Then
            Timer1.Interval = TIMER_TEMPINFO
        End If
        
        If Credit = 0 Then
            Select Case OnTab
                Case TAB_NOCREDIT
                    ShowTab TAB_LEADERBOARD
                Case TAB_LEADERBOARD
                    TopOf = TopOf + 1
                    
                    ShowTab TAB_VIEWIMAGE
                
                    ViewWinner TopOf
                    
                Case TAB_VIEWIMAGE
                    If TopOf = TopNumberOf Or TopOf = ImgCount Then TopOf = 0
                    ShowTab TAB_NOCREDIT
                    
                Case TAB_VOTEONIMAGE, TAB_GENERATE
                    ShowTab TAB_LEADERBOARD
            End Select
        End If
        
        Label2.Caption = "Left Arrow (BACK)     Right Arrow (NEXT)"
        
        Select Case OnTab
            Case TAB_GENERATE
                Command2.BackColor = &H8000000F
            Case TAB_VOTEONIMAGE
                Command4.BackColor = &H8000000F
        End Select


        If RandomPositive(1, 5) = 1 Then
            If Ambient.Count > 0 Then
                Ambient(RandomPositive(1, Ambient.Count)).PlaySound
            End If
        End If
    End If
    
End Sub

'important timer
Private Sub Timer2_Timer()

    TestForUSBKey
    
    #If USECOIN = -1 Then
        Static coinop As Boolean
        If CoinCheck Then
            If Not coinop Then
                coinop = True
                CoinIn
            End If
        Else
            coinop = False
        End If
    #End If
End Sub
