Attribute VB_Name = "modMain"
#Const modMain = -1
Option Explicit


'################### COMPILE CONDITIONS ###################
'
' In the properties of this project are compile conditions
' that are vital to the operation or debugging of this app
'
' Here is an explination:
'
'  VBIDE = -1/0     Sets whether or not the app is running from the VB IDE environment (=-1) or as a compiled executable (=0)
'
'  SAFEMODE = -1/0  Sets whether or not to put the system into restrictive mode as a Kiosk (=0) (i.e. unable to use most of the
'                   options in CTRL+ALT+DEL) or to keep the Windows 11 system options as normal (=-1) safely allowing TaskMngr
'                   *screen savers, pwoer savers, hot keys, mouse, and everything but accessibility in CTRL+ALT+DEL is disabled.
'
'  USESD = -1/0     Sets whether or not the application should start the Stable-Diffusion Automatic1111 webui (=-1) or
'                   to fake it (=0) (i.e. debigging other portions like graphics not invovled with generating an image).
'                   In the manner the way the application closes when this is enabled, the envirnoment becomes volatile.
'
'  USECOIN = -1/0   Sets whether or not to use the hardware coin drop (=-1), or to fake it (=0) by using the Insert key
'                   The hardware coin drop relies on inpout32.dll and a parallel port with a data pin and ground connected
'                   to the coin drop box switch. See modePeekPoke for more information on the values and settings it up.
'
'  USEESC = -1/0    Sets whether or not the ESCAPE key will immediatly close the program (=-1) or the ESCAPE key will be ignored (=0)
'
'  During a production compile of this application for an Arcade box, you will want the following compile condition values set:
'
'  VBIDE = 0 : SAFEMODE = 0 : USESD = -1 : USECOIN = -1 : USEESC = 0
'
'  During a development run environment of this application for an Arcade box, I usually use the following compile condition values set:
'
'  VBIDE = -1 : SAFEMODE = -1 : USESD = -1 : USECOIN = 0 : USEESC = -1
'
'  Other compile considtions that are not nessisarily used may be set for different module awareness are the module names as follows:
'
'  modBitValue = -1 : modCOmmon = -1 : modDatabase = -1 : modFiles = -1 : modFolders = -1 : modGraphics = -1 : modGuid = -1
'  modKillVisible = -1 : modKioskApp = -1 : modMain = -1 : modPeekPoke = -1 : modProcess = -1 : modReg64Bit = -1 : modSettings = -1
'
'  Last but not least the project name Txt2ImgKiosk = -1 is again for module awareness if, at all even used.
'
' About setting it up:
'
'  A working automatic1111 webui with webui.bat is required, it will create a similar to webui-user.bat called mywebui.bat
'  The application itself can run from with in the webui folder, or one level above if your automatic1111 folder is named:
'  webui or stable-diffusion-webui or stable-diffusion-webui-master, tweakable in Sub Main(), also, see AppPath for detail.
'
'##########################################################

Public Const GlobalBackColor = &HC0C0C0

Public Const VotePeriod = 28 'this is what a election period will be in days before votes reset, if in period vote mode
Public Const ResetVotes = 14 'during the period, if this many votes are not met, it will not reset votes on the period
                            'and possibly change the terms of the election to be no term, or having a term of VotePeriod

Public Const TopNumberOf = 12  'what number to make the top 10 so to speak, or top 5, anything from 1 to 12
Public Const TotalImages = 50 'as much as you want to be held in the scroll image cache for voting and must
    'be above the TopNumberOf, theoretically significantly, these are not viewable unless you have credits in.

Public Const TaperVoteReset = True 'sets whether or not the top votes are reset to their positions in the rank as their vote count
                                'when the vote term changes from no term to term or term to no term, if false, votes reset to zero

'end customization

Public VotingTerm As String 'if this value is before the current date, then we are in no timeline term voting, else term voting
Public PeriodicVotes As Long 'determined by the voting term from the above variable, when in term voting, this is how many votes
                            'are seen so far in the term, if we are in no term voting, this is how many votes have occured in
                            'the current virtual voting term, that if met, will change the no term voting to term voting. in
                            'term voting, when the term timeline is met and votes are enough it remains term voting then the
                            'votes of every picture are reset to zero, in no term voting the votes don't get reset, as scores.

Public LastError As String
Public WebUIURL As String
Public SDPath As String


Public Sub RegistrySet(ByVal IO As Long)
    'the critical registry entries to change Windows 11 to a restrictive Kiosk mode in a legacy app's effort.
    'due to potential errors while theses are enabled (IO=1) they are called off for errors that are also popup
    'supressed otherwise the potential to lock out the system from recovery requiring a windows reset is high.
    
    RegWriteDWORD HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\System", "DontDisplayNetworkSelectionUI", IO, True
    RegWriteDWORD HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\System", "DontDisplayNetworkSelectionUI", IO, False
    
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", IO, True
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", IO, False

    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", IO, True
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoWinKeys", IO, False
    
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableLockWorkstation", IO, True
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableLockWorkstation", IO, False

    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskmgr", IO, True
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableTaskmgr", IO, False

    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableChangePassword", IO, True
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableChangePassword", IO, False
    
    RegWriteDWORD HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "HideFastUserSwitching", IO, True
    RegWriteDWORD HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "HideFastUserSwitching", IO, False
    
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", IO, True
    RegWriteDWORD HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", IO, False
End Sub

Public Sub Install()
    'the entry point function to setup the app/visibility as a Kiosk or Windows Shell
    'handling absents of Windows features appearences and voiding Sleeps/Screensaver
    
    If Not frmMain.Visible Then
    
    #If VBIDE = 0 Then
        #If SAFEMODE = 0 Then

            RegistrySet 1
            
        #Else

            RegistrySet 0
            
        #End If
            
<<<<<<< HEAD
        RegWriteString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", AppPath & App.EXEName & ".exe", True
        RegWriteString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", AppPath & App.EXEName & ".exe", False
=======
        RegWriteString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", AppPath & "Txt2ImgKiosk.exe", True
        RegWriteString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", AppPath & "Txt2ImgKiosk.exe", False
>>>>>>> de56a4e81b38f5d2d86a0a4069a0f175e8796827
        Do While IsProccessEXERunning("explorer.exe", False)
            KillApp "explorer.exe", False
        Loop
    
       ActivatePowerMonitor
    #End If
    
    frmMain.Show

    End If
End Sub

Public Sub Uninstall()
    'opposite of the above install function in every way
    If frmMain.Visible Then

        frmMain.Hide
                
        '#If VBIDE = 0 Then
        
            DeactivatePowerMonitor

            RegistrySet 0
            
            RegWriteString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe", True
            RegWriteString HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell", "explorer.exe", False
            Do While IsProccessEXERunning("explorer.exe", False)
                KillApp "explorer.exe", False
            Loop
            RunProcess Environ("SystemRoot") & "\Explorer.exe"

        '#End If

    End If

End Sub

Public Function RandomPositive(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'returns a random positive number from LowerBound to UpperBound
    RandomPositive = CLng((UpperBound - LowerBound) * Rnd + LowerBound)
End Function

Public Sub MyDebugPrint(ByVal Msg As String)
    'due to the nature of the Kiosk app legacy mode
    'errors are supressed to a debug text file
    
    Dim FileNum As Integer
    FileNum = FreeFile
    Open AppPath & "debug.txt" For Append As #FileNum
        Print #FileNum, Msg
    Close #FileNum

    Debug.Print Msg
    
    RegistrySet 0
End Sub


Public Sub TestForUSBKey()

    Dim drv As String
    Dim found As Boolean
    Dim tmp As String
    drv = "A"
    Do Until drv = "Z" Or found
        On Error Resume Next
        tmp = Dir(drv & ":")
        If Err.Number = 0 Then
            On Error GoTo 0
            If PathExists(drv & ":\KIOSK", True) Then
                found = True
                If ReadFile(drv & ":\KIOSK") = "6E98DE51-5380-D7AC-D780-5351DE986E6E" Then
                    Uninstall
                End If
            End If
        Else
            Err.Clear
            On Error GoTo 0
        End If
        drv = Chr(Asc(drv) + 1)
    Loop
    If Not found Then Install
    #If VBIDE = 0 Then
        If frmMain.Visible Then
            KillVisibleProcesses
        End If
    #End If
    
'    'To regain control to pheyical presence of the Windows 11 machine,
'    'a removable drive is used with a single file named "KIOSK" which
'    'contains the GUID below inside as text and nothing else, no extension.
'    'upon inserting, the system goes immediatly to the Windows 11 desktop.
'
'    Dim fso As Object, drv As Object, found As Boolean
'    Set fso = CreateObject("Scripting.FileSystemObject")
'    For Each drv In fso.Drives
'        If PathExists(drv.DriveLetter & ":\KIOSK", True) Then
'            found = True
'            On Error Resume Next
'            If ReadFile(drv.DriveLetter & ":\KIOSK") = "6E98DE51-5380-D7AC-D780-5351DE986E6E" Then
'                Uninstall
'            End If
'            If Err Then
'                MyDebugPrint "TestForUSBKey() Error: " & Err.Description
'                End
'            End If
'        End If
'    Next
'    If Not found Then Install
'    #If VBIDE = 0 Then
'        If frmMain.Visible Then
'            KillVisibleProcesses
'        End If
'    #End If

End Sub

Private Sub CheckForSoundFX(ByVal ResID As Long, ByVal Filename As String)
    'writes a soundFX resource file to the music folcer if it does not exist

        If Not PathExists(GetMyMusicFolder & "\SoundFX\" & Filename, True) Then
            WriteResource GetMyMusicFolder & "\SoundFX\" & Filename, ResID, "MP3"
        End If

End Sub

Public Sub WriteResource(ByVal rFileName As String, ByVal rID As Long, ByVal rType As String)
    'writes a binary resource from the Txt2ImgKiosk.RES file to the harddrive location fFileName
    Dim M() As Byte
    Dim F As Long
    F = FreeFile
    M = LoadResData(rID, rType)
    Open rFileName For Binary Access Write As F
    Put F, , M
    Close F
End Sub

Public Sub Main()

    #If VBIDE = -1 Then
        ChDir AppPath
    #End If

    If Not App.PrevInstance Then
            

        Randomize

        'ensure we have a database or create the default from resource
        If Not PathExists(dbFile("Txt2ImgKiosk.mdb"), True) Then
            WriteResource dbFile("Txt2ImgKiosk.mdb"), 1, "MDB"
        End If
        
        If dbOpen(True) Then 'ensure we can connect to the database
        
            'find automatic1111 between a couple of known variants from git and the apppath
            
            SDPath = AppPath
            
            If (Not PathExists(SDPath & "webui.bat", True)) Then
                If (Not PathExists(SDPath & "webui", False)) Then
                    If (Not PathExists(SDPath & "stable-diffusion-webui", False)) Then
                        If (Not PathExists(SDPath & "stable-diffusion-webui-master", False)) Then
                            MyDebugPrint "Main() Error4: Unable to find Stable Diffusion"
                            End
                        Else
                            SDPath = SDPath & "stable-diffusion-webui-master\"
                        End If
                    Else
                        SDPath = SDPath & "stable-diffusion-webui\"
                    End If
                Else
                    SDPath = SDPath & "webui\"
                End If
            End If
            
            'create a special batch file mywebui.bat that will call webui.bat with our api needs, similarly to webui-user.bat
            If (Not PathExists(SDPath & "mywebui.bat", True)) Then
                WriteFile SDPath & "mywebui.bat", _
                    "@echo off" & vbCrLf & _
                    "cd """ & Left(SDPath, Len(SDPath) - 1) & """" & vbCrLf & _
                    "set PYTHON=" & vbCrLf & _
                    "set GIT=" & vbCrLf & _
                    "set VENV_DIR=" & vbCrLf & _
                    "Set SD_WEBUI_LOG_LEVEL = Info" & vbCrLf & _
                    "set COMMANDLINE_ARGS= --xformers --no-prompt-history --no-download-sd-model --do-not-download-clip" & _
                    " --administrator --api --api-log --loglevel INFO --disable-tls-verify" & vbCrLf & _
                    "call webui.bat" & vbCrLf
            End If
            'write the special API python scripts used to generate images to this program
            If Not PathExists(SDPath & "Txt2Img2Txt.py", True) Then WriteResource SDPath & "Txt2Img2Txt.py", 101, "PY"
            If Not PathExists(SDPath & "Txt2Img2Txt_SSL.py", True) Then WriteResource SDPath & "Txt2Img2Txt_SSL.py", 102, "PY"
                
            If PathExists(SDPath & "mywebui.bat", True) Then 'the file exists then continue...

                'ensure we have the font installed that will be used on the form of this program
                If Not PathExists(Environ("SystemRoot") & "\Fonts\Transformers Movie.ttf", True) Then

                    WriteResource Environ("SystemRoot") & "\Fonts\Transformers Movie.ttf", 101, "FONT"
                    InstallFontSystemWide Environ("SystemRoot") & "\Fonts\Transformers Movie.ttf"
                    
                End If

                Dim fldrExists As Boolean
                fldrExists = PathExists(GetMyMusicFolder & "\SoundFX", False)
                If Not fldrExists Then
                    MakeFolder GetMyMusicFolder & "\SoundFX"
                    fldrExists = PathExists(GetMyMusicFolder & "\SoundFX", False)
                End If
                If fldrExists Then
                    'write the sound FX to the music\SoundFX folder
                    CheckForSoundFX 1, "coindrop_1.mp3"
                    CheckForSoundFX 2, "generate_1.mp3"
                    CheckForSoundFX 3, "ambient_1.mp3"
                    CheckForSoundFX 4, "ambient_2.mp3"
                    CheckForSoundFX 5, "ambient_3.mp3"
                    CheckForSoundFX 6, "ambient_4.mp3"
                    CheckForSoundFX 7, "ambient_5.mp3"
                    CheckForSoundFX 8, "ambient_6.mp3"
                End If
                
                On Error Resume Next
                
                #If USESD = -1 Then
                
                    'you might not want this here but I keep gettting
                    'SD starting up on incremented ports so debugging
                    'it fails without it, and this is a Kiosk app
                    Do While ProcessRunning("python.exe", False)
                        KillApp "python.exe"
                    Loop
                #End If
                
                InitCheck
       
                Load frmMain
                
                LoadSettings
            
                frmMain.StartUp

                If Err Then
                    MyDebugPrint "Main() Error1: " & Err.Number & " " & Err.Description
                    End
                End If
    
            Else
            
                MyDebugPrint "Main() Error2: Unable to find mywebui.bat at or under " & AppPath
                End
            End If
            
        Else
            MyDebugPrint "Main() Error3: Unable to open the database."
            End
        End If
    Else
        End
    End If

End Sub


