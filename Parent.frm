VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Parent 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remote JukeBox"
   ClientHeight    =   3030
   ClientLeft      =   3165
   ClientTop       =   1920
   ClientWidth     =   3960
   Icon            =   "Parent.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Parent.frx":08CA
   ScaleHeight     =   3030
   ScaleWidth      =   3960
   Begin VB.Timer AnimationTimer 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   1560
      Top             =   3120
   End
   Begin VB.FileListBox PlayListInfo 
      Height          =   285
      Left            =   2400
      Pattern         =   "*.rjk"
      TabIndex        =   2
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin MSWinsockLib.Winsock SendSock 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin MSWinsockLib.Winsock Sock 
      Left            =   2160
      Top             =   3120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.ListBox Display 
      BackColor       =   &H80000009&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog LoadList 
      Left            =   1560
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Load PlayList"
   End
   Begin VB.Label DirectoryButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Directory"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   6
      ToolTipText     =   "Select PlayList Directory"
      Top             =   840
      Width           =   780
   End
   Begin VB.Label OpenIPWindowButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "IP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Change Local IP Address"
      Top             =   600
      Width           =   195
   End
   Begin VB.Line Line4 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   3960
      Y1              =   5
      Y2              =   5
   End
   Begin VB.Line Line3 
      BorderColor     =   &H8000000C&
      X1              =   5
      X2              =   5
      Y1              =   0
      Y2              =   4140
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000009&
      X1              =   3940
      X2              =   3940
      Y1              =   0
      Y2              =   4140
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   0
      X2              =   3960
      Y1              =   4120
      Y2              =   4120
   End
   Begin VB.Label CreatePlayList 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Create"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Create Playlist"
      Top             =   1320
      Width           =   570
   End
   Begin VB.Image Back 
      Height          =   345
      Left            =   870
      Picture         =   "Parent.frx":89ACC
      ToolTipText     =   "Back"
      Top             =   600
      Width           =   780
   End
   Begin VB.Image StopButton 
      Height          =   345
      Left            =   1590
      Picture         =   "Parent.frx":8A912
      ToolTipText     =   "Stop"
      Top             =   840
      Width           =   780
   End
   Begin VB.Image Forward 
      Height          =   345
      Left            =   2310
      Picture         =   "Parent.frx":8B758
      ToolTipText     =   "Forward"
      Top             =   600
      Width           =   780
   End
   Begin VB.Image Play 
      Height          =   345
      Left            =   1590
      Picture         =   "Parent.frx":8C59E
      ToolTipText     =   "Play"
      Top             =   360
      Width           =   780
   End
   Begin VB.Label loadplaylist 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Load"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Load Playlist"
      Top             =   1080
      Width           =   435
   End
   Begin VB.Image AnimationBox 
      Height          =   375
      Left            =   3480
      Picture         =   "Parent.frx":8D3E4
      ToolTipText     =   "Block/Unblock Connections"
      Top             =   120
      Width           =   375
   End
   Begin MediaPlayerCtl.MediaPlayer Playa 
      CausesValidation=   0   'False
      Height          =   495
      Left            =   240
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   0   'False
      AllowScan       =   0   'False
      AllowChangeDisplaySize=   0   'False
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   0   'False
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   0   'False
      EnableContextMenu=   0   'False
      EnablePositionControls=   0   'False
      EnableFullScreenControls=   0   'False
      EnableTracker   =   0   'False
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   0   'False
      ShowAudioControls=   0   'False
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   0   'False
      ShowStatusBar   =   0   'False
      ShowTracker     =   0   'False
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
End
Attribute VB_Name = "Parent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim List As Variant
Dim Current, AnimationMarker As Integer
Dim Loaded, StopBool, ClientConnected, ConnectionOpen As Boolean
Dim PLDrive, PLPath, PLName, ClientAddress, LocalAddress As String
Dim ButtonPictures() As Picture
Dim ClickedButtonPictures() As Picture
Dim Animation() As Picture
Dim ToggleConnectionPicture As Picture
Dim AnimationPath() As String
Dim PlayListArray() As String
Dim DataPath As String


Private Sub AnimationBox_Click()
    ToggleConnection
End Sub

Private Sub AnimationTimer_Timer()
    AnimationBox.Picture = Animation(AnimationMarker)
    AnimationMarker = AnimationMarker + 1
    If AnimationMarker > UBound(Animation) Then
        AnimationMarker = 0
    End If
End Sub

Private Sub Back_Click()
    If Current - 1 < 0 Or Not Loaded Then
        Exit Sub
    End If
    Current = Current - 1
    Display.ListIndex = Current
    If Not StopBool Then
        Play_Click
    End If
End Sub

Private Sub Back_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Back.Picture = ClickedButtonPictures(0)
End Sub

Private Sub Back_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Back.Picture = ButtonPictures(0)
End Sub



Private Sub CreatePlayList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CreatePlayList.ForeColor = &HFFFF&
    CreatePlayList.BackColor = &H80000008
End Sub

Private Sub CreatePlayList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CreatePlayList.ForeColor = &H80000008
    CreatePlayList.BackColor = &H80000009
End Sub

Private Sub DirectoryButton_Click()
    SelectDirectoryScreen.Show vbModal
End Sub

Private Sub DirectoryButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DirectoryButton.ForeColor = &HFFFF&
    DirectoryButton.BackColor = &H80000008
End Sub

Private Sub DirectoryButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DirectoryButton.ForeColor = &H80000008
    DirectoryButton.BackColor = &H80000009
End Sub

Private Sub loadplaylist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    loadplaylist.ForeColor = &HFFFF&
    loadplaylist.BackColor = &H80000008
End Sub

Private Sub loadplaylist_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    loadplaylist.ForeColor = &H80000008
    loadplaylist.BackColor = &H80000009
End Sub



Private Sub OpenIPWindowButton_Click()
    ChangeLocalIPScreen.Show vbModal
End Sub

Private Sub OpenIPWindowButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OpenIPWindowButton.ForeColor = &HFFFF&
    OpenIPWindowButton.BackColor = &H80000008
End Sub

Private Sub openIPWindowbutton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OpenIPWindowButton.ForeColor = &H80000008
    OpenIPWindowButton.BackColor = &H80000009
End Sub

Private Sub ToggleConnection()
    On Error GoTo ConnectionError
    If ConnectionOpen Then
        ConnectionOpen = False
        DisconnectClient
        Sock.Close
        AnimationBox.Picture = ToggleConnectionPicture
    ElseIf Not ConnectionOpen Then
        Sock.Bind 7777, LocalAddress
        ConnectionOpen = True
        AnimationBox.Picture = Animation(0)
    End If
    Exit Sub
ConnectionError:
    MsgBox "Error Opening Connection on " + Sock.LocalIP + " Port 7777", vbOKOnly, "ERROR"
End Sub

Public Function ChangeLocalIP(NewLocalIP As String) As Boolean
    On Error GoTo ConnectionError
    ConnectionOpen = False
    DisconnectClient
    Sock.Close
    AnimationBox.Picture = ToggleConnectionPicture
    LocalAddress = NewLocalIP
    Sock.Bind 7777, NewLocalIP
    ConnectionOpen = True
    AnimationBox.Picture = Animation(0)
    ChangeLocalIP = True
    Exit Function
ConnectionError:
    MsgBox "Error Opening Connection on " + NewLocalIP + " Port 7777", vbOKOnly, "ERROR"
    ChangeLocalIP = False
End Function


Private Sub CreatePlaylist_Click()
    CreatePlaylistForm.Show vbModal
End Sub

Private Sub Display_DblClick()
    PlayTrack (Display.ListIndex)
End Sub

Private Sub PlayTrack(Track As Integer)
    Current = Track
    On Error GoTo invalidtrack
    PlayFile (List(Track))
invalidtrack:
End Sub

Private Sub Form_Load()
    ReDim AnimationPath(3) As String
    ReDim Animation(3) As Picture
    ReDim ButtonPictures(3) As Picture
    ReDim ClickedButtonPictures(3) As Picture

    On Error GoTo FileError
    DataPath = PlayListInfo.path + "\" + "DAT\"
    openPLPath
    changePLPath PLDrive, PLPath
    LocalAddress = ""
    ConnectionOpen = False
    BackSelectedBool = False
    PlaySelectedBool = False
    ForwardSelectedBool = False
    StopSelectedBool = False
    ClientConnected = False
    StopBool = True
    AnimationMarker = 0
    Current = 0
    
    On Error GoTo ExitingProgram
    Load CreatePlaylistForm
    
    On Error GoTo FileError
    AnimationPath(0) = DataPath + "Phone1.bmp"
    AnimationPath(1) = DataPath + "Phone2.bmp"
    AnimationPath(2) = DataPath + "Phone3.bmp"
    AnimationPath(3) = DataPath + "Phone4.bmp"
    
    Set ToggleConnectionPicture = LoadPicture(DataPath + "blockphone.bmp")
    
    Set ButtonPictures(0) = LoadPicture(DataPath + "back.bmp")
    Set ButtonPictures(1) = LoadPicture(DataPath + "play.bmp")
    Set ButtonPictures(2) = LoadPicture(DataPath + "forward.bmp")
    Set ButtonPictures(3) = LoadPicture(DataPath + "stop.bmp")
    
    Set ClickedButtonPictures(0) = LoadPicture(DataPath + "highlightedback.bmp")
    Set ClickedButtonPictures(1) = LoadPicture(DataPath + "highlightedplay.bmp")
    Set ClickedButtonPictures(2) = LoadPicture(DataPath + "highlightedforward.bmp")
    Set ClickedButtonPictures(3) = LoadPicture(DataPath + "highlightedstop.bmp")
    
    Set Animation(0) = LoadPicture(AnimationPath(0))
    Set Animation(1) = LoadPicture(AnimationPath(1))
    Set Animation(2) = LoadPicture(AnimationPath(2))
    Set Animation(3) = LoadPicture(AnimationPath(3))
    
    On Error GoTo ConnectionError
    ClientAddress = ""
    Sock.Bind 7777
    LocalAddress = Sock.LocalIP
    ConnectionOpen = True
    Debug.Print Sock.LocalIP
    Debug.Print "localaddress: " + LocalAddress
    Exit Sub

ConnectionError:
    LocalAddress = Sock.LocalIP
    MsgBox "Error Opening Connection on " + Sock.LocalIP + " Port 7777", vbOKOnly, "ERROR"
    AnimationBox.Picture = ToggleConnectionPicture
    Exit Sub
FileError:
    Debug.Print "parent"
    MsgBox "Unable To Locate Program Files." + vbNewLine + "Reinstallation may be necessary.", vbOKOnly, "ERROR"
ExitingProgram:
    MsgBox "Exiting Program!", vbOKOnly, "ERROR"
    Unload Me
End Sub

Private Sub openPLPath()
    On Error GoTo notsaved
    Open DataPath + "rjbox.dat" For Input As #1
        Input #1, PLDrive
        Input #1, PLPath
    Close #1
    Exit Sub
notsaved:
    PLDrive = Left(PlayListInfo.path, 2)
    PLPath = PlayListInfo.path + "\" + "PlayList\"
End Sub

Private Sub DisplayError()
    Debug.Print "error"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ClientConnected Then
        sendData ClientAddress, "exiting" + vbNewLine
    End If
    Sock.Close
    SendSock.Close
    Unload CreatePlaylistForm
    Unload SelectDirectoryScreen
End Sub

Private Sub Forward_Click()
    If Not Loaded Then
        Exit Sub
    End If
    If Current + 1 > UBound(List) - 1 Then
        Current = -1
        Forward_Click
        Exit Sub
    End If
    Current = Current + 1
    Display.ListIndex = Current
    If Not StopBool Then
        Play_Click
    End If
End Sub

Private Sub Forward_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Forward.Picture = ClickedButtonPictures(2)
End Sub

Private Sub Forward_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Forward.Picture = ButtonPictures(2)
End Sub

Private Sub LoadPlaylist_Click()
    Dim SplitFile As Variant
    LoadList.FileName = "*.rjk"
    LoadList.Filter = ".rjk"
    LoadList.DefaultExt = "rjk"
    LoadList.ShowOpen
    LoadList.InitDir = PLPath
    If Not LoadList.FileName = "*.rjk" And Not LoadList.FileName = "" And Not Right(LoadList.FileName, 1) = "\" Then
        LoadPlayListInfo LoadList.FileName, False
        If ClientConnected And Loaded Then
            SendPLConfirm (ClientAddress)
            SendSongList (ClientAddress)
        End If
    End If
End Sub

Private Sub LoadPlayListInfo(ByVal FileName As String, ByVal RemoteRequest As Boolean)
        Dim BrokenName As Variant
        Debug.Print FileName
        FileList = ""
        BrokenName = Split(FileName, "\")
        
        On Error GoTo OpenError
        Debug.Print FileName
        Open FileName For Input As #1
        Do
            Input #1, FileName
            If Not FileName = "" Then
                FileList = FileList + FileName + vbNewLine
            End If
        Loop While True
    Exit Sub
OpenError:
    Debug.Print FileList
    If FileList = "" And Not RemoteRequest Then
        Close #1
        Loaded = False
        MsgBox "Invalid File", vbOKOnly, "Invalid File"
    ElseIf FileList = "" And RemoteRequest Then
        Close #1
        SendError "Error Loading PlayList: " + BrokenName(UBound(BrokenName))
        Loaded = False
    ElseIf Not FileList = "" Then
        Close #1
        PLName = BrokenName(UBound(BrokenName))
        List = Split(FileList, vbNewLine)
        If Not StopBool Then
            Current = -1
            StopBool = True
        ElseIf StopBool Then
            Current = 0
        End If
        Loaded = True
        Displayfiles
    End If
End Sub

Private Sub Play_Click()
    Current = Display.ListIndex
    If Current < 0 Then
        Current = 0
    End If
    If Loaded Then
        PlayFile (List(Current))
    End If
End Sub

Private Sub PlayFile(File As String)
    Playa.Open File
    Display.ListIndex = Current
    StopBool = False
    If ClientConnected Then
        NotifyCurrentSong (ClientAddress)
    End If
End Sub

Private Sub Play_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Play.Picture = ClickedButtonPictures(1)
End Sub

Private Sub Play_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Play.Picture = ButtonPictures(1)
End Sub

Private Sub Playa_PlayStateChange(ByVal OldState As Long, ByVal NewState As Long)
    If NewState = 0 And Not StopBool Then
        Forward_Click
    End If
End Sub

Private Sub SendSock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "ERROR", vbOKOnly, Description
    If ConnectionOpen Then
        ToggleConnection
    Else
        DisconnectClient
    End If
End Sub

Private Sub Sock_DataArrival(ByVal bytesTotal As Long)
    Dim RemoteCommand As String
    Dim SplitCommand As Variant
    Sock.GetData RemoteCommand, vbString
    Debug.Print RemoteCommand
    SplitCommand = Split(RemoteCommand, Right(RemoteCommand, 1))
    If connected And Not ClientAddress = Sock.RemoteHostIP Then
        Exit Sub
    End If
    
    ClientAddress = Sock.RemoteHostIP
    If Not ClientConnected Then
        ClientConnected = True
        AnimationTimer.Enabled = True
    End If
    
    If SplitCommand(0) = "play" Then
        If Not Loaded Or Not StopBool Then
            Exit Sub
        End If
        Play_Click
    ElseIf SplitCommand(0) = "stop" Then
        If Not Loaded Or StopBool Then
            Exit Sub
        End If
        Stopbutton_Click
    ElseIf SplitCommand(0) = "forward" Then
        If Not Loaded Then
            Exit Sub
        End If
        Forward_Click
    ElseIf SplitCommand(0) = "back" Then
        If Not Loaded Then
            Exit Sub
        End If
        Back_Click
    ElseIf SplitCommand(0) = "playlist" Then
        SendPlayListInfo (Sock.RemoteHostIP)
        Debug.Print StopBool
        If Not StopBool Then
            Debug.Print "inside"
            NotifyCurrentSong (Sock.RemoteHostIP)
        End If
        If Loaded Then
            SendPLConfirm (Sock.RemoteHostIP)
            SendSongList (Sock.RemoteHostIP)
        End If
    ElseIf SplitCommand(0) = "load" Then
        On Error GoTo bottom
        Debug.Print PlayListArray(CInt(SplitCommand(1)))
        LoadPlayListInfo PlayListArray(CInt(SplitCommand(1))), True
        If Loaded Then
            SendPLConfirm (Sock.RemoteHostIP)
            SendSongList (Sock.RemoteHostIP)
        End If
    ElseIf SplitCommand(0) = "playtrack" Then
        PlayTrack (CInt(SplitCommand(1)))
    ElseIf SplitCommand(0) = "exiting" Then
        DisconnectClient
        Debug.Print "client disconnected"
    End If
bottom:
End Sub

Private Sub DisconnectClient()
    ClientConnected = False
    AnimationTimer.Enabled = False
    AnimationMarker = 0
    AnimationBox.Picture = Animation(0)
End Sub

Private Sub SendPLConfirm(RequestAddress As String)
    Dim data As String
    If RequestAddress = "" Then
        Exit Sub
    End If
    data = "confirmplaylist" + vbNewLine + PLName + vbNewLine + vbNewLine
    sendData RequestAddress, data
End Sub

Private Sub SendError(Message As String)
    sendData ClientAddress, "error" + vbNewLine + Message + vbNewLine + vbNewLine
End Sub

Private Sub NotifyCurrentSong(RequestAddress As String)
    Dim data, FileName As String
    Dim BrokenName As Variant
    If RequestAddress = "" Then
        Exit Sub
    End If
    BrokenName = Split(List(Current), "\")
    FileName = BrokenName(UBound(BrokenName))
    If FileName = "" Then
        Exit Sub
    End If
    data = "songtitle" + vbNewLine + FileName + vbNewLine + vbNewLine
    sendData RequestAddress, data
End Sub

Private Sub SendSongList(RequestAddress As String)
    Dim data, FileName
    If Not Loaded Then
        Exit Sub
    End If
    If Not UBound(List) > 0 Then
        Exit Sub
    End If
    data = "songlist" + vbNewLine
    For i = 0 To UBound(List) - 1
        FileName = Split(List(i), "\")
        data = data + FileName(UBound(FileName)) + vbNewLine
    Next i
    data = data + vbNewLine
    Debug.Print data
    sendData RequestAddress, data
End Sub

Private Sub SendPlayListInfo(RequestAddress As String)
    Dim data As String
    Dim X As Integer
    data = "playlist" + vbNewLine
    PlayListInfo.Refresh
    updatePlayListArray
    X = PlayListInfo.ListCount
    Debug.Print "*******"
    Debug.Print X
    For i = 0 To X
        data = data + PlayListInfo.List(i) + vbNewLine
    Next i
    Debug.Print data
    sendData RequestAddress, data
End Sub

Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Winsock ERROR", vbOKOnly, Description
    If ConnectionOpen Then
        ToggleConnection
    Else
        DisconnectClient
    End If
End Sub

Private Sub Stopbutton_Click()
    StopBool = True
    Playa.Stop
    sendData ClientAddress, "stopped" + vbNewLine + vbNewLine
End Sub

Private Sub Displayfiles()
    Dim BrokenFilename As Variant
    Dim FileName As String
    Display.Clear
    On Error GoTo error
    For i = 0 To UBound(List) - 1
        BrokenFilename = Split(List(i), "\")
        Display.AddItem (BrokenFilename(UBound(BrokenFilename)))
    Next i
error:
End Sub

Private Sub sendData(ByVal TargetAddress As String, ByVal data As String)
    Dim error As Integer
    error = 0
    If Not ClientConnected Or Not ConnectionOpen Then
        Exit Sub
    End If
    Debug.Print "data sent"
    SendSock.RemoteHost = TargetAddress
    SendSock.RemotePort = 6666
    On Error GoTo retry
    SendSock.sendData data
    Exit Sub
retry:
    Debug.Print "sock error"
    If error < 2 Then
        error = error + 1
        SendSock.sendData data
    End If
End Sub

Private Sub StopButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StopButton.Picture = ClickedButtonPictures(3)
End Sub

Private Sub StopButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StopButton.Picture = ButtonPictures(3)
End Sub

Public Function getLocalIP() As String
    getLocalIP = LocalAddress
End Function

Public Function getDataPath() As String
    getDataPath = DataPath
End Function

Public Function getPLPath() As String
    getPLPath = PLPath
End Function

Public Function getPLDrive() As String
    getPLDrive = PLDrive
End Function

Public Sub changePLPath(ByVal drive As String, ByVal path As String)
On Error GoTo PLError
    LoadList.InitDir = path
    PlayListInfo.path = path
    PLPath = path
    PLDrive = drive
    updatePlayListArray
    SavePLPath drive, path
    If ClientConnected Then
        SendPlayListInfo (ClientAddress)
    End If
    Exit Sub
PLError:
    MsgBox "Error updating path: " + path, vbOKOnly, "ERROR"
End Sub

Private Sub updatePlayListArray()
    ReDim PlayListArray(PlayListInfo.ListCount)
    For i = 0 To PlayListInfo.ListCount
        PlayListArray(i) = PlayListInfo.path + "\" + PlayListInfo.List(i)
    Next i
End Sub

Private Sub SavePLPath(ByVal drive As String, ByVal path As String)
    On Error GoTo errorsaving
    Open DataPath + "rjbox.dat" For Output As #1
        Write #1, drive
        Write #1, path
    Close #1
    Exit Sub
errorsaving:
    MsgBox "Error Saving PlayList Directory Information.", vbOKOnly, "ERROR"
End Sub

