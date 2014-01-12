VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form CreatePlaylistForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create PlayList"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7770
   Icon            =   "CreatePlaylistForm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CreatePlaylistForm.frx":08CA
   ScaleHeight     =   4155
   ScaleWidth      =   7770
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer AnimationTimer 
      Enabled         =   0   'False
      Interval        =   750
      Left            =   1080
      Top             =   120
   End
   Begin VB.DriveListBox DriveBox 
      Height          =   315
      Left            =   5255
      TabIndex        =   3
      Top             =   120
      Width           =   2415
   End
   Begin MSComDlg.CommonDialog PlayListDialog 
      Left            =   1440
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save PlayList"
   End
   Begin VB.FileListBox FileList 
      Height          =   2235
      Left            =   5255
      Pattern         =   "*.mp3"
      TabIndex        =   2
      Top             =   1800
      Width           =   2415
   End
   Begin VB.DirListBox DirList 
      Height          =   1215
      Left            =   5255
      TabIndex        =   1
      Top             =   480
      Width           =   2415
   End
   Begin VB.ListBox Display 
      Height          =   3960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label NewList 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3480
      TabIndex        =   7
      Top             =   1440
      Width           =   525
   End
   Begin VB.Label Edit 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4200
      TabIndex        =   6
      Top             =   1440
      Width           =   495
   End
   Begin VB.Image AnimationBox 
      Height          =   1065
      Left            =   3120
      Picture         =   "CreatePlaylistForm.frx":89ACC
      Top             =   120
      Width           =   2130
   End
   Begin VB.Image Remove 
      Height          =   345
      Left            =   3720
      Picture         =   "CreatePlaylistForm.frx":911C2
      Top             =   2640
      Width           =   780
   End
   Begin VB.Image Add 
      Height          =   345
      Left            =   3720
      Picture         =   "CreatePlaylistForm.frx":92008
      Top             =   1920
      Width           =   780
   End
   Begin VB.Label CloseButton 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4320
      TabIndex        =   5
      Top             =   3120
      Width           =   690
   End
   Begin VB.Label Save 
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3360
      TabIndex        =   4
      Top             =   3120
      Width           =   615
   End
End
Attribute VB_Name = "CreatePlaylistForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim List, PrevDrive, EditedFileName, EditedPath As String
Dim Saved, Editing, Edited As Boolean
Dim AnimationIndex As Integer
Dim Animation() As Picture
Dim ButtonPictures() As Picture
Dim SelectedButtonPictures() As Picture

Private Sub Add_Click()
    addFile FileList.path + "\" + FileList.FileName
    If FileList.ListIndex + 1 < FileList.ListCount Then
        FileList.ListIndex = FileList.ListIndex + 1
    End If
End Sub

Private Sub Add_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Add.Picture = SelectedButtonPictures(0)
End Sub

Private Sub Add_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Add.Picture = ButtonPictures(0)
End Sub

Private Sub AnimationBox_Click()
    If AnimationTimer.Enabled Then
        AnimationTimer.Enabled = False
    Else
        AnimationTimer.Enabled = True
    End If
End Sub

Private Sub AnimationTimer_Timer()
    AnimationBox.Picture = Animation(AnimationIndex)
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex > UBound(Animation) - 1 Then
        AnimationIndex = 0
    End If
End Sub

Private Sub CloseButton_Click()
    If Not Editing And Not List = "" And Not Saved Then
        CloseConfirm.Show vbModal
        If CloseConfirm.getResult Then
            Save_Click
        End If
    End If
    
    If Editing And Edited And Not Saved Then
        CloseConfirm.Show vbModal
        If CloseConfirm.getResult Then
            Save_Click
        End If
    End If
    
    List = ""
    For i = 0 To Display.ListCount - 1
        Display.RemoveItem (0)
    Next i
    Clear
    Hide
End Sub

Private Sub addFile(tempString As String)
    If Not Right(tempString, 1) = "\" And Not tempString = "" Then
        List = List + tempString + vbNewLine
        If Editing Then
            Edited = True
        End If
        displayList
        Display.ListIndex = Display.ListCount - 1
    End If
End Sub

Private Sub displayList()
    Dim broken, BrokenFilename As Variant
    Dim FileName As String
    broken = Split(List, vbNewLine)
    Display.Clear
    For i = 0 To (UBound(broken) - 1)
        FileName = broken(i)
        BrokenFilename = Split(FileName, "\")
        FileName = BrokenFilename(UBound(BrokenFilename))
        Display.AddItem (FileName)
    Next i
End Sub

Private Sub CloseButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.ForeColor = &HFFFF&
    CloseButton.BackColor = &H80000008
End Sub

Private Sub CloseButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CloseButton.ForeColor = &H80000008
    CloseButton.BackColor = &H80000009
End Sub

Private Sub DirList_Change()
    FileList.path = DirList.path
    If FileList.ListCount > 0 Then
        FileList.ListIndex = 0
    End If
End Sub

Private Sub DriveBox_Change()
    On Error GoTo DriveUnavailable
    DirList.path = DriveBox.drive
    PrevDrive = DriveBox.drive
    Exit Sub
DriveUnavailable:
    MsgBox DriveBox.drive + " Device Unavailable", vbOKOnly, "ERROR"
    DriveBox.drive = PrevDrive
End Sub

Private Sub Edit_Click()
    Dim tempFile As String
    Dim broken As Variant
    PlayListDialog.FileName = "*.rjk"
    PlayListDialog.Filter = ".rjk"
    PlayListDialog.DefaultExt = "rjk"
    PlayListDialog.InitDir = Parent.getPLPath
    PlayListDialog.DialogTitle = "Open PlayList"
    On Error GoTo cancelled
    PlayListDialog.ShowOpen
    
    If Not PlayListDialog.FileName = "" And Not PlayListDialog.FileName = "*.rjk" Then
        List = ""
        Open PlayListDialog.FileName For Input As #1
        On Error GoTo error
        Do
            Input #1, tempFile
            addFile tempFile
            Edited = False
        Loop While True
    End If
error:
        Close #1
        Debug.Print "list: " + List
        If Not List = "" Then
            EditedPath = ""
            EditedFileName = ""
            Editing = True
            broken = Split(PlayListDialog.FileName, "\")
            For i = 0 To UBound(broken) - 1
                If i = 0 Then
                    EditedPath = broken(0)
                Else
                    EditedPath = EditedPath + "\" + broken(i)
                End If
            Next i
            EditedFileName = broken(UBound(broken))
            Caption = "Editing: " + EditedFileName
            Debug.Print EditedPath
            Debug.Print EditedFileName
        Else
            MsgBox "Invalid File", vbOKOnly, "ERROR"
        End If
cancelled:
End Sub

Private Sub Edit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Edit.ForeColor = &HFFFF&
    Edit.BackColor = &H80000008
End Sub

Private Sub Edit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Edit.ForeColor = &H80000008
    Edit.BackColor = &H80000009
End Sub

Private Sub FileList_DblClick()
    addFile FileList.path + "\" + FileList.FileName
End Sub

Private Sub Form_Load()
    ReDim ButtonPictures(2) As Picture
    ReDim SelectedButtonPictures(2) As Picture
    ReDim Animation(4) As Picture
    AnimationIndex = 1
    EditedFileName = ""
    List = ""
    Saved = False
    Editing = False
    Edited = False
    PrevDrive = DriveBox.drive
    Load CloseConfirm
    On Error GoTo FileError
    Set Animation(0) = LoadPicture(Parent.getDataPath + "Animation0.bmp")
    Set Animation(1) = LoadPicture(Parent.getDataPath + "Animation1.bmp")
    Set Animation(2) = LoadPicture(Parent.getDataPath + "Animation2.bmp")
    Set Animation(3) = LoadPicture(Parent.getDataPath + "Animation3.bmp")
    Set ButtonPictures(0) = LoadPicture(Parent.getDataPath + "add.bmp")
    Set ButtonPictures(1) = LoadPicture(Parent.getDataPath + "remove.bmp")
    Set SelectedButtonPictures(0) = LoadPicture(Parent.getDataPath + "selectedadd.bmp")
    Set SelectedButtonPictures(1) = LoadPicture(Parent.getDataPath + "selectedremove.bmp")
    Exit Sub
FileError:
        Debug.Print "create"
        MsgBox "Unable To Locate Program Files." + vbNewLine + "Reinstallation may be necessary.", vbOKOnly, "ERROR"
        Unload Parent
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload CloseConfirm
End Sub

Private Sub NewList_Click()
    Clear
End Sub

Private Sub NewList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewList.ForeColor = &HFFFF&
    NewList.BackColor = &H80000008
End Sub

Private Sub NewList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    NewList.ForeColor = &H80000008
    NewList.BackColor = &H80000009
End Sub

Private Sub Remove_Click()
    Dim NewList As String
    Dim broken As Variant
    Dim Index As Integer
    
    Index = Display.ListIndex
    If Index = -1 Then
        Exit Sub
    End If

    broken = Split(List, vbNewLine)
    For i = 0 To UBound(broken) - 1
        If Not i = Index Then
            NewList = NewList + broken(i) + vbNewLine
        End If
    Next i
    List = NewList
    Display.RemoveItem (Index)
    If Editing Then
        Edited = True
    End If
    On Error GoTo indexerror
    Display.ListIndex = Index
    Exit Sub
indexerror:
    Display.ListIndex = Display.ListCount - 1
End Sub

Private Sub Remove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Remove.Picture = SelectedButtonPictures(1)
End Sub

Private Sub Remove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Remove.Picture = ButtonPictures(1)
End Sub

Private Sub Save_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Save.ForeColor = &HFFFF&
    Save.BackColor = &H80000008
End Sub

Private Sub Save_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Save.ForeColor = &H80000008
    Save.BackColor = &H80000009
End Sub

Private Sub Save_Click()
    If Editing And Not Edited Then
        Exit Sub
    End If
    
    If Editing Then
        PlayListDialog.InitDir = EditedPath
        PlayListDialog.FileName = EditedFileName
    Else
        PlayListDialog.InitDir = Parent.getPLPath
        PlayListDialog.FileName = "*.rjk"
    End If
    PlayListDialog.Filter = ".rjk"
    PlayListDialog.DefaultExt = "rjk"
    PlayListDialog.DialogTitle = "Save PlayList"

    On Error GoTo error
    If Not Display.ListCount = 0 Then
        PlayListDialog.ShowSave
        If Not PlayListDialog.FileName = "" And Not PlayListDialog.FileName = "*.rjk" Then
            SavePlayList PlayListDialog.FileName
        End If
    End If
error:
End Sub

Private Sub SavePlayList(FileName As String)
    Dim broken As Variant
    broken = Split(List, vbNewLine)
    On Error GoTo SaveError
    Debug.Print FileName
    Open FileName For Output As #1
        For i = 0 To UBound(broken) - 1
            Debug.Print broken(i)
            Write #1, broken(i)
        Next i
        Saved = True
    Close #1
    Clear
    Exit Sub
SaveError:
    Close #1
    MsgBox "Error Saving PlayList.", vbOKOnly, "ERROR"
End Sub

Private Sub Clear()
    Caption = "Create PlayList"
    List = ""
    Display.Clear
    Edited = False
    Editing = False
    Saved = False
End Sub
