VERSION 5.00
Begin VB.Form SelectDirectoryScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select PlayList Directory"
   ClientHeight    =   2850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "SelectDirectoryScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   975
   End
   Begin VB.DirListBox DirectoryBox 
      Height          =   1665
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.DriveListBox DriveBox 
      Height          =   315
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Look in: "
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   160
      Width           =   615
   End
End
Attribute VB_Name = "SelectDirectoryScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PrevDrive As String

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub DriveBox_Change()
    On Error GoTo DriveError
    DirectoryBox.path = DriveBox.drive
    PrevDrive = DriveBox.drive
    Exit Sub
DriveError:
        MsgBox DriveBox.drive + " Device Unavailable", vbOKOnly, "ERROR"
        DriveBox.drive = PrevDrive
End Sub

Private Sub Form_Load()
    DisableX.DisableX Me.hwnd
    On Error GoTo assignmenterror
    DriveBox.drive = Parent.getPLDrive
    PrevDrive = Parent.getPLDrive
    DirectoryBox.path = Parent.getPLPath
    Exit Sub
assignmenterror:
    MsgBox "Unable to Locate PlayList Location: " + vbNewLine + Parent.getPLPath, vbOKOnly, "ERROR"
End Sub

Private Sub OK_Click()
    Parent.changePLPath DriveBox.drive, DirectoryBox.path
    Unload Me
End Sub
