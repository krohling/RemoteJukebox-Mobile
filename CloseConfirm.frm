VERSION 5.00
Begin VB.Form CloseConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save PlayList?"
   ClientHeight    =   1290
   ClientLeft      =   2685
   ClientTop       =   2970
   ClientWidth     =   3210
   Icon            =   "CloseConfirm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   3210
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NoCommand 
      Cancel          =   -1  'True
      Caption         =   "No"
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton YesCommand 
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Would You Like To Save Your PlayList Before Continuing?"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "CloseConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Result As Boolean

Public Function getResult() As Boolean
    getResult = Result
End Function

Private Sub Form_Load()
    DisableX.DisableX Me.hwnd
End Sub

Private Sub NoCommand_Click()
    Result = False
    Hide
End Sub

Private Sub YesCommand_Click()
    Result = True
    Hide
End Sub
