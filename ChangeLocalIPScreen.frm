VERSION 5.00
Begin VB.Form ChangeLocalIPScreen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1815
   ClientLeft      =   2580
   ClientTop       =   2865
   ClientWidth     =   4530
   Icon            =   "ChangeLocalIPScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1815
   ScaleWidth      =   4530
   Begin VB.CommandButton Cancel 
      BackColor       =   &H80000009&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton Connect 
      BackColor       =   &H80000009&
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox IPTextField 
      Height          =   375
      Left            =   1538
      MaxLength       =   15
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"ChangeLocalIPScreen.frx":08CA
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ChangeLocalIPScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Connect_Click()
    If Parent.ChangeLocalIP(IPTextField.Text) Then
        Hide
        MsgBox "Local IP Changed Successfully To " + IPTextField.Text, vbOKOnly, "Connection Successful!"
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Caption = "Current Local IP: " + Parent.getLocalIP
End Sub
