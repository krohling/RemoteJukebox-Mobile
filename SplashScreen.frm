VERSION 5.00
Begin VB.Form SplashScreen 
   BackColor       =   &H80000009&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   1800
   ClientTop       =   2115
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   Picture         =   "SplashScreen.frx":0000
   ScaleHeight     =   2505
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   Begin VB.Timer SplashTimer 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   1680
      Top             =   480
   End
End
Attribute VB_Name = "SplashScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo loadError
    Load Parent
    SplashTimer.Enabled = True
    Exit Sub
loadError:
    Unload Me
End Sub

Private Sub SplashTimer_Timer()
    Hide
    Parent.Show
    Unload Me
End Sub
