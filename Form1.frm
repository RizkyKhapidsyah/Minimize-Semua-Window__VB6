VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Minimize Semua Window"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call keybd_event(VK_LWIN, 0, 0, 0)
  Call keybd_event(&H4D, 0, 0, 0)
  Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

