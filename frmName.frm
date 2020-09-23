VERSION 5.00
Begin VB.Form frmName 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Set your name:"
   ClientHeight    =   795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   795
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   285
      Left            =   2535
      TabIndex        =   2
      Top             =   390
      Width           =   1020
   End
   Begin VB.TextBox txtMyName 
      Height          =   300
      Left            =   105
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "Player"
      Top             =   360
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "State your name:"
      Height          =   240
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   4005
   End
End
Attribute VB_Name = "frmName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    YourName = txtMyName
    frmMain.Enabled = True
    frmMain.lblYou.Caption = strConnect2Start
    frmMain.ChatAdd txtMyName, 3
    frmMain.Caption = "X O X O - ''" & txtMyName & "'' -- By Michael Vainshtein"
    
    If frmMain.Ws.State = 7 Then frmMain.WsSend "OppName=" & YourName
    
    Unload Me
End Sub

Private Sub Form_Load()
    Show
    Randomize
    If YourName = "" Then
        txtMyName = txtMyName & Trim(Str(Int(1 + Rnd * 50)))
    Else: txtMyName = YourName
    End If
    txtMyName.SetFocus
End Sub

Private Sub txtMyName_GotFocus()
txtMyName.SelStart = 0
txtMyName.SelLength = Len(txtMyName)
End Sub

Private Sub txtMyName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdDone_Click
End Sub
