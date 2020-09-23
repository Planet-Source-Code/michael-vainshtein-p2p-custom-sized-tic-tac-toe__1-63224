VERSION 5.00
Begin VB.Form frmJoin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Join game"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   2235
      TabIndex        =   3
      Top             =   855
      Width           =   1590
   End
   Begin VB.CommandButton cmdJoin 
      Caption         =   "Join"
      Height          =   300
      Left            =   105
      TabIndex        =   2
      Top             =   855
      Width           =   1590
   End
   Begin VB.TextBox txtHostIP 
      Height          =   300
      Left            =   75
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "frmJoin.frx":0000
      Top             =   480
      Width           =   3825
   End
   Begin VB.Label Label1 
      Caption         =   "Join this host:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   1
      Top             =   180
      Width           =   1665
   End
End
Attribute VB_Name = "frmJoin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    IHost = False
    frmMain.Enabled = True
    frmMain.Ws.Close
    Unload Me
End Sub

Private Sub cmdJoin_Click()
    
    RemoteIP = txtHostIP

    cmdJoin.Enabled = False
    cmdJoin.Caption = "Please wait..."
    txtHostIP = "Connecting to host: " & txtHostIP
    txtHostIP.Enabled = False
    
    CloseSock
    Do While frmMain.Ws.State <> 7
        CloseSock
        frmMain.Ws.Connect RemoteIP, RemotePort
        Wait (6)
    Loop
End Sub

Private Sub txtHostIP_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdJoin_Click
End Sub

Private Sub txtHostIP_GotFocus()
txtHostIP.SelStart = 0
txtHostIP.SelLength = Len(txtHostIP)
End Sub
