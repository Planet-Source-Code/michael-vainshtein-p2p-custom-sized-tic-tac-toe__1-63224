VERSION 5.00
Begin VB.Form frmWait 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listening to port......"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelHost 
      Caption         =   "Cancel hosting"
      Height          =   285
      Left            =   1677
      TabIndex        =   3
      Top             =   870
      Width           =   1270
   End
   Begin VB.TextBox txtIP 
      Alignment       =   2  'Center
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   900
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   435
      Width           =   2925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Your IP:"
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Please wait for an opponent. . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   45
      TabIndex        =   0
      Top             =   120
      Width           =   4305
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancelHost_Click()
    frmMain.Ws.Close
    frmMain.Enabled = True
    IHost = False
    Unload Me
End Sub

Private Sub Form_Load()
    Show
    txtIP = frmMain.Ws.LocalIP
    txtIP.Left = Width / 2 - txtIP.Width / 2
    Center txtIP, cmdCancelHost
End Sub




Public Sub Center(a As Control, b As Control)
b.Move a.Left + a.Width / 2 - b.Width / 2
End Sub
