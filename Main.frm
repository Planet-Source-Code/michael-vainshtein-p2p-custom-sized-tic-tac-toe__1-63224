VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X O X O X O X O  - By Michael Vainshtein"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrPing 
      Interval        =   4000
      Left            =   7425
      Top             =   630
   End
   Begin VB.PictureBox BlueO 
      Height          =   285
      Left            =   4680
      Picture         =   "Main.frx":4DF2
      ScaleHeight     =   225
      ScaleWidth      =   120
      TabIndex        =   19
      Top             =   4320
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox BlueX 
      Height          =   570
      Left            =   3255
      Picture         =   "Main.frx":5C46
      ScaleHeight     =   510
      ScaleWidth      =   495
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.PictureBox RedX 
      Height          =   495
      Left            =   2580
      Picture         =   "Main.frx":6A9A
      ScaleHeight     =   435
      ScaleWidth      =   390
      TabIndex        =   17
      Top             =   4185
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.PictureBox RedO 
      Height          =   345
      Left            =   1950
      Picture         =   "Main.frx":78EE
      ScaleHeight     =   285
      ScaleWidth      =   315
      TabIndex        =   16
      Top             =   4215
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox BG 
      Height          =   645
      Left            =   720
      Picture         =   "Main.frx":8742
      ScaleHeight     =   585
      ScaleWidth      =   615
      TabIndex        =   15
      Top             =   3960
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.TextBox txtChatSend 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   195
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   8085
      Width           =   6735
   End
   Begin VB.TextBox txtChat 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   195
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   7290
      Width           =   6990
   End
   Begin VB.Timer tmrCircles 
      Interval        =   1
      Left            =   1320
      Top             =   1500
   End
   Begin VB.Timer tmrStatus 
      Interval        =   100
      Left            =   7500
      Top             =   2340
   End
   Begin VB.Frame fraSep 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   -90
      Width           =   60
   End
   Begin VB.Frame fraSep 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   0
      Left            =   3990
      TabIndex        =   5
      Top             =   -90
      Width           =   60
   End
   Begin MSWinsockLib.Winsock Ws 
      Left            =   645
      Top             =   465
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Cell 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Index           =   0
      Left            =   195
      ScaleHeight     =   390
      ScaleWidth      =   390
      TabIndex        =   0
      Top             =   450
      Width           =   450
   End
   Begin VB.Label lblPing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   7185
      TabIndex        =   20
      Top             =   270
      Width           =   45
   End
   Begin VB.Shape Shape2 
      DrawMode        =   6  'Mask Pen Not
      Height          =   1590
      Left            =   7425
      Top             =   3750
      Width           =   450
   End
   Begin VB.Shape Shape1 
      DrawMode        =   6  'Mask Pen Not
      Height          =   2160
      Left            =   7425
      Top             =   3180
      Width           =   615
   End
   Begin VB.Label lblVote 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Click here to: "
      Height          =   705
      Index           =   1
      Left            =   7410
      TabIndex        =   14
      Top             =   3180
      Width           =   435
   End
   Begin VB.Label lblVote 
      BackStyle       =   0  'Transparent
      Caption         =   "V o t e  a t   P S C"
      Height          =   2130
      Index           =   0
      Left            =   7920
      TabIndex        =   13
      Top             =   3195
      Width           =   120
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   80
      Left            =   6945
      Shape           =   1  'Square
      Top             =   8025
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   79
      Left            =   8100
      Shape           =   1  'Square
      Top             =   3090
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   78
      Left            =   8205
      Shape           =   1  'Square
      Top             =   3090
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   77
      Left            =   7980
      Shape           =   1  'Square
      Top             =   2910
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   76
      Left            =   7725
      Shape           =   1  'Square
      Top             =   2910
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   75
      Left            =   7845
      Shape           =   1  'Square
      Top             =   2910
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   74
      Left            =   7530
      Shape           =   1  'Square
      Top             =   2910
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   73
      Left            =   7635
      Shape           =   1  'Square
      Top             =   2910
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   72
      Left            =   7410
      Shape           =   1  'Square
      Top             =   2910
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   47
      Left            =   7845
      Shape           =   1  'Square
      Top             =   1800
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   71
      Left            =   7215
      Shape           =   1  'Square
      Top             =   6405
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   70
      Left            =   6990
      Shape           =   1  'Square
      Top             =   -195
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   69
      Left            =   6990
      Shape           =   1  'Square
      Top             =   19
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   68
      Left            =   6990
      Shape           =   1  'Square
      Top             =   233
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   64
      Left            =   6990
      Shape           =   1  'Square
      Top             =   447
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   63
      Left            =   6990
      Shape           =   1  'Square
      Top             =   660
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   67
      Left            =   7185
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   66
      Left            =   7410
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   65
      Left            =   7305
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   62
      Left            =   7740
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   61
      Left            =   7965
      Shape           =   1  'Square
      Top             =   720
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   60
      Left            =   7860
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   59
      Left            =   7620
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   58
      Left            =   7500
      Shape           =   1  'Square
      Top             =   705
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   57
      Left            =   7965
      Shape           =   1  'Square
      Top             =   975
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   56
      Left            =   7965
      Shape           =   1  'Square
      Top             =   855
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   55
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1230
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   54
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1080
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   53
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1455
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   52
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1335
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   51
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1680
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   50
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1560
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   49
      Left            =   7965
      Shape           =   1  'Square
      Top             =   1800
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   48
      Left            =   7425
      Shape           =   1  'Square
      Top             =   1800
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   46
      Left            =   7635
      Shape           =   1  'Square
      Top             =   1800
      Width           =   150
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   45
      Left            =   7005
      Shape           =   1  'Square
      Top             =   1830
      Width           =   375
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   44
      Left            =   7245
      Shape           =   1  'Square
      Top             =   7875
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   39
      Left            =   7215
      Shape           =   1  'Square
      Top             =   7710
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   38
      Left            =   7215
      Shape           =   1  'Square
      Top             =   7530
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   37
      Left            =   7215
      Shape           =   1  'Square
      Top             =   7305
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   36
      Left            =   7980
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   35
      Left            =   8100
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   29
      Left            =   7875
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   34
      Left            =   7170
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   32
      Left            =   7290
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   40
      Left            =   7530
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   43
      Left            =   7650
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   28
      Left            =   7410
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   42
      Left            =   0
      Shape           =   1  'Square
      Top             =   0
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   41
      Left            =   7770
      Shape           =   1  'Square
      Top             =   8025
      Width           =   60
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   33
      Left            =   7815
      Shape           =   1  'Square
      Top             =   7890
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   31
      Left            =   7815
      Shape           =   1  'Square
      Top             =   7440
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   30
      Left            =   7815
      Shape           =   1  'Square
      Top             =   7665
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   27
      Left            =   7815
      Shape           =   1  'Square
      Top             =   6540
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   26
      Left            =   7815
      Shape           =   1  'Square
      Top             =   6765
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   25
      Left            =   7815
      Shape           =   1  'Square
      Top             =   6990
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   24
      Left            =   7815
      Shape           =   1  'Square
      Top             =   7215
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   23
      Left            =   7815
      Shape           =   1  'Square
      Top             =   5655
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   22
      Left            =   7815
      Shape           =   1  'Square
      Top             =   5880
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   21
      Left            =   7815
      Shape           =   1  'Square
      Top             =   6105
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   20
      Left            =   7815
      Shape           =   1  'Square
      Top             =   6330
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   19
      Left            =   7815
      Shape           =   1  'Square
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   18
      Left            =   7815
      Shape           =   1  'Square
      Top             =   5445
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   17
      Left            =   7410
      Shape           =   1  'Square
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   16
      Left            =   7612
      Shape           =   1  'Square
      Top             =   5235
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   15
      Left            =   7035
      Shape           =   1  'Square
      Top             =   6405
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   14
      Left            =   7215
      Shape           =   1  'Square
      Top             =   6630
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   13
      Left            =   7215
      Shape           =   1  'Square
      Top             =   6855
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   12
      Left            =   7215
      Shape           =   1  'Square
      Top             =   7080
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   11
      Left            =   7035
      Shape           =   1  'Square
      Top             =   6195
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   10
      Left            =   7035
      Shape           =   1  'Square
      Top             =   5970
      Width           =   150
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   9
      Left            =   7035
      Shape           =   1  'Square
      Top             =   5745
      Width           =   150
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   8
      Left            =   7005
      Shape           =   1  'Square
      Top             =   5175
      Width           =   375
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   7
      Left            =   7005
      Shape           =   1  'Square
      Top             =   4752
      Width           =   375
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   6
      Left            =   7005
      Shape           =   1  'Square
      Top             =   4330
      Width           =   375
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   4
      Left            =   7005
      Shape           =   1  'Square
      Top             =   3908
      Width           =   375
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   3
      Left            =   7005
      Shape           =   1  'Square
      Top             =   3486
      Width           =   375
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   2
      Left            =   7005
      Shape           =   1  'Square
      Top             =   3064
      Width           =   375
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   1
      Left            =   7005
      Shape           =   1  'Square
      Top             =   2642
      Width           =   375
   End
   Begin VB.Shape s 
      BorderWidth     =   2
      Height          =   375
      Index           =   5
      Left            =   7035
      Shape           =   1  'Square
      Top             =   5520
      Width           =   150
   End
   Begin VB.Label lblM 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   15
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   6000
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "MishaSoft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   6930
      TabIndex        =   9
      Top             =   8235
      Width           =   1095
   End
   Begin VB.Shape s 
      Height          =   375
      Index           =   0
      Left            =   7005
      Shape           =   1  'Square
      Top             =   2220
      Width           =   375
   End
   Begin VB.Label lblTotalWins 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4995
      TabIndex        =   8
      Top             =   45
      Width           =   90
   End
   Begin VB.Label lblMode 
      BackStyle       =   0  'Transparent
      Caption         =   "Winsock Disconnected"
      ForeColor       =   &H0000FFFF&
      Height          =   405
      Left            =   7005
      TabIndex        =   6
      Top             =   1425
      Width           =   1005
   End
   Begin VB.Shape shpMode 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   6990
      Shape           =   3  'Circle
      Top             =   1065
      Width           =   345
   End
   Begin VB.Label lblYou 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   90
   End
   Begin VB.Label lblWin 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7110
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblTurn 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4125
      TabIndex        =   2
      Top             =   45
      Width           =   900
   End
   Begin VB.Label lblCoords 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7125
      TabIndex        =   1
      Top             =   2000
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Menu mnuConn 
      Caption         =   "Co&nnect"
      Begin VB.Menu mnuHost 
         Caption         =   "&Host a Game..."
      End
      Begin VB.Menu mnuJoin 
         Caption         =   "&Join a Game..."
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuName 
         Caption         =   "Change &Name..."
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuGraphicsOn 
         Caption         =   "Additional &Graphics"
      End
      Begin VB.Menu mnuShowLog 
         Caption         =   "Show &Log"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type Coords
    X As Integer
    Y As Integer
End Type

Const N = 15
Const WinCond = 4

Dim Cl, Darker 'Circles control vars
Dim You
Dim Turn
Dim Matrix(N - 1, N - 1) As String
Dim NL
Dim TotalWins(1) As Integer
Dim Ping

Private Sub Cell_Click(Index As Integer)
    If Testing = True Then You = Turn

    If You = Turn And Matrix(Cell2Coords(Index).X, Cell2Coords(Index).Y) = "" Then
        SetMark Index
        WsSend "OpWent=" & Index
        If IHost And CheckWin <> "False" Then OnWin
    End If
    txtChatSend.SetFocus
End Sub

Private Sub Cell_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cell_Click (Index)
End Sub

Private Sub Cell_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblCoords = "Index: " & Index & NL & "(" & Cell2Coords(Index).X & ", " & Cell2Coords(Index).Y & ")" & NL
    
    Dim i
    For i = 0 To s.Count - 1
        Darker(i) = 2
    Next
End Sub


Private Sub Form_Load()
    Dim i, j
    
    Show
    NL = vbNewLine
    Turn = "x"
    
    Randomize

    frmData.Visible = False
    frmData.Move Me.Left - frmData.Width, Me.Top
    
    mnuName_Click
    mnuGraphicsOn_Click
    
    InitCircles
    LoadCells
    For i = 0 To N - 1
        For j = 0 To N - 1
            Matrix(i, j) = ""
        Next
    Next
    
    UpdateMode
End Sub

Private Function Cell2Coords(ByVal a As Integer) As Coords
    Cell2Coords.X = a Mod 15
    Cell2Coords.Y = (a \ N)
End Function

Private Function Coords2Cell(X As Integer, Y As Integer) As Integer
    Coords2Cell = Y * 15 + X
End Function


Public Sub SetMark(Index As Integer)
    Dim s
    
    'Alternative code for text 'X's and 'O's (instead of pix)
    'Uncomment this and comment the next block
'    If Turn = You Then
'        Cell(Index).ForeColor = vbBlue
'    Else: Cell(Index).ForeColor = vbRed
'    End If
'
    'Cell(Index).FontSize = 28
    'Cell(Index).CurrentX = Cell(Index).ScaleWidth / 2 - (Cell(Index).TextWidth(Turn) * 0.5)
    'Cell(Index).CurrentY = Cell(Index).ScaleHeight / 2 - (Cell(Index).TextHeight(Turn) * 0.55)
    'Cell(Index).Print Turn
    
    If Turn = You Then
        If You = "x" Then
            Cell(Index).PaintPicture BlueX.Picture, 0, 0, Cell(Index).Width - 15, Cell(Index).Height - 15
        Else: Cell(Index).PaintPicture BlueO.Picture, 0, 0, Cell(Index).Width - 15, Cell(Index).Height - 15
        End If
    Else
        If You = "x" Then
             Cell(Index).PaintPicture RedO.Picture, 0, 0, Cell(Index).Width - 15, Cell(Index).Height - 15
        Else: Cell(Index).PaintPicture RedX.Picture, 0, 0, Cell(Index).Width - 15, Cell(Index).Height - 15
        End If
    End If
    
    '|-----------------------------------------------
    Matrix(Cell2Coords(Index).X, Cell2Coords(Index).Y) = Turn
    SwapTurn
End Sub

Public Sub SwapTurn()
    If Turn = "x" Then
        Turn = "o"
    ElseIf Turn = "o" Then Turn = "x"
    End If
    
UpdateLabels
End Sub


Public Function CheckWin(Optional Flag As Integer = 1) As String
On Error Resume Next
    Dim i, j, K As Integer
    Dim tmp As Boolean
    Dim Who As String
    CheckWin = "False"
    Who = ""
    
    For i = 0 To N - 1
        For j = 0 To N - 1
            If Matrix(i, j) <> "" And CheckWin = "False" Then
                Who = Matrix(i, j)
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i - K, j) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 1: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i + K, j) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 2: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i, j - K) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 3: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i, j + K) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 4: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i - K, j - K) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 5: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i + K, j + K) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 6: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i - K, j + K) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 7: Exit Function
                tmp = True
                For K = 1 To WinCond
                    If Matrix(i, j) <> Matrix(i + K, j - K) Then tmp = False: Exit For
                Next
                If tmp = True Then CheckWin = Who
                If tmp = True And Flag = 1 Then HighlightWin i, j, 8: Exit Function
                tmp = True
            End If
        Next
    Next
End Function

Public Sub LoadCells()
' DONT add "DoEvents" here.
' Adding it will cause bugs
    Dim i, j As Integer
    Cell(0).BackColor = vbWhite
    For i = 1 To (N * N - 1)
        Load Cell(i)
        With Cell(i)
            .Left = Cell(0).Left + (i Mod N) * Cell(i).Width
            .Top = Cell(0).Top + (i \ N) * Cell(0).Height
            .Cls
'Uncommecnt those 2 lines to get test values
'            Cell(i).Print Cell2Coords(i).x
'            Cell(i).Print Cell2Coords(i).y
        End With
    Next
    For i = 1 To (N * N - 1) / 2
        Cell(i).Visible = True
        Cell(N * N - i).Visible = True
        Wait 0.0035
    Next
End Sub

Public Sub UnloadObjects()
    Dim i
    For i = 1 To (N * N - 1)
        Unload Cell(i)
    Next
    
    
    LoadCells
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i
    For i = 0 To s.Count - 1
        Darker(i) = 2
    Next
End Sub

Private Sub Form_Terminate()
    mnuExit_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mnuExit_Click
    End
End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblVote_Click(Index As Integer)
    Dim Q
    Q = Shell("explorer ""http://k.1asphost.com/mishasoft/TTTredirect.asp""", vbNormalFocus)
    MsgBox "An explorer window has opened in the background. " & NL & "Vote and leave a comment. Thank you."
End Sub

Private Sub mnuConn_Click()
If lblYou = strConnect2Start Then lblYou = ""
End Sub

Private Sub mnuExit_Click()
    CloseSock
    End
End Sub

Private Sub mnuGraphicsOn_Click()
    Graphics mnuGraphicsOn.Checked
    mnuGraphicsOn.Checked = Not mnuGraphicsOn.Checked
End Sub

Private Sub mnuHost_Click()
    If Ws.State = 7 Then
        If vbYes = MsgBox("This will close active connection. " & NL & "Proceed?", "Connecting") Then Exit Sub
    End If
    
    IHost = True
    CloseSock
    Ws.LocalPort = RemotePort
    Ws.Listen
    frmWait.Show 0, Me
    Me.Enabled = False
End Sub

Private Sub mnuJoin_Click()
    If Ws.State = 7 Then
        If vbYes = MsgBox("This will close active connection. " & NL & "Proceed?", "Connecting") Then Exit Sub
    End If
    
    CloseSock
    IHost = False
    Me.Enabled = False
    frmJoin.Show
End Sub

Private Sub mnuName_Click()
    frmMain.Enabled = False
    frmName.Show
End Sub


Private Sub mnuOptions_Click()
    mnuGraphicsOn = True
    If Ws.State = 7 Then mnuGraphicsOn = False
End Sub

Private Sub mnuShowLog_Click()
    frmData.Visible = Not frmData.Visible
    mnuShowLog.Checked = Not mnuShowLog.Checked
End Sub

Private Sub tmrPing_Timer()
    WsSend "Ping=hi"
    Ping = Timer
End Sub

Private Sub tmrStatus_Timer()
    UpdateMode
End Sub

Private Sub txtChat_Change()
    Dim i
    For i = 0 To s.Count - 1
        Darker(i) = 2
    Next
End Sub

Private Sub txtChatSend_Change()
    If Len(txtChatSend.Text) > Len(vbNewLine) And txtChatSend <> "" Then
        If Ws.State = 7 And Right(txtChatSend, Len(vbNewLine)) = vbNewLine Then
            ChatAdd Left(txtChatSend, Len(txtChatSend.Text) - Len(vbNewLine)), 1
            WsSend "Chat=" & Left(txtChatSend, Len(txtChatSend.Text) - Len(vbNewLine))
            txtChatSend = ""
        End If
    End If
    Dim i
    For i = 0 To s.Count - 1
        Darker(i) = 2
    Next
End Sub

Private Sub Ws_Connect()
    Dim i, j
    If Not IHost Then
        frmJoin.Hide
        Me.Enabled = True
        WsSend "OppName=" & YourName
        ChatAdd "Connected to host: " & Ws.RemoteHostIP
        
        For i = 0 To N - 1
            For j = 0 To N - 1
                Matrix(i, j) = ""
            Next
        Next
        Turn = "x"
        UnloadObjects
    End If
End Sub

Private Sub Ws_ConnectionRequest(ByVal requestID As Long)
    Dim i, j
    Dim s As String
    If IHost Then
        Ws.Close
        Ws.Accept requestID
        
        frmWait.Hide
        Me.Enabled = True

        s = "ClientSide=o"
        You = "x"
        i = Int(Rnd * 2)
        If i = 2 Then You = "o": s = "ClientSide=x"
        WsSend s
        WsSend "OppName=" & YourName
        
        ChatAdd Ws.RemoteHostIP & " connected."
        UpdateLabels
        
        Turn = "x"
        
        For i = 0 To N - 1
            For j = 0 To N - 1
                Matrix(i, j) = ""
            Next
        Next
        
        UnloadObjects
    End If
End Sub

Private Sub Ws_DataArrival(ByVal bytesTotal As Long)
    Dim Data, tmp
    Ws.GetData Data, vbString
    Dim a, X, b
    X = Data
    frmData.txtData = frmData.txtData & X & NL
    Do While X <> ""
        a = Split(X, "|")
        X = Right(X, Len(X) - Len(a(0)))
        X = Replace(X, "|", "", , 1)
        b = Split(a(0), "=")
        Analize_Command b(0), b(1)
    Loop
    UpdateMode
End Sub

Public Sub Analize_Command(CmdName, CmdParams)
    Select Case LCase(CmdName)
        Case "clientside"
            You = CmdParams
            UpdateLabels
        Case "opwent"
            SetMark (CmdParams)
            If IHost And CheckWin <> "False" Then OnWin
        Case "chat"
            ChatAdd CmdParams, 2
        Case "oppname"
            OppName = CmdParams
            ChatAdd OppName, 4
        Case "win"
            CheckWin
'If you experiance problems due to slow connections try uncommenting this line:
'            Wait (1)
            OnWin
        Case "ping"
            If CmdParams = "hi" Then
                WsSend "ping=bye"
            Else
                Ping = Ping - Timer
                lblPing = "Ping: " & Ping / 2
            End If
        Case "exit"
            MsgBox "Opponent closed the connection.", vbExclamation, "Game over"
    End Select
    UpdateLabels
End Sub

Public Sub ChatAdd(ByVal s As String, Optional Flag As Integer = 0)
    Const SysMsg = "/= "
    txtChat.BackColor = RGB(255, 200, 200)
    If txtChat = "" Then NL = ""
    Select Case Flag
        Case 0: txtChat = txtChat & NL & SysMsg & s & StrReverse(SysMsg)
        Case 1: txtChat = txtChat & NL & YourName & ": " & s
        Case 2: txtChat = txtChat & NL & OppName & ": " & s
        Case 3: txtChat = txtChat & NL & SysMsg & "You've set your name to: ''" & s & "''" & StrReverse(SysMsg)
        Case 4: txtChat = txtChat & NL & SysMsg & "Opponent has set his name to: ''" & s & "''" & StrReverse(SysMsg)
    End Select
    txtChat.SelStart = Len(txtChat)
    NL = vbNewLine
    Wait (0.1)
    txtChat.BackColor = vbWhite
    Wait (0.1)
    txtChat.BackColor = RGB(255, 200, 200)
    Wait (0.1)
    txtChat.BackColor = vbWhite
End Sub

Public Sub WsSend(ByVal s As String)
    If Ws.State = 7 Then
        Ws.SendData s & "|"
        UpdateMode True
    ' Just a notification message which i found to be very annoying while testing
    ' uncomment if you with
    'Else: MsgBox "You are not connect to an opponent.", vbExclamation, "Cannot preform task"
    End If
End Sub

Public Sub UpdateMode(Optional Sending As Boolean = False)
    Dim s As String
    If Sending = True Then
        s = "Sending"
        If Not shpMode.BackColor = vbYellow Then shpMode.BackColor = vbYellow
        Exit Sub
    End If
    
    Select Case Ws.State
    Case 2
        s = "Listening"
        If Not shpMode.BackColor = vbYellow Then shpMode.BackColor = vbYellow
    Case 6
        s = "clsed by peer"
        If Not shpMode.BackColor = RGB(255, 100, 100) Then shpMode.BackColor = RGB(255, 100, 100)
    Case 7
        s = "Connected"
        If Not shpMode.BackColor = vbGreen Then shpMode.BackColor = vbGreen
    Case Else
        s = "Disconnected"
        If Not shpMode.BackColor = vbRed Then shpMode.BackColor = vbRed
    End Select
    
    s = "Peer state:" & NL & s
    If Not s = lblMode Then lblMode = s
    'Adds the winsock's state to the label.
    'Used for debugging
    'lblMode = "Winsock (" & Ws.State & ")" & NL & lblMode
End Sub

Private Sub Ws_SendComplete()
    UpdateMode
End Sub

Public Sub HighlightWin(ByVal i As Integer, ByVal j As Integer, Flag As Integer)
    Dim K
    Select Case Flag
    Case 1
         For K = 0 To WinCond
            Highlight i - K, j
         Next
    Case 2
         For K = 0 To WinCond
             Highlight i + K, j
         Next
    Case 3
         For K = 0 To WinCond
             Highlight i, j - K
         Next
    Case 4
         For K = 0 To WinCond
             Highlight i, j + K
         Next
    Case 5
         For K = 0 To WinCond
             Highlight i - K, j - K
         Next
    Case 6
         For K = 0 To WinCond
             Highlight i + K, j + K
         Next
    Case 7
         For K = 0 To WinCond
             Highlight i - K, j + K
         Next
    Case 8
         For K = 0 To WinCond
             Highlight i + K, j - K
         Next
    End Select
End Sub

Public Sub Highlight(ByVal X As Integer, ByVal Y As Integer)
    Cell(Coords2Cell(X, Y)).BackColor = RGB(194, 235, 233)
    Cell(Coords2Cell(X, Y)).FontSize = 20
    Cell(Coords2Cell(X, Y)).Cls
    Cell(Coords2Cell(X, Y)).CurrentX = Cell(Coords2Cell(X, Y)).ScaleWidth / 2 - (Cell(Coords2Cell(X, Y)).TextWidth(CheckWin(2)) * 0.5)
    Cell(Coords2Cell(X, Y)).CurrentY = Cell(Coords2Cell(X, Y)).ScaleHeight / 2 - (Cell(Coords2Cell(X, Y)).TextHeight(CheckWin(2)) * 0.55)
    Cell(Coords2Cell(X, Y)).Print CheckWin(2)
End Sub



Public Sub UpdateLabels()
    lblYou = "You are '" & UCase(You) & "'"
    
    lblTurn = UCase(Turn) & " Go!"
    If You = Turn Then lblTurn.ForeColor = vbBlue
    If You <> Turn Then lblTurn.ForeColor = vbRed
    
    lblTotalWins = "Total wins: [" & TotalWins(0) & " : " & TotalWins(1) & "]"
    
    fraSep(0).Left = lblYou.Left + lblYou.Width + 60
    lblTurn.Left = fraSep(0).Left + 120
    fraSep(1).Left = lblTurn.Left + lblTurn.Width
    lblTotalWins.Left = fraSep(1).Left + 120
End Sub

Public Sub OnWin()
    Dim Winner, i, j, s
    Winner = CheckWin(2)
    If Winner = You Then
        lblYou = "You've won!"
        TotalWins(0) = TotalWins(0) + 1
        ChatAdd "You've Won!"
    End If
    If Winner <> You And Winner <> "False" Then
        lblYou = "You've lost!"
        ChatAdd "You've Lost!"
        TotalWins(1) = TotalWins(1) + 1
    End If
    UpdateLabels

    
    NL = vbNewLine
    Turn = "x"


    For i = 0 To N - 1
        For j = 0 To N - 1
            Matrix(i, j) = ""
        Next
    Next
    
    If IHost Then
        WsSend "Win=" & Winner
        
        s = "ClientSide=o"
        You = "x"
        i = Int(Rnd * 2)
        If i = 2 Then You = "o": s = "ClientSide=x"
        WsSend s
    End If
    
    UpdateLabels
    UpdateMode False
    
    UnloadObjects
End Sub





Private Sub tmrCircles_Timer()
    Dim i
    For i = 0 To s.Count - 1
        s(i).BorderColor = RGB(Cl(i), Cl(i), 250)
        If Darker(i) = 1 Then Cl(i) = Cl(i) - 150
        If Cl(i) < 80 Then
            Darker(i) = 0
            Cl(i) = 80
        End If
        If Darker(i) = 2 Then Cl(i) = Cl(i) + 5
        If Cl(i) > 200 Then
            Darker(i) = 0
            Cl(i) = 200
        End If
    Next
End Sub

Public Sub InitCircles()
    ReDim Darker(s.Count)
    ReDim Cl(s.Count)
    Dim i
    
    
    lblM(0) = ""
    For i = 0 To s.Count - 1
        Cl(i) = 239
        s(i).BorderColor = RGB(Cl(i), Cl(i), Cl(i))
        s(i).Shape = s(0).Shape
        s(i).BorderColor = s(0).BorderColor
        s(i).BorderStyle = s(0).BorderStyle
        s(i).BorderWidth = s(0).BorderWidth
        's(I).Move s(I).Left, s(I).Top, s(0).Width, s(0).Height
        Darker(i) = 0
        If i > 0 Then Load lblM(i)
        lblM(i).Visible = True
        lblM(i).Move s(i).Left, s(i).Top, s(i).Width, s(i).Height
    Next
End Sub


Private Sub lblM_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i
    For i = 0 To s.Count - 1
        If Index <> i Then Darker(i) = 2
    Next
    If Cl(Index) <> 80 Then Darker(Index) = 1
End Sub


Public Sub Graphics(Off As Boolean)
On Error GoTo ERR
    Dim i
    If Off Then
        Picture = LoadPicture("")
        lblMode.ForeColor = vbBlack
        For i = 0 To s.Count - 1
            s(i).Visible = False
        Next
    Else
        Picture = BG.Picture
        lblMode.ForeColor = &H8000000F
        For i = 0 To s.Count - 1
            s(i).Visible = True
        Next
    End If
Exit Sub
ERR: MsgBox "Error occurred." & NL & "Make sure bg.bmp is in the same folder as this EXE."
End Sub


Public Sub Vote4Me()
    SetMark (0): SetMark (15): SetMark (30)
    SetMark (45): SetMark (61): SetMark (47): SetMark (32): SetMark (17): SetMark (2): SetMark (14): SetMark (12): SetMark (13): SetMark (27): SetMark (42): SetMark (43): SetMark (44): SetMark (57)
    SetMark (72): SetMark (73): SetMark (74): SetMark (8): SetMark (9): SetMark (10): SetMark (24): SetMark (39): SetMark (54): SetMark (69): SetMark (5): SetMark (6): SetMark (22): SetMark (37)
    SetMark (19): SetMark (34): SetMark (49): SetMark (65): SetMark (66): SetMark (52): SetMark (90): SetMark (105): SetMark (120): SetMark (121): SetMark (92): SetMark (107): SetMark (122): SetMark (137): SetMark (152): SetMark (95): SetMark (110): SetMark (125): SetMark (140): SetMark (155):
    SetMark (111): SetMark (127): SetMark (113): SetMark (99): SetMark (114): SetMark (129): SetMark (144): SetMark (159): SetMark (102): SetMark (103): SetMark (104): SetMark (132): SetMark (133): SetMark (134): SetMark (162): SetMark (163): SetMark (164): SetMark (101): SetMark (116): SetMark (131): SetMark (146): SetMark (161)
    Turn = "x"
End Sub
