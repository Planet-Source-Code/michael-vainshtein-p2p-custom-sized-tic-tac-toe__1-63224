Attribute VB_Name = "mdlGlobal"
Option Explicit
Global Const Testing = False
Global Const RemotePort = 1991
Global IHost As Boolean
Global YourName, OppName As String 'Your\Opponent name
Global RemoteIP As String
Global Const strConnect2Start = "Press connect to start a game"


Public Sub GlobalInit()
    IHost = False
    YourName = ""
    OppName = "Opponenet"
End Sub


Public Sub CloseSock()
    If frmMain.Ws.State = 7 Then frmMain.WsSend "Exit=exit"
    frmMain.Ws.Close
End Sub


Public Sub Wait(s As Double, Optional DoEv As Boolean = True)
    Dim T
    T = Timer
    Do While s + T > Timer
        If DoEv = True Then DoEvents
    Loop
End Sub
