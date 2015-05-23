VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   960
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock wsFTP 
      Left            =   1440
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8535
   End
   Begin MSWinsockLib.Winsock wsBNET 
      Left            =   480
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wsListen 
      Left            =   0
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Menu mnuLaunch 
      Caption         =   "Launch Menu"
      Begin VB.Menu mnuSCLaunch 
         Caption         =   "SC Beta"
      End
      Begin VB.Menu mnuBWLaunch 
         Caption         =   "BW Beta"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Select Server"
   End
   Begin VB.Menu mnuReadme 
      Caption         =   "~> README <~"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


' The StarCraft Beta Gateway, written by Filip Jaroš , aka l2k-Shadow

' Thanks TO:
' l2k-Spec-Ops_X for showing me the beta.
' Lead@USEast, for helping me with packet logging and setting up his old computer as a server
' so i could test by myself through RDC, lol :)
' Doral@USEast, setup a PvPGN server
' Testers (Entropy aka Physician, Lead, Doral, l2k-Minosha, others..)

Option Base 0
Option Explicit
Option Compare Text

Public Sub Form_Load()
    'load it
    Me.Show
    frmMain.wsListen.Tag = frmMain.wsListen.LocalIP ' that's your local ip
    BotVersion = Chr$(118) & Chr$(51) & Chr$(46) & Chr$(48) & Chr$(49) ' version
    BotCaption = Chr$(83) & Chr$(116) & Chr$(97) & Chr$(114) & Chr$(67) & Chr$(114) & Chr$(97) & Chr$(102) & Chr$(116) & Chr$(32) & Chr$(66) & Chr$(101) & Chr$(116) & Chr$(97) & Chr$(32) & Chr$(71) & Chr$(97) & Chr$(116) & Chr$(101) & Chr$(32) & BotVersion & Chr$(32) & Chr$(45) & Chr$(32) & Chr$(66) & Chr$(121) & Chr$(58) & Chr$(32) & Chr$(108) & Chr$(50) & Chr$(107) & Chr$(45) & Chr$(83) & Chr$(104) & Chr$(97) & Chr$(100) & Chr$(111) & Chr$(119) ' caption
    Me.Caption = BotCaption
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'resize
    Text1.Height = Me.ScaleHeight
    Text1.Width = Me.ScaleWidth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    Unload frmSetup
End Sub

Private Sub mnuConfig_Click()
    'config
    frmSetup.Show
End Sub

Private Sub mnuReadme_Click()
    ' readme
    ShellExecute Me.hWnd, "open", App.Path & "\readme.txt", vbNullString, Left$(App.Path, 3), 1
End Sub

Private Sub mnuBWLaunch_Click()
    'launchbeta
    If CheckStuff("BroodWarBeta") Then
        If FindWindow("SWarClass", vbNullString) <> 0 Then ' Is Starcraft already running?
            ' Yep
            MsgBox "StarCraft is already running on your computer!", vbExclamation
        Else
            Client = 1
            RewriteBattle Client
            ResetConnections
            ' Nope, we're clear to launch beta
            ShellExecute Me.hWnd, "open", InstallPath & "\StarCraft.exe", vbNullString, Left$(InstallPath, 3), 1
        End If
    End If
End Sub

Private Sub mnuSCLaunch_Click()
    'launch beta
    If CheckStuff("StarcraftBeta") Then
        If FindWindow("SWarClass", vbNullString) <> 0 Then ' Is Starcraft already running?
            ' Yep
            MsgBox "StarCraft is already running on your computer!", vbExclamation
        Else
            Client = 0
            RewriteBattle Client
            Call FileCopy(App.Path & "\files\scexe.beta", InstallPath & "\StarCraft.exe")
            ResetConnections
            ' Nope, we're clear to launch beta
            ShellExecute Me.hWnd, "open", InstallPath & "\StarCraft.exe", vbNullString, Left$(InstallPath, 3), 1
        End If
    End If
End Sub

Private Sub Timer1_Timer()
' the timer to spoof 0x00
Dim p As New clsPacketBuffer
    p.Send wsListen, &H0
    Set p = Nothing
    Timer1.Enabled = False
End Sub

Private Sub wsBNET_Close()
    'bnet closed
    AddChat "wsBNET_Close"
    ResetConnections
End Sub

Private Sub wsBNET_Connect()
    'and we're connected! send the protocol byte
    Call wsBNET.SendData(Chr$(&H1))
End Sub

Private Sub wsBNET_DataArrival(ByVal bytesTotal As Long)
' parsing data that we get from bnet
Static Buffer As String
Dim Data As String, Length As Long
    Call wsBNET.GetData(Data, vbString, bytesTotal)
    Buffer = Buffer & Data
    While Len(Buffer) >= 4
        Length = GetWORD(Mid$(Buffer, 3, 2))
        If Length > Len(Buffer) Then Exit Sub
        Data = Left$(Buffer, Length)
        Call ParseRecvData(Data)
        Buffer = Mid$(Buffer, Length + 1)
    Wend
End Sub

Private Sub wsBNET_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'ouch error
    AddChat "wsBNET Error: " & Number & " - " & Description
    ResetConnections
End Sub

Private Sub wsFTP_Close()
    'ftp closed
    AddChat "wsFTP_Close"
End Sub

Private Sub wsFTP_ConnectionRequest(ByVal requestID As Long)
    ' Beta needs something from FTP
    Call wsFTP.Close
    Call wsFTP.Accept(requestID)
End Sub

Private Sub wsFTP_DataArrival(ByVal bytesTotal As Long)
Dim Data As String
    ' parsing what beta wants from FTP, will be IX86ver3.mpq
    Call wsFTP.GetData(Data, vbString, bytesTotal)
    If Len(Data) > 1 Then Call SendFTPFile("IX86ver3.mpq")
End Sub

Private Sub wsFTP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' this can't be good btw
    AddChat "wsFTP Error: " & Number & " - " & Description
End Sub

Private Sub wsListen_Close()
    ' beta closed connection
    AddChat "wsListen_Close"
    ResetConnections
End Sub

Private Sub wsListen_ConnectionRequest(ByVal requestID As Long)
    'beta is trying to connect to bnet!
    Call wsBNET.Connect(ServerToIP(G("Server")), 6112) ' gate connects to bnet
    While wsBNET.State <> sckConnected
        DoEvents
    Wend
    Call wsListen.Close
    Call wsListen.Accept(requestID) ' accept
    wsFTP.LocalPort = 6112
    Call wsFTP.Listen ' listen for FTP requests
End Sub

Private Sub wsListen_DataArrival(ByVal bytesTotal As Long)
' parse data from beta
Static Buffer As String
Dim Data As String, Length As Long
    Call wsListen.GetData(Data, vbString, bytesTotal)
    If Not LDataArrived And Left$(Data, 1) = Chr$(&H1) Then
        Data = Mid$(Data, 2)
        LDataArrived = True
    End If
    Buffer = Buffer & Data
    While Len(Buffer) >= 4
        Length = GetWORD(Mid$(Buffer, 3, 2))
        If Length > Len(Buffer) Then Exit Sub
        Data = Left$(Buffer, Length)
        Call ParseSendData(Data)
        Buffer = Mid$(Buffer, Length + 1)
    Wend
End Sub

Private Sub wsListen_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'error ;/
    AddChat "wsListen Error: " & Number & " - " & Description
    ResetConnections
End Sub
