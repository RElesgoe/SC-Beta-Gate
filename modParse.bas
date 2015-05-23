Attribute VB_Name = "modParse"

Option Base 0
Option Explicit
Option Compare Text

Public Sub ParseRecvData(ByVal Data As String)
Dim p As New clsPacketBuffer, t As String, i As Long
    'AddChat "Recv: wsBNET - " & StrToHex(Data)
    Select Case Asc(Mid$(Data, 2, 1))
        Case &H15
            AddChat "S->C: Ignored Packet 0x" & OneHex(Asc(Mid$(Data, 2, 1)))
        Case &H6
            Call Send0x06(frmMain.wsListen, Data)
        Case &H9
            ' join game packet
            Call Send0x09(frmMain.wsListen, Data, "S->C: Filtered Packet 0x09, Sending 0x09")
        Case &HF
            ' The famous chatting packet
            Call Send0x0F(frmMain.wsListen, Data, "S->C: Filtered Packet 0x0F, Sending 0x0F")
        Case &H1C
            ' Today's bnet runs on 0x1C, but our oldie needs 0x08
            p.InsertDWORD IsNot(GetDWORD(Mid$(Data, 5, 4)))
            p.Send frmMain.wsListen, &H8
            AddChat "S->C: Filtered Packet 0x1C, Sending 0x08"
            InGame = True ' we're in game
        Case &H3A
            ' Today's bnet runs on 0x3A, but our oldie needs 0x29
            p.InsertDWORD IsNot(GetDWORD(Mid$(Data, 5, 4)))
            p.Send frmMain.wsListen, &H29
            AddChat "S->C: Filtered Packet 0x3A, Sending 0x29"
        Case &H3D
            ' Today's bnet runs on 0x3D, but our oldie needs 0x2A
            p.InsertDWORD IsNot(GetDWORD(Mid$(Data, 5, 4)))
            p.Send frmMain.wsListen, &H2A
            AddChat "S->C: Filtered Packet 0x3D, Sending 0x2A"
        Case Else
            AddChat "S->C: 0x" & OneHex(Asc(Mid$(Data, 2, 1))) & " (" & StrToHex(Data) & ")"
            frmMain.wsListen.SendData Data
    End Select
    Set p = Nothing
End Sub

Public Sub ParseSendData(ByVal Data As String)
Dim p As New clsPacketBuffer, i As Integer, splt() As String, tStr As String
Dim GameName As String, GamePass As String, GameType As Long, PacketID As Integer
    If Len(Data) < 4 Then Exit Sub
    'AddChat "Recv: wsListen - " & StrToHex(Data)
    PacketID = Asc(Mid$(Data, 2, 1))
    Select Case PacketID
        Case &H2
            ' 0x02 means we left the game or disonnected from bnet
            frmMain.wsBNET.SendData Data
            AddChat "C->S: 0x02"
            If InGame Then
                InGame = False
            Else
                AddChat "Resetting connections..."
                ResetConnections
            End If
        Case &H12, &H15, &H2B, &H2C, &H2D
            ' defunct/useless packets which we have to or can ignore
            AddChat "C->S: Ignored Packet 0x" & OneHex(PacketID)
        Case &H6
            Call TSend0x06(frmMain.wsBNET, Data, "C->S: Filtered Packet 0x06, Sending 0x06")
        Case &H7
            ' this is the beta sending the information about its version to battle.net
            Call Send0x07(frmMain.wsBNET, Data, "C->S: Filtered Packet 0x07, Sending 0x07")
            frmMain.Timer1.Interval = 1000 ' again the stuck thing, let's spoof 0x00
            frmMain.Timer1.Enabled = True
        Case &H9
            ' i don't feel like making a new sub so here
            If Len(Data) = &H17 Then 'is the game retreiving the game list or trying to join game?
                ' retrieving game list
                p.InsertDWORD &H0
                p.InsertDWORD &H0
                p.Send frmMain.wsListen, &H9
                ' send "no games found" back. Why? beta crashes otherwise
                AddChat "C->S->C: Filtered Packet 0x09, Sending Back"
            Else
                ' getting one game info
                GameName = GetLastString(Data, 3) ' Game Name!
                GamePass = GetLastString(Data, 2) ' Game Pass!
                p.InsertDWORD &H0 ' product conditions, set to 0, ..
                p.InsertDWORD &H0 ' same thing
                p.InsertDWORD &H0 ' same thing
                p.InsertDWORD &H1 ' number of games to get info on
                p.InsertSTRING GameName ' name
                p.InsertSTRING GamePass ' pass
                p.InsertSTRING vbNullString ' statstring , we're retrieving it so this is null
                p.Send frmMain.wsBNET, &H9 ' send away
            End If
        Case &HC
            frmMain.wsBNET.SendData Data
            AddChat "C->S: Filtered Packet 0x0C, Sending 0x0C"
            SendEvent &H6, "- You are running SC Beta Gate " & BotVersion
            ' get the news
            'tStr = GetURL("http://www.energydl.com/shadow/BETA/News.txt") ' if you want
            'SendEvent &H6, tStr ' and send the news
        Case &H26
            ' get the profile, edit that one a little bit to suit today's protocol
            ' removes the name if you're retrieving your own
            If MyUsername <> vbNullString And InStr(Data, MyUsername) Then ' retrieving your own?
                'yep remove the name!
                p.InsertVOID Mid$(Data, 5, 12)
                p.InsertVOID Mid$(Data, Len(MyUsername) + 17)
                p.Send frmMain.wsBNET, &H26 ' and send
            Else
                'nope, you're fine
                frmMain.wsBNET.SendData Data
            End If
            AddChat "C->S: Filtered Packet 0x26, Sending 0x26"
        Case &H27
            ' writing your profile
            p.InsertVOID Mid$(Data, 5, 8)
            p.InsertVOID Mid$(Data, Len(MyUsername) + 13) 'remove name
            p.Send frmMain.wsBNET, &H27 ' and send
            AddChat "C->S: Filtered Packet 0x27, Sending 0x27"
        Case &H29
            MyUsername = TrimString(Mid$(Data, 33))
            frmMain.wsBNET.SendData Data
        Case &H30
            ' cd-key packet. since we are logging in with shareware, we don't send this
            p.InsertDWORD &H1
            p.InsertSTRING vbNullString
            p.Send frmMain.wsListen, &H30
        Case Else
            ' else, just forward those
            AddChat "C->S: 0x" & OneHex(PacketID) & " (" & StrToHex(Data) & ")"
            frmMain.wsBNET.SendData Data
    End Select
    Set p = Nothing
End Sub

Public Sub SendEvent(ByVal EventID As Long, ByVal Message As String, Optional ByVal Username As String)
' send event!
Dim p As New clsPacketBuffer, splt() As String, i As Integer
    splt = Split(Message, vbCrLf)
    For i = 0 To UBound(splt)
        p.InsertDWORD EventID ' Event ID
        p.InsertDWORD &H0 ' FLAGS!
        p.InsertDWORD &H0 ' PING!
        p.InsertDWORD &H0 ' This used to be the user's IP, it's null now of course
        p.InsertDWORD &H0 ' crap
        p.InsertDWORD &H0 ' crap
        p.InsertSTRING Username ' name
        p.InsertSTRING splt(i) ' and the message
        p.Send frmMain.wsListen, &HF ' send away
    Next i
    Set p = Nothing ' free the buffer!
End Sub

Public Sub SendFTPFile(ByVal File As String)
Dim p As New clsPacketBuffer, tPath As String, FF As Integer, fLen As Long
    'When beta requests a file (will only be IX86ver3.mpq, since we sent it that file in 0x06)
    tPath = App.Path & "\Files\" & File ' Path to file
    AddChat "DL: " & File
    fLen = FileLen(tPath) ' Length of file
    p.InsertDWORD Len(File) + 25 ' Length of FTP header (not including the file)
    p.InsertDWORD fLen ' there's our file length
    p.InsertDWORD &H0 ' banner stuff 0 works
    p.InsertDWORD &H0 ' more banner stuff 0 works
    p.InsertDWORD &H4341AC00 ' filetime
    p.InsertDWORD &H1C50B25 ' filetime
    p.InsertSTRING File 'file name
    AddChat "DLING: " & tPath & " " & fLen
    FF = FreeFile ' Free file
    Open tPath For Binary As #FF ' Open
        Do Until EOF(FF) ' Until there is no more in the file
            p.InsertVOID Input(fLen, FF) ' put into buffer
        Loop
    Close #FF ' Close!
    p.SendRaw frmMain.wsFTP ' send away!
    AddChat "C->S->C: Sending Back File " & File
End Sub

Private Sub Send0x06(ByVal WS As Winsock, ByVal s As String)
Dim p As New clsPacketBuffer
    p.InsertDWORD &H4341AC00 ' first part of filetime, lol.
    p.InsertDWORD &H1C50B25 ' second part of filetime
    p.InsertSTRING "IX86ver3.mpq" ' Yeah so the client thinks it will get the real file.
    p.InsertSTRING "A=125933019 B=665814511 C=736475113 4 A=A+S B=B^C C=C^A A=A^B" 'some bullshit hash
    p.Send WS, &H6 ' Good luck client, make sure you hash this without failure, it will matter much
    Set p = Nothing ' don't forget to free!
End Sub

Private Sub TSend0x06(ByVal WS As Winsock, ByVal IncomingData As String, ByVal s As String)
Dim p As New clsPacketBuffer
    p.InsertVOID "CAMPRHSS" ' Starcraft Shareware for Mac
    p.InsertDWORD &HA5 ' verbyte
    p.InsertDWORD &H0 ' no clue
    p.Send WS, &H6
    Set p = Nothing ' Free your packet buffer!
    AddChat s
End Sub

Private Sub Send0x07(ByVal WS As Winsock, ByVal IncomingData As String, ByVal s As String)
Dim p As New clsPacketBuffer, Version As Long, ExeInfo As String, Checksum As Long
    p.InsertVOID "CAMPRHSS" ' Starcraft Shareware for Mac
    p.InsertDWORD &HA5 ' verbyte
    p.InsertDWORD &H0 ' version of game
    p.InsertDWORD &H0 ' checksum of game files
    p.InsertSTRING vbNullString ' exe information
    p.Send WS, &H7
    Set p = Nothing
    AddChat s
End Sub

Private Sub Send0x09(ByVal WS As Winsock, ByVal IncomingData As String, ByVal s As String)
Dim p As New clsPacketBuffer
    ' Toughest packet
    If (Client = 0 And InStr(IncomingData, ",,")) Then ' Joining game or just game info?
        p.InsertVOID Mid$(IncomingData, 5, 4) ' Number of games
        p.InsertDWORD &H0 ' game type/parameter some crap
        p.InsertWORD &HFFFF ' no clue
        p.InsertWORD &H1 ' no clue
        p.InsertWORD &H2 ' Address family (always AF_INET)
        p.InsertVOID Mid$(IncomingData, 19, 6) ' Host's IP & Port
        p.InsertDWORD &H0 ' no clue
        p.InsertDWORD &H0 ' no clue
        p.InsertDWORD &H4 ' no clue
        p.InsertDWORD &H1 ' time elapsed since game was created let's make it 1 second :)
        p.InsertSTRING GetLastString(IncomingData, 3) ' Game Name!
        p.InsertSTRING GetLastString(IncomingData, 2) ' Game Pass!
        p.InsertVOID FormStatString(GetLastString(IncomingData)) ' Game Statstring, what a pain in the ass
        p.Send WS, &H9 ' Send away!
        InGame = True ' Yep, we're in
    Else
        WS.SendData IncomingData ' game info, leave as is
    End If
    Set p = Nothing ' FREE IT!
    AddChat s
End Sub

Private Sub Send0x0F(ByVal WS As Winsock, ByVal IncomingData As String, ByVal s As String)
Dim p As New clsPacketBuffer, t As String
    Select Case GetDWORD(Mid$(IncomingData, 5, 4)) ' What event?
        Case &H1, &H2, &H3, &H9 ' User, Join, Leave, Flags
            t = GetLastString(IncomingData) ' Message
            t = Replace(t, "PXES", "RATS") ' Client is on BW, Make him SC :)
            p.InsertVOID Mid$(IncomingData, 5, 24) ' First chunk of crap
            p.InsertSTRING TrimString(Mid$(IncomingData, 29)) ' Username
            If t <> vbNullString And InStr(t, "RATS") = 0 Then 'Is the user on SC?
                ' Nope, make him SC
                p.InsertSTRING "RATS 0 0 0 0 0 0 0 0 RATS" ' Insert SC statstring
            Else
                ' Yep he's on SC
                p.InsertSTRING t ' Leave message alone
            End If
            p.Send WS, &HF ' Send away!
        Case Else
            WS.SendData IncomingData ' Don't have to edit other events, send them as is
    End Select
    Set p = Nothing ' Free your packet buffer!
    AddChat s
End Sub
