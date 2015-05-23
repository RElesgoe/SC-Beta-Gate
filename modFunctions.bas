Attribute VB_Name = "modFunctions"

Option Base 0
Option Explicit
Option Compare Text

' good ole functions

Public Sub RewriteBattle(ByVal Client As Long)
' tricky!
Dim hFile As Long, s As String, ol As OVERLAPPED, t As String
    s = frmMain.wsListen.LocalIP
    hFile = CreateFile(InstallPath & "\battle.snp", &H40000000, &H2, 0&, &H3, &H0, &H0)
    If hFile Then
        If Client Then
            'bw beta
            ol.offset = &H21968
            WriteFile hFile, ByVal s, Len(s), 0, ol
            ol.offset = &H21968 + Len(s)
            t = String$(43 - Len(s), vbNullChar)
            WriteFile hFile, ByVal t, 43 - Len(s), 0, ol
            ol.offset = &H219A4
            WriteFile hFile, ByVal s, Len(s), 0, ol
            ol.offset = &H219A4 + Len(s)
            t = String$(35 - Len(s), vbNullChar)
            WriteFile hFile, ByVal t, 35 - Len(s), 0, ol
        Else
            'sc beta
            ol.offset = &H1EDF4
            WriteFile hFile, ByVal s, Len(s), 0, ol
            ol.offset = &H1EDF4 + Len(s)
            t = String$(39 - Len(s), vbNullChar)
            WriteFile hFile, ByVal t, 39 - Len(s), 0, ol
            ol.offset = &H1EE2C
            WriteFile hFile, ByVal s, Len(s), 0, ol
            ol.offset = &H1EE2C + Len(s)
            t = String$(87 - Len(s), vbNullChar)
            WriteFile hFile, ByVal t, 87 - Len(s), 0, ol
        End If
        CloseHandle hFile
    Else
        AddChat "Failed to open Battle.snp for writing!"
    End If
End Sub

Public Function CheckStuff(ByVal RegKey As String) As Boolean
' check if beta is installed
Dim hKey As Long
    If G("Server") = vbNullString Then
        frmMain.Hide
        frmSetup.Show
        Exit Function
    End If
    InstallPath = String$(256, vbNullChar)
    Call RegOpenKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Blizzard Entertainment\" & RegKey, hKey)
    Call RegQueryValueEx(hKey, "InstallPath", 0, 1, InstallPath, 256)
    Call RegCloseKey(hKey)
    InstallPath = TrimString(InstallPath)
    If InstallPath = vbNullString Then
        MsgBox RegKey & " is not installed!", vbExclamation
        Exit Function
    End If
    frmMain.Text1.Text = vbNullString ' clear chat
    CheckStuff = True ' good!
End Function

Public Sub AddChat(ByVal Text As Variant)
' add text to box
    With frmMain.Text1
        If Len(.Text) >= 5000 Then
            .SelStart = 0
            .SelLength = InStr(.Text, vbCrLf) + 1
            .SelText = vbNullString
        End If
        .SelStart = Len(.Text)
        .SelText = Text & vbCrLf
    End With
End Sub

Public Sub W(ByVal Key As String, ByVal sString As String) ' write to ini file
    Call WritePrivateProfileString("StarCraft Beta Gate", Key, sString, App.Path & "\Config.ini")
End Sub

Public Function G(ByVal Key As String) As String ' get from ini file
Dim sRet As String * 128
    Call GetPrivateProfileString("StarCraft Beta Gate", Key, vbNullString, sRet, 128, App.Path & "\Config.ini")
    G = TrimString(sRet)
End Function

Public Function TrimString(ByVal sString As String, Optional ByVal Delimiter As String = vbNullChar)
' just cut string off once it reaches delimiter
Dim i As Integer
    i = InStr(sString, Delimiter)
    If i = 0 Then
        TrimString = sString
    Else
        TrimString = Left$(sString, i - 1)
    End If
End Function

'DWORD/WORD functions

Public Function MakeDWORD(ByVal Data As Long) As String
Dim sRet As String * 4
    Call CopyMemory(ByVal sRet, Data, 4)
    MakeDWORD = sRet
End Function

Public Function GetDWORD(ByVal Data As String) As Long
    Call CopyMemory(GetDWORD, ByVal Data, 4)
End Function

Public Function MakeWORD(ByVal Data As Long) As String
Dim sRet As String * 2
    Call CopyMemory(ByVal sRet, Data, 2)
    MakeWORD = sRet
End Function

Public Function GetWORD(ByVal Data As String) As Long
    Call CopyMemory(GetWORD, ByVal Data, 2)
End Function

' yeah i know this function won't always work but it works for this case so i'm using it!

Public Function GetLastString(ByVal Data As String, Optional ByVal Count As Integer = 1) As String
Dim i As Integer, c As Integer
    For i = Len(Data) To 0 Step -1
        If i = 0 Then
            GetLastString = TrimString(Data)
            Exit Function
        Else
            If Mid$(Data, i, 1) = vbNullChar Then
                If c = Count Then
                    GetLastString = TrimString(Mid$(Data, i + 1))
                    Exit Function
                Else
                    c = c + 1
                End If
            End If
        End If
    Next i
End Function

Public Function FormStatString(ByVal Old As String) As String
' this function took a while to code, forms the old statstring that beta wants in S->C 0x09
Dim i As Integer, c As Integer, j As Integer, Name As String, Map As String, NewSS As String
Dim p As New clsPacketBuffer
    For i = 1 To Len(Old)
        If c = 11 Then
            Name = TrimString(Mid$(Old, i), Chr$(&HD))
            For j = Len(Name) To 25
                Name = Name & Chr$(&HFF)
            Next j
            Map = TrimString(Mid$(Old, i + Len(Name) + 1), Chr$(&HD))
            For j = Len(Map) To 56
                Map = Map & Chr$(&HFF)
            Next j
            Exit For
        End If
        If Mid$(Old, i, 1) = "," Then c = c + 1
    Next i
    p.InsertDWORD &HFFFFFF0A
    p.InsertDWORD &HFFFFFFFF
    p.InsertDWORD &HFFFFFFFF
    p.InsertDWORD &HFF80FF80
    p.InsertDWORD &H160401FF
    p.InsertDWORD &HFF01FF01
    p.InsertVOID Name
    p.InsertSTRING Map
    FormStatString = p.Buffer
    Set p = Nothing
End Function

Public Sub ResetConnections()
' reset winsocks for use!
Dim i As Integer
    Call frmMain.wsListen.Close
    Call frmMain.wsBNET.Close
    Call frmMain.wsFTP.Close
    frmMain.wsListen.LocalPort = 6112
    Call frmMain.wsListen.Listen
    LDataArrived = False
    MyUsername = vbNullString
    AddChat "Your local IP is: " & frmMain.wsListen.Tag
    AddChat "Listening for connections on port " & frmMain.wsListen.LocalPort & "..."
End Sub

Public Function IsNot(ByVal Data As Long) As Long
' just converting stuff
    Select Case Data
        Case &H0
            IsNot = &H1
        Case Else
            IsNot = &H0
    End Select
End Function

Public Function ServerToIP(ByVal Server As String) As String
' from config to ip
    Select Case Server
        Case "U.S. East"
            ServerToIP = "63.240.202.120"
        Case "U.S. West"
            ServerToIP = "63.241.83.110"
        Case "Europe"
            ServerToIP = "213.248.106.204"
        Case "Asia"
            ServerToIP = "211.233.0.54"
        Case Else
            ServerToIP = Server
    End Select
End Function

Public Function MakeIPDWORD() As String
' turn your ip into DWORD
Dim splt() As String, i As Integer
    splt = Split(frmMain.wsListen.Tag, ".")
    For i = 0 To UBound(splt)
        MakeIPDWORD = MakeIPDWORD & Chr$(CLng(splt(i)))
    Next i
End Function

' Debugging functions below

Public Function HexToStr(ByVal strText As String) As String
Dim i As Long, t As String
    strText = Replace(strText, " ", "")
    For i = 1 To Len(strText) Step 2
        HexToStr = HexToStr & Chr$(CLng("&H" & Mid$(strText, i, 2)))
    Next i
End Function

Public Function StrToHex(ByVal strText As String) As String
Dim i As Long, t As String
    If strText = vbNullString Then Exit Function
    For i = 1 To Len(strText)
        t = Hex(Asc(Mid$(strText, i, 1)))
        If Len(t) = 1 Then t = "0" & t
        StrToHex = StrToHex & t & Space(1)
    Next i
    StrToHex = Left$(StrToHex, Len(StrToHex) - 1)
End Function

Public Function OneHex(ByVal sID As Integer) As String
    OneHex = Hex(sID)
    If Len(OneHex) = 1 Then OneHex = "0" & OneHex
End Function
