VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacketBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

' my little packet buffer

Option Base 0
Option Explicit
Option Compare Text

Public Buffer As String
    
'CopyMemory Functions

Private Function MakeMemory(ByVal Data As Long, ByVal Length As Long) As String
Dim strOut As String
    strOut = String$(Length, vbNullChar)
    CopyMemory ByVal strOut, Data, Length
    MakeMemory = strOut
End Function

Public Sub InsertDWORD(ByVal Data As Long)
    Buffer = Buffer & MakeMemory(Data, 4)
End Sub

Public Sub InsertWORD(ByVal Data As Long)
    Buffer = Buffer & MakeMemory(Data, 2)
End Sub

Public Sub InsertSTRING(ByVal Data As String)
    Buffer = Buffer & Data & vbNullChar
End Sub

Public Sub InsertVOID(ByVal Data As String)
    Buffer = Buffer & Data
End Sub

Public Sub Send(ByVal WS As Winsock, ByVal PacketID As Long)
    'AddChat "Send: " & WS.Name & " - " & Hex(PacketID) & " - " & StrToHex(Buffer)
    If WS.State = sckConnected Then Call WS.SendData(Chr(&HFF) & Chr(PacketID) & MakeMemory(Len(Buffer) + 4, 2) & Buffer)
    Buffer = vbNullString
End Sub

Public Sub SendRaw(ByVal WS As Winsock)
    If WS.State = sckConnected Then Call WS.SendData(Buffer)
    Buffer = vbNullString
End Sub
