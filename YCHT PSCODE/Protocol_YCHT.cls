VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Protocol_YCHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  Protocol_YCHT   :   Sunday, June 15, 2003
''  Author          :   Coozzzzz (Christopher Brim)
''  Contact         :   Coozzzzz@yahoo.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  Simply put this class module encapsulates the basic foundation of the YCHT
''  protocol allowing for easy use and abilities for multiple instances of it
''  within a few additional lines.
''
''  This is free to use within your projects or to learn by if you can manage it
''  without proper comments. Credit for using this within your projects is not
''  needed unless you wish to do so. However, kindly don't repost this anywhere
''  unless the comments are remained intact.
''
''  As for credits go most of the packets and methods used have been from packet
''  monitoring and experience and slight descriptions from texts online. The two
''  byte packet length method used has been used from a previous source and is
''  by-far the most reasonable method unless bugs are found from it.
''
''  If you happen to trigger the "Unhandled Packet" error and/or wish to improve
''  the methods I use please contact me and I could possibly improve them on this
''  as well if they are better implemented methods than the ones i've used. This
''  also applies to bugs you find within this. As I do not test this much I don't
''  run into too many. :)
''
''  -Coozzzzz
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Server_Login As String 'login.yahoo.com
Public Server_Chat As String  'cs8.chat.yahoo.com
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private chatUsername As String
Private chatPassword As String
Private chatRoom As String
Private chatCookie As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum FriendStatus
    GamesJoined = 0
    GamesLeft = 1
    OnlineTrue = 2
    OnlineFalse = 3
    ChatJoined = 4
    ChatLeft = 5
    OnlineInChat = 6
End Enum
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Event Away(strMsg As String)
Public Event Connected(isConnected As Boolean)
Public Event Error(strError As String)
Public Event FriendStatus(strFriend As String, fStatus As FriendStatus)
Public Event ReceivedEmail(emailCount As String)
Public Event ReceivedEmote(strUser As String, strMsg As String)
Public Event ReceivedInvite(strRoom As String, strUser As String)
Public Event ReceivedMessage(strUser As String, strMsg As String)
Public Event ReceivedPrivateMessage(strUser As String, strMsg As String)
Public Event RoomJoined(strRoom As String, strRoomTopic As String)
Public Event UserEntered(strUser As String)
Public Event UserLeft(strUser As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Requires a reference to "Mswinsck.ocx"
Private WithEvents Winsock As MSWinsockLib.Winsock
Attribute Winsock.VB_VarHelpID = -1
Private Sub Class_Initialize()
    Set Winsock = New MSWinsockLib.Winsock
End Sub
Private Sub Class_Terminate()
    Winsock.Close
    Set Winsock = Nothing
End Sub
Public Property Let Login_Username(strUsername As String)
    chatUsername = strUsername
End Property
Public Property Let Login_Password(strPassword As String)
    chatPassword = strPassword
End Property
Private Function Cookie_Get(strUsername As String, strPassword As String) As String
    Cookie_Get = "GET /config/login?.tries=1&.src=chat&.last=&promo=&lg=&.intl=&.bypass=&.chkP=Y&login=" & strUsername & "&passwd=" & strPassword & " HTTP/1.1" & vbCrLf
    Cookie_Get = Cookie_Get & "Accept: */*" & vbCrLf
    Cookie_Get = Cookie_Get & "Accept -Language: en -us" & vbCrLf
    Cookie_Get = Cookie_Get & "User-Agent: Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.1)" & vbCrLf
    Cookie_Get = Cookie_Get & "Host: login.yahoo.com" & vbCrLf
    Cookie_Get = Cookie_Get & "Connection: Close" & vbCrLf & vbCrLf
End Function
Private Function Cookie_Parse(strCookie As String)
    Dim Pos1 As Integer, Pos2 As Integer
    Pos1 = InStr(1, strCookie, "Y=v=")
    Pos2 = InStr(Pos1 + 1, strCookie, ";")
    If Pos1 > 0 And Pos2 > 0 Then
        Cookie_Parse = Mid(strCookie, Pos1, Len(strCookie) - (Len(Mid(strCookie, 1, Pos1 - 1) + Mid(strCookie, Pos2 + 1))))
        strCookie = Mid(strCookie, Pos2 + 1)
        Pos1 = InStr(1, strCookie, "T=z=")
        Pos2 = InStr(Pos1 + 1, strCookie, ";")
        If Pos1 > 0 And Pos2 > 0 Then
            Cookie_Parse = Cookie_Parse & " " & Mid(strCookie, Pos1, Len(strCookie) - 1 - (Len(Mid(strCookie, 1, Pos1 - 1) + Mid(strCookie, Pos2 + 1))))
        Else
            Cookie_Parse = "Invalid"
            Exit Function
        End If
    Else
        Cookie_Parse = "Invalid"
        Exit Function
    End If
End Function
Private Function Packet_Assemble(packetType As String, packetData As String) As String
    If Len(packetData) > 512 Then packetData = Left(packetData, 512)
    Packet_Assemble = "YCHT" & Chr$(0) & Chr$(0) & Chr$(1) & Chr$(102) & Chr$(0) & Chr$(0) & Chr(0) & packetType & Chr$(0) & Chr$(0) & (Chr(Int(Len(packetData) / 256)) & Chr(Int(Len(packetData) Mod 256))) & packetData
End Function
Private Sub Packet_Handle(strRecv As String)
    Dim packets() As String, i As Integer
    packets = Split(strRecv, "YCHT" & Chr(0) & Chr(0))
    If IsArray(packets) = True Then
        For i = 1 To UBound(packets)
            packets(i) = "YCHT" & Chr(0) & Chr(0) & packets(i)
            Packet_Parse Mid(packets(i), 17), Asc(Mid(packets(i), 12, 1))
        Next i
    End If
End Sub
Private Sub Packet_Parse(strRecv As String, action As Integer)
    Dim strPacket() As String, strUsers() As String
    Dim i As Integer, j As Integer
    strPacket = Split(strRecv, Chr(&HC0) & Chr(&H80))
    If IsArray(strPacket) = True Then
        If strPacket(0) = "*** Danger Will Robinson!!!" Then
            RaiseEvent Error(strPacket(0))
            Exit Sub
        End If
        Select Case action
            Case "1"    'Connected
                RaiseEvent Connected(True)
            Case "2"    'Disconnected
                RaiseEvent Error(strPacket(0))
                RaiseEvent Connected(False)
            Case "17"   'User Enter
                Select Case strPacket(0)
                    Case "That room is full.  Try a similar room?"
                        RaiseEvent Error(strPacket(0))
                        Exit Sub
                    Case Else
                        If InStr(1, strRecv, chatUsername) > 0 Then
                            chatRoom = strPacket(0)
                            RaiseEvent RoomJoined(chatRoom, strPacket(1))
                        End If
                End Select
                strUsers = Split(strPacket(4), "")
                For i = 0 To UBound(strUsers) Step 5
                    If Len(strUsers(i)) > 1 Then RaiseEvent UserEntered(Replace(strUsers(i), Chr(32) & Chr(1), ""))
                Next i
            Case "18"   'User Leave
                RaiseEvent UserLeft(strPacket(1))
            Case "23"   'Invitation
                If InStr(1, strPacket(0), "has been invited to join this room", vbTextCompare) > 0 Then
                    'No need to notify user when person has been invited since they
                    'probably won't get it anyways.. seems to be a yahoo! problem
                Else
                    RaiseEvent ReceivedInvite(strPacket(1), strPacket(2))
                End If
            Case "65"   'Message
                RaiseEvent ReceivedMessage(strPacket(1), strPacket(2))
            Case "67"   'Emote
                RaiseEvent ReceivedEmote(strPacket(1), strPacket(2))
            Case "69"   'Private Message
                RaiseEvent ReceivedPrivateMessage(strPacket(0), strPacket(2))
            Case "70"   'User Away/Back
                RaiseEvent Away(strPacket(0))
            Case "101"  'Logoff successful
                RaiseEvent Error(strPacket(0))
                RaiseEvent Connected(False)
            Case "104"  'Friend Online/Offline
                For i = 0 To UBound(strPacket) Step 3
                    Select Case strPacket(i)
                        Case "0" 'Offline
                            Select Case strPacket(i + 2)
                                Case "1" 'Games
                                    RaiseEvent FriendStatus(strPacket(i + 1), GamesLeft)
                                Case "2" 'Chat
                                    RaiseEvent FriendStatus(strPacket(i + 1), ChatLeft)
                                Case "4" 'Online
                                    RaiseEvent FriendStatus(strPacket(i + 1), OnlineFalse)
                            End Select
                        Case "1" 'Online
                            Select Case strPacket(i + 2)
                                Case "1" 'Games
                                    RaiseEvent FriendStatus(strPacket(i + 1), GamesJoined)
                                Case "2" 'Chat
                                    RaiseEvent FriendStatus(strPacket(i + 1), ChatJoined)
                                Case "4" 'Online
                                    RaiseEvent FriendStatus(strPacket(i + 1), OnlineTrue)
                            End Select
                        Case "3" 'Multiple Users Online
                            'Having problems figuring out which packets are for which statuses
                            'so the below will be ignored
                            Exit Sub
                            'Debug.Print Replace(strRecv, Chr(&HC0) & Chr(&H80), "  ")
                            For j = 1 To UBound(strPacket) Step 2
                                Select Case strPacket(j + 1)
                                    Case "1" 'Chat
                                        RaiseEvent FriendStatus(strPacket(j), ChatJoined)
                                    Case "2" 'Online
                                        RaiseEvent FriendStatus(strPacket(j), OnlineTrue)
                                    Case "3" 'Online In Chat
                                        RaiseEvent FriendStatus(strPacket(j), OnlineInChat)
                                End Select
                            Next j
                            Exit For
                    End Select
                Next i
            Case "105"  'New E-mail
                RaiseEvent ReceivedEmail(strPacket(0))
            Case Else   'Unknown Packet
                RaiseEvent Error("Unhandled packet. Report this number (" & action & ") and possibly what the purpose of it is and i'll implement it.")
        End Select
    End If
End Sub
Private Function ErrorHandle_Winsock() As Boolean
    If Winsock.State = 7 Then
        ErrorHandle_Winsock = True
    Else
        ErrorHandle_Winsock = False
        RaiseEvent Error("Winsock control not connected.")
    End If
End Function
Public Sub ychtConnect()
    chatCookie = ""
    Winsock.Close
    Winsock.Connect Server_Login, 80
End Sub
Public Sub ychtClose()
    Winsock.Close
    RaiseEvent Connected(False)
End Sub
Public Sub ychtJoinRoom(strRoomName As String)
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(113), "join " & strRoomName)
End Sub
Public Sub ychtSendMessage(strMsg As String)
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(65), chatRoom & Chr(1) & strMsg)
End Sub
Public Sub ychtSendPrivateMessage(strUser As String, strMsg As String)
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(113), "tell " & strUser & " " & strMsg)
End Sub
Public Sub ychtSendInvite(strUser As String)
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(113), "invite " & strUser)
End Sub
Public Sub ychtGotoUser(strUser As String)
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(113), "goto " & strUser)
End Sub
Public Sub ychtSendEmote(strMsg As String)
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(67), chatRoom & Chr(1) & strMsg)
End Sub
Public Sub ychtPing()
    If ErrorHandle_Winsock() = True Then Winsock.SendData Packet_Assemble(Chr(98), "")
End Sub
Private Sub Winsock_Close()
    RaiseEvent Connected(False)
End Sub
Private Sub Winsock_Connect()
    If Len(chatCookie) = 0 Then
        Winsock.SendData Cookie_Get(chatUsername, chatPassword)
    Else
        Winsock.SendData Packet_Assemble(Chr(1), chatUsername & Chr(1) & chatCookie)
    End If
End Sub
Private Sub Winsock_DataArrival(ByVal bytesTotal As Long)
    Dim recv As String
    Winsock.GetData recv
    'Debug.Print recv
    If Len(chatCookie) = 0 Then
        chatCookie = Cookie_Parse(recv)
        If chatCookie = "Invalid" Then
            RaiseEvent Error("Invalid cookie returned by server." & vbCrLf & vbCrLf & "Possibly Invalid Username/Password Combination.")
            Winsock.Close
        Else
            Winsock.Close
            Winsock.Connect Server_Chat, 8002
        End If
    Else
        Packet_Handle recv
    End If
End Sub
