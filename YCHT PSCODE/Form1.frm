VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "YCHT Protocol Class"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   8550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3900
      Width           =   5535
   End
   Begin VB.TextBox Text1 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   120
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   5760
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  This is just a demonstration form for using the Class Module provided with
''  this project. This form is commented to help you understand the methods and
''  events returned by an instance of the class module but the module itself is
''  not commented well. Read the headers within the Class Module for further info.
''
''  Btw, I know this form does not look "good". I actually spent more time
''  commenting it than working on the form. It's a waste of time for me to make
''  it look "good" when it's just a demonstration of the class module.
''
''  -Coozzzzz
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Define our variable to our class module with events (WithEvents)
Private WithEvents ycht As Protocol_YCHT
Attribute ycht.VB_VarHelpID = -1
Private Sub Form_Load()
    'Initialize our class module
    Set ycht = New Protocol_YCHT
    'Set the Chat and Login servers we'll be using
    ycht.Server_Chat = "cs8.chat.yahoo.com"
    ycht.Server_Login = "login.yahoo.com"
    'Change the below Username/Password to the one you wish to use
    ycht.Login_Username = "Username"
    ycht.Login_Password = "Password"
    'Connect to YCHT Server
    ycht.ychtConnect
End Sub
Private Sub ycht_Away(strMsg As String)
    'Event is fired when a user goes away or comes back
    '----
    'strMsg includes the username and message such as "coozzzzz is back."
    MsgBox strMsg, vbExclamation
End Sub
Public Sub ycht_Connected(isConnected As Boolean)
    'Event is fired when our connection state changes
    '----
    'isConnected=true   : when connected
    'isConnected=false  : when disconnected
    If isConnected = True Then ycht.ychtJoinRoom "The Bored Room:1"
End Sub
Private Sub ycht_FriendStatus(strFriend As String, fStatus As FriendStatus)
    'Event is fired when a friend's status changes to Online/Offline/Chat/Games
    '----
    'I've used an Enum to handle the Statuses and should be self explanatory
    'in the below usage
    Select Case fStatus
        Case FriendStatus.ChatJoined
            Text1.Text = Text1.Text & strFriend & " in in chat." & vbCrLf & vbCrLf
        Case FriendStatus.ChatLeft
            Text1.Text = Text1.Text & strFriend & " left chat." & vbCrLf & vbCrLf
        Case FriendStatus.GamesJoined
            Text1.Text = Text1.Text & strFriend & " is in games." & vbCrLf & vbCrLf
        Case FriendStatus.GamesLeft
            Text1.Text = Text1.Text & strFriend & " left games." & vbCrLf & vbCrLf
        Case FriendStatus.OnlineFalse
            Text1.Text = Text1.Text & strFriend & " is offline." & vbCrLf & vbCrLf
        Case FriendStatus.OnlineTrue
            Text1.Text = Text1.Text & strFriend & " is online." & vbCrLf & vbCrLf
    End Select
End Sub
Private Sub ycht_ReceivedEmail(emailCount As String)
    'Event is fired when you receive a new e-mail on the name you're currently using
    '----
    'emailCount     : Count of how many emails you have
    MsgBox "We've received a new e-mail! (total of " & emailCount & " email(s)).", vbInformation
End Sub
Private Sub ycht_ReceivedInvite(strRoom As String, strUser As String)
    'Event is fired when you receive an invitation. (from testing i've noticed there
    'are problems with this between YMSG/YCHT.. however it works if invited from a
    'YMSG protocol user."
    '----
    'strRoom    : The room you are invited to
    'strUser    : The user who invited you
    Dim lRet As Long
    lRet = MsgBox("You are invited to '" & strRoom & "' by " & strUser & ".", vbQuestion + vbYesNo, "Invitation")
    If lRet = vbYes Then ycht.ychtJoinRoom strRoom
End Sub
Private Sub ycht_ReceivedPrivateMessage(strUser As String, strMsg As String)
    'Event is fired when you receive a private message
    '----
    'strUser    : User who sent you the private message
    'strMsg     : The message the User sent
    Text1.Text = Text1.Text & "(Private Message) " & strUser & " says, " & parse_HTML(strMsg) & vbCrLf & vbCrLf
End Sub
Public Sub ycht_Error(strError As String)
    'Event is fired when a handled error occurs within the Class Module
    '----
    'strError   : The error message
    MsgBox strError, vbCritical
End Sub
Private Sub ycht_RoomJoined(strRoom As String, strRoomTopic As String)
    'Event is fired when you join a new room. I've added this so you know when you
    'should clear your User List to add new Users.
    '----
    'strRoom        : The room you joined
    'strRoomTopic   : The topic of the room you joined
    List1.Clear
    Form1.Caption = App.Title & " - " & strRoom
End Sub
Private Sub ycht_UserEntered(strUser As String)
    'Event is fired when a new user joins the room you're in. It also fires multiple
    'times when you join a new room. Optionally if you wish to ignore users you should
    'do this here. Have it check if the user is ignored and if so, do not add them to
    'the list. This is good so you can check through the UserList on who's messages
    'to display instead of checking through the entire ignore list for each message.
    '----
    'strUser    : The user that joined the room
    Dim i As Integer
    'Check if user exists on our list before adding them
    For i = List1.ListCount - 1 To 0 Step -1
        'If user already exists on list then exit sub
        If StrComp(strUser, List1.List(i), vbTextCompare) = 0 Then Exit Sub
    Next i
    'Add user to list
    List1.AddItem strUser
End Sub
Private Sub ycht_UserLeft(strUser As String)
    'Event is fired when a user leaves the room you're currently in
    '----
    'strUser    : The user that left the room
    Dim i As Integer
    'Loop through our UserList
    For i = List1.ListCount - 1 To 0 Step -1
        'If the user exists in our list then remove them
        '(suggested that you should compare each list(i) with strUser in lowercase)
        If StrComp(strUser, List1.List(i), vbTextCompare) = 0 Then List1.RemoveItem i
    Next i
End Sub
Private Sub ycht_ReceivedMessage(strUser As String, strMsg As String)
    'Event is fired when you receive a new message within the room you're currently in
    '----
    'strUser    : The user that sent the message
    'strMsg     : The message the user sent
    Text1.Text = Text1.Text & strUser & " says, " & parse_HTML(strMsg) & vbCrLf & vbCrLf
End Sub
Private Sub ycht_ReceivedEmote(strUser As String, strMsg As String)
    'Event is fired when you receive a new emote within the room you're currently in
    '----
    'strUser    : The user that sent the message
    'strMsg     : The message the user sent
    Text1.Text = Text1.Text & strUser & " " & parse_HTML(strMsg) & vbCrLf & vbCrLf
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    'The below is just to handle chat commands such as /join, /goto, /invite, etc
    Dim tSplit() As String
    If KeyCode = "13" Then
        If Len(Text2.Text) > 0 Then
            If Mid(Text2.Text, 1, 5) = "/join" Then
                ycht.ychtJoinRoom Mid(Text2.Text, 7)
            ElseIf Mid(Text2.Text, 1, 3) = "/pm" Then
                tSplit = Split(Mid(Text2.Text, 4), " ")
                If IsArray(tSplit) Then
                    If UBound(tSplit) = 2 Then ycht.ychtSendPrivateMessage tSplit(1), tSplit(2)
                End If
            ElseIf Mid(Text2.Text, 1, 5) = "/goto" Then
                ycht.ychtGotoUser Mid(Text2.Text, 7)
            ElseIf Mid(Text2.Text, 1, 1) = ":" Then
                ycht.ychtSendEmote Mid(Text2.Text, 2)
            ElseIf Mid(Text2.Text, 1, 7) = "/invite" Then
                ycht.ychtSendInvite Mid(Text2.Text, 9)
            Else
                ycht.ychtSendMessage Text2.Text
            End If
            Text2.Text = ""
        End If
    End If
End Sub
Private Sub Text1_Change()
    'The below just keeps the last line in the text box visible (auto-scrolling)
    Text1.SelStart = Len(Text1.Text)
End Sub
Private Function parse_HTML(strCheck As String) As String
    'The below just parses certain html tags from strings.. it's a bit rough but
    'it's not that important to optimize at the moment
    Dim Pos1 As Integer, Pos2 As Integer
reparse1:
    Pos1 = InStr(1, LCase(strCheck), "")
    If Pos1 > 0 Then
        Pos2 = InStr(Pos1, LCase(strCheck), "m")
        If Pos2 > 0 Then
            If Pos1 = 1 Then
                strCheck = Mid(strCheck, Pos2 + 1)
            Else
                strCheck = Mid(strCheck, 1, Pos1 - 1) & Mid(strCheck, Pos2 + 1)
            End If
        Else
            parse_HTML = strCheck
            GoTo reparse2
        End If
    Else
        parse_HTML = strCheck
        GoTo reparse2
    End If
    GoTo reparse1
reparse2:
    Pos1 = InStr(1, LCase(strCheck), "<font")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "</")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "</")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "<b")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "<alt")
    If Pos1 = 0 Then Pos1 = InStr(1, LCase(strCheck), "<fade")
    If Pos1 > 0 Then
        Pos2 = InStr(Pos1, LCase(strCheck), ">")
        If Pos2 > 0 Then
            If Pos1 = 1 Then
                strCheck = Mid(strCheck, Pos2 + 1)
            Else
                strCheck = Mid(strCheck, 1, Pos1 - 1) & Mid(strCheck, Pos2 + 1)
            End If
        Else
           parse_HTML = strCheck
            Exit Function
        End If
    Else
        parse_HTML = strCheck
        Exit Function
    End If
    GoTo reparse2
End Function
