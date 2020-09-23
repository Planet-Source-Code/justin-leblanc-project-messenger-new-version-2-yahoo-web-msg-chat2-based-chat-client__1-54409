Attribute VB_Name = "MODHandles"
Option Explicit

Public Function AuthenticationHandle(Data As String, Cookie As String, Authentication As Winsock, YahooChat2 As Winsock)
Dim IntC(1 To 2) As Integer
Dim CParts(1 To 2) As String
  If InStr(1, Data, "ERROR: Invalid NCC Login") Then
    Authentication.Close
    FrmMessenger.lblLogin.Caption = "Login to Yahoo!"
  Else
    IntC(1) = InStr(1, Data, "Y=v=")
  If IntC(1) = 0 Then
    Authentication.Close
    FrmMessenger.lblLogin.Caption = "Login to Yahoo!"
  Else
    IntC(2) = InStr(IntC(1) + 1, Data, ";") + 1
    CParts(1) = Mid(Data, IntC(1), IntC(2) - IntC(1))
    '
    IntC(1) = InStr(1, Data, "T=z=")
    IntC(2) = InStr(IntC(1) + 1, Data, ";")
    CParts(2) = Mid(Data, IntC(1), IntC(2) - IntC(1))
    Cookie = CParts(1) & Chr(&H20) & CParts(2)
    '
    Authentication.Close
    Call SockConnect(YServer(2), YPort(2), YahooChat2)
  End If
  End If
End Function

Public Function YahooChat2Handle(Data As String, YahooChat2 As Winsock)
Dim WhoName As String
Dim WhoNick As String
Dim UserMessage As String
Dim MessageType As String
Select Case Mid(Data, 12, 1)
  Case Chr(&H1E) 'Login
    SessionKey(2) = Mid(Data, 17, 4)
  Case Chr(&H55) 'buddylist/alias
    YahooChat2.SendData LoginToChat(YCurrentId, SessionKey(2))
  Case Chr(&H96) 'Chat Login Confermation
    ''''''
  Case Chr(&H6)  'User Messages you
'59 4D-53 47 00 00 00 00 00 60                     YMSG.....`
'00 06 00 00 00 01 EB 19-F2 F6 35 C0 80 64 72 61   ......Î.Úˆ5¿Ädra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 34 C0 80 71 71   ma_stinks¿Ä4¿Äqq
'2E 37 38 39 C0 80 31 34-C0 80 68 65 79 C0 80 31   .789¿Ä14¿Ähey¿Ä1
'35 30 C0 80 30 C0 80 31-35 31 C0 80 71 71 2E 37   50¿Ä0¿Ä151¿Äqq.7
'38 39 20 71 71 2E 37 38-39 20 63 38 39 39 30 62   89 qq.789 c8990b
'32 31 2F 36 39 37 30 2F-37 37 66 39 39 63 30 34   21/6970/77f99c04
'C0 80 31 35 32 C0 80 30-C0 80                     ¿Ä152¿Ä0¿Ä
 'If InStr(1, Data, "4¿Ä") > 0 Then
 '   WhoName = Split(Data, "4¿Ä")(1)
 '   WhoName = Split(WhoName, "¿Ä")(0)
 '   UserMessage = Split(Data, "¿Ä14¿Ä")(1)
 '   UserMessage = Split(UserMessage, "¿Ä")(0)
 '   Call SetPMText(WhoName, UserMessage)
 'End If
  Case Chr(&H98) 'Users/User joins Room /Room Info (Room space,Voice Token)
    Call ParseChatUsers(Data, FrmChat.ListView, FrmChat.RTBChat)
    FrmChat.ListView.ColumnHeaders(1).Text = "Users: " & FrmChat.ListView.ListItems.Count
        'roomname and room message parse
        If InStr(1, Data, "¿Ä128¿Ä") > 0 Then
          InChat = True
          ChatRoom = Split(Data, "104¿Ä")(1)
          ChatRoom = Split(ChatRoom, "¿Ä")(0)
        If InStr(1, Data, "¿Ä105¿Ä") > 0 Then
          ChatMessage = Split(Data, "¿Ä105¿Ä")(1)
          ChatMessage = Split(ChatMessage, "¿Ä")(0)
        Else
          ChatMessage = ""
        End If
        If InStr(1, Data, "¿Ä130¿Ä") > 0 Then
        'voice token and roomspace parse
          YVoiceToken = Split(Data, "¿Ä130¿Ä")(1)
          YVoiceToken = Split(YVoiceToken, "¿Ä")(0)
          YRoomSpace = Split(Data, "¿Ä129¿Ä")(1)
          YRoomSpace = Split(YRoomSpace, "¿Ä")(0)
        End If
        FrmChat.Caption = ChatRoom & " ~ Chat"
        '
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text)
        FrmChat.RTBChat.SelText = vbCrLf & ChatRoom & " (" & ChatMessage & ")" & vbCrLf & vbCrLf
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text) - Len(ChatRoom & " (" & ChatMessage & ")") - 4
        FrmChat.RTBChat.SelLength = Len(ChatRoom) + 1
        FrmChat.RTBChat.SelColor = &HC000&
        FrmChat.RTBChat.SelBold = True
        FrmChat.RTBChat.SelFontSize = 10
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text)
      End If
  Case Chr(&H9B) 'User Leaves
'59 4D-53 47 6C 79 31 33 00 5E                     YMSGly13.^
'00 9B 00 00 00 01 94 95-C4 A0 31 30 34 C0 80 51   .õ....îïƒ†104¿ÄQ
'51 27 73 20 52 6F 6F 6D-3A 31 C0 80 31 30 35 C0   Q's Room:1¿Ä105¿
'80 57 65 6C 63 6F 6D 65-20 74 6F 20 4D 79 20 52   ÄWelcome to My R
'6F 6F 6D C0 80 31 30 38-C0 80 31 C0 80 31 30 39   oom¿Ä108¿Ä1¿Ä109
'C0 80 71 71 2E 37 38 39-C0 80 31 31 32 C0 80 30   ¿Äqq.789¿Ä112¿Ä0
'C0 80 31 31 33 C0 80 33-33 37 39 32 C0 80 31 34   ¿Ä113¿Ä33792¿Ä14
'31 C0 80 51 51 C0 80 00-                          1¿ÄQQ¿Ä.
  If Mid(Data, 13, 4) = "ˇˇˇˇ" Then Exit Function
    WhoName = Split(Data, "¿Ä109¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    Call RemoveItem(WhoName, FrmChat.ListView)
  If FrmMessenger.mnuFilterJoinLeave.Checked = True Then Exit Function
      If LCase(WhoName) = LCase(YCurrentId) Then Exit Function
      If InStr(1, Data, "¿Ä141¿Ä") > 0 Then
        WhoNick = Split(Data, "¿Ä141¿Ä")(1)
        WhoNick = Split(WhoNick, "¿Ä")(0)
      End If
      If Len(WhoNick) > 0 Then
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text)
        FrmChat.RTBChat.SelText = WhoNick & " (" & WhoName & ") Leaves Room" & vbCrLf
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text) - Len(WhoNick & " (" & WhoName & ") Leaves Room") - 2
        FrmChat.RTBChat.SelLength = Len(WhoNick & " (" & WhoName & ")")
        FrmChat.RTBChat.SelColor = &HFF&
        FrmChat.RTBChat.SelItalic = True
        FrmChat.RTBChat.SelFontSize = 10
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text)
      Else
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text)
        FrmChat.RTBChat.SelText = WhoName & " Enters Room" & vbCrLf
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text) - Len(WhoName & " Leaves Room") - 2
        FrmChat.RTBChat.SelLength = Len(WhoName)
        FrmChat.RTBChat.SelColor = &HFF&
        FrmChat.RTBChat.SelItalic = True
        FrmChat.RTBChat.SelFontSize = 10
        FrmChat.RTBChat.SelStart = Len(FrmChat.RTBChat.Text)
      End If
  Case Chr(&HA8) 'user send message in chat
'59 4D-53 47 37 00 00 00 00 4F                     YMSG7....O
'00 A8 00 00 00 01 00 00-00 00 31 30 34 C0 80 47   .®........104¿ÄG
'65 6E 65 72 61 74 69 6F-6E 20 58 3A 32 34 C0 80   eneration X:24¿Ä
'31 30 39 C0 80 6E 73 74-79 5F 62 72 6F 6F 6B 65   109¿Änsty_brooke
'31 33 C0 80 31 31 37 C0-80 74 65 6C 6C 20 6D 65   13¿Ä117¿Ätell me
'20 77 68 61 74 20 79 6F-75 20 6C 69 6B 65 C0 80    what you like¿Ä
'31 32 34 C0 80 31 C0 80-00                        124¿Ä1¿Ä.
    WhoName = Split(Data, "¿Ä109¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    UserMessage = Split(Data, "¿Ä117¿Ä")(1)
    UserMessage = Split(UserMessage, "¿Ä")(0)
    MessageType = Split(Data, "¿Ä124¿Ä")(1)
    MessageType = Split(MessageType, "¿Ä")(0)
      If CheckIfIgnored(WhoName, FrmChat.ListView) = True Then Exit Function
      If MessageType = "1" Then
        Call SetChatText(SetNickName(WhoName, FrmChat.ListView), MessageFilter(UserMessage), FrmChat.RTBChat)
      ElseIf MessageType = "2" Then
        Call SetEmoteChatText(SetNickName(WhoName, FrmChat.ListView), MessageFilter(UserMessage), FrmChat.RTBChat)
      ElseIf MessageType = "3" Then
        Call SetThinkChatText(SetNickName(WhoName, FrmChat.ListView), MessageFilter(UserMessage), FrmChat.RTBChat)
      End If
End Select
End Function

Public Function WebMessengerHandle(Data As String, WebMsg As Winsock)
Dim UserMessage As String
Dim WhoName As String
Dim FriendsList As String
Dim Frm As New FrmAdd
Dim Frm2 As New FrmRemove
Dim Frm3 As New FrmUserAdd
Dim AuthUrl As String
Dim AuthToken As String
Dim X As Integer
Dim i As Long
Dim Names() As String
Dim Buffer As String

Select Case Mid(Data, 12, 1)
  Case Chr(&H55) 'buddylist/alias
    SessionKey(1) = Mid(Data, 17, 4)
With FrmMessenger
    .lblLogin.Caption = "Login to Yahoo!"
    .FMHide.Visible = False
    .TreeView.Visible = True
    .Toolbar.Visible = True
    .CBBStatus.Visible = True
    .mnuSignin.Caption = "Sign Out"
    .mnuTools.Visible = True
    .mnuView.Visible = True
  If FrmLogin.CKBGeneral(2).Value = 1 Then
    .CBBStatus.ComboItems(2).Selected = True
  If WebMsg.State = sckConnected Then WebMsg.SendData WebMsgInvisible(SessionKey(1))
  Else
    .CBBStatus.ComboItems(1).Selected = True
  End If
  If InStr(1, Data, "87¿Ä") > 0 Then
    Call ParseIdentitys(Data)
    Call ParseBuddiesAndGroups(Data, .TreeView)
  End If
End With
    FrmLogin.CMDGeneral(0).Caption = "Sign Out"
    ''''''
    Call SockConnect(YServer(3), YPort(3), FrmLogin.SockChat)
  Case Chr(&H28) 'authentication reply
    'if add user form
    If AddUserForm = True Then
      Frm.CBBIdentity.Text = YCurrentId
        For i = 0 To UBound(YId)
          If Len(YId(i)) > 0 Then Frm.CBBIdentity.AddItem YId(i)
        Next i
    If Len(YGroups(1)) > 0 Then
      For X = 1 To UBound(YGroups)
        Frm.CBBGroup.Text = YGroups(1)
        If Len(YGroups(X)) > 0 Then Frm.CBBGroup.AddItem YGroups(X)
      Next X
    End If
      AuthUrl = Split(Data, "225¿Ä")(1)
      AuthToken = Split(AuthUrl, "226¿Ä")(1)
      AuthToken = Split(AuthToken, "¿Ä")(0)
      AuthUrl = Split(AuthUrl, "¿Ä226")(0)
      Frm.txtToken.Text = AuthToken
      Frm.WBImg.Navigate AuthUrl
      Frm.txtWho.Text = AddName
      AddName = ""
      Frm.Show
    Else
    'if remove user form
      AuthUrl = Split(Data, "225¿Ä")(1)
      AuthToken = Split(AuthUrl, "226¿Ä")(1)
      AuthToken = Split(AuthToken, "¿Ä")(0)
      AuthUrl = Split(AuthUrl, "¿Ä226")(0)
      Frm2.txtToken.Text = AuthToken
      Frm2.WBImg.Navigate AuthUrl
      Frm2.txtWho.Text = RemoveName
      RemoveName = ""
      Frm2.Show
    End If
  Case Chr(&H4B) ' User Typing
'59 4D-53 47 00 00 00 00 00 40                     YMSG.....@
'00 4B 00 00 00 01 F0 D1-2A 0B 35 C0 80 6A 61 79   .K....—*.5¿Äjay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 34 C0 80-71 71 2E 37 38 39 C0 80   _up¿Ä4¿Äqq.789¿Ä
'31 34 C0 80 20 C0 80 31-33 C0 80 31 C0 80 34 39   14¿Ä ¿Ä13¿Ä1¿Ä49
'C0 80 54 59 50 49 4E 47-C0 80                     ¿ÄTYPING¿Ä
    WhoName = Split(Data, "4¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    Call Typing(WhoName)
  Case Chr(&H6) 'User Messages You/Users status message sent
'59 4D-53 47 00 00 00 00 00 44                     YMSG.....D
'00 06 00 00 00 01 F0 D1-2A 0B 35 C0 80 6A 61 79   ......—*.5¿Äjay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 34 C0 80-71 71 2E 37 38 39 C0 80   _up¿Ä4¿Äqq.789¿Ä
'39 37 C0 80 31 C0 80 31-34 C0 80 79 6F C0 80 36   97¿Ä1¿Ä14¿Äyo¿Ä6
'33 C0 80 3B 30 C0 80 36-34 C0 80 30 C0 80         3¿Ä;0¿Ä64¿Ä0¿Ä
    
 'there is one flaw in using web messenger protocol you MUS have the user
 'on your buddy list to message them back
 'but since i combined chat2 protocol and web messenger protocol you can :P
 'chat2 allows you to pm them even if the user is not on your list, plus it allows you to go into chat
 If InStr(1, Data, "4¿Ä") > 0 Then
    WhoName = Split(Data, "4¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
 If DenyPmNonBuddy(WhoName, FrmMessenger.mnuNonfriends, FrmMessenger.TreeView) = False Then Exit Function
 If FrmMessenger.mnuAll.Checked = True Then Exit Function
    UserMessage = Split(Data, "¿Ä14¿Ä")(1)
    UserMessage = Split(UserMessage, "¿Ä")(0)
    Call SetPMText(WhoName, MessageFilter(UserMessage))
 End If
 'YMSG     1    √Y 5¿Äqq.789¿Ä10¿Ä99¿Ä19¿ÄBroken(Featuring Amy Lee¿Ä
 If InStr(1, Data, "5¿Ä") > 0 And InStr(1, Data, "19¿Ä") > 0 Then
    WhoName = Split(Data, "5¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    UserMessage = Split(Data, "¿Ä19¿Ä")(1)
    UserMessage = Split(UserMessage, "¿Ä")(0)
    '''set user status
 End If
  Case Chr(&HF) 'add request
       If FrmMessenger.mnuBlockAdd.Checked = True Then Exit Function
       If InStr(1, Data, "7¿Ä") > 0 Then
        WhoName = Split(Data, "7¿Ä")(1)
        WhoName = Split(WhoName, "¿Ä")(0)
        i = GetNodeIndex(WhoName, FrmMessenger.TreeView)
      If i > 0 Then FrmMessenger.TreeView.Nodes.Item(i).Image = 2
      End If
      If InStr(1, Data, "3¿Ä") > 0 And InStr(1, Data, "14¿Ä") > 0 Then
        WhoName = Split(Data, "3¿Ä")(1)
        WhoName = Split(WhoName, "¿Ä")(0)
        UserMessage = Split(Data, "14¿Ä")(1)
        UserMessage = Split(UserMessage, "¿Ä")(0)
        Frm3.txtWho.Text = LCase(WhoName)
        Frm3.txtMessage.Text = UserMessage
        Frm3.Show
      End If
  Case Chr(&H1)  ' User Comes Online
'59 4D-53 47 00 00 00 00 00 69                     YMSG.....i
'00 01 00 00 00 01 F0 DF-93 30 30 C0 80 64 72 61   ......ﬂì00¿Ädra
'6D 61 5F 73 74 69 6E 6B-73 C0 80 31 C0 80 64 72   ma_stinks¿Ä1¿Ädr
'61 6D 61 5F 73 74 69 6E-6B 73 C0 80 37 C0 80 71   ama_stinks¿Ä7¿Äq
'71 2E 37 38 39 C0 80 31-30 C0 80 30 C0 80 31 31   q.789¿Ä10¿Ä0¿Ä11
'C0 80 43 42 39 41 31 35-45 39 C0 80 31 37 C0 80   ¿ÄCB9A15E9¿Ä17¿Ä
'30 C0 80 31 33 38 C0 80-31 C0 80 31 33 C0 80 31   0¿Ä138¿Ä1¿Ä13¿Ä1
'C0 80 31 39 38 C0 80 30-C0 80 32 31 33 C0 80 30   ¿Ä198¿Ä0¿Ä213¿Ä0
'C0 80 00                                          ¿Ä.
      If InStr(1, Data, "¿Ä7¿Ä") > 0 Then
        Buffer = Mid(Data, 21, Len(Data) - 20)
        Names = Split(Buffer, "¿Ä7¿Ä")
          For X = 1 To UBound(Names)
              Buffer = Split(Names(X), "¿Ä10¿Ä")(0)
              i = GetNodeIndex(Buffer, FrmMessenger.TreeView)
            If i > 0 Then FrmMessenger.TreeView.Nodes.Item(i).Image = 2
          Next X
      End If
  Case Chr(&H2)  ' User Goes Offline
'59 4D-53 47 00 00 00 00 00 45                     YMSG.....E
'00 02 00 00 00 01 F0 DF-93 30 37 C0 80 71 71 2E   ......ﬂì07¿Äqq.
'37 38 39 C0 80 31 30 C0-80 30 C0 80 31 31 C0 80   789¿Ä10¿Ä0¿Ä11¿Ä
'43 42 39 41 31 35 45 39-C0 80 31 37 C0 80 30 C0   CB9A15E9¿Ä17¿Ä0¿
'80 34 37 C0 80 32 C0 80-31 33 C0 80 30 C0 80 31   Ä47¿Ä2¿Ä13¿Ä0¿Ä1
'39 38 C0 80 30 C0 80 32-31 33 C0 80 30 C0 80      98¿Ä0¿Ä213¿Ä0¿Ä
    WhoName = Split(Data, "7¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    i = GetNodeIndex(WhoName, FrmMessenger.TreeView)
      If i > 0 Then FrmMessenger.TreeView.Nodes.Item(i).Image = 1
  Case Chr(&H83)
'59 4D-53 47 00 00 00 00 00 40                     YMSG.....@
'00 83 00 00 00 01 F0 C5-C3 69 31 C0 80 6A 61 79   .É....≈√i1¿Äjay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 36 36 C0-80 30 C0 80 37 C0 80 79   _up¿Ä66¿Ä0¿Ä7¿Äy
'6F 79 6F 2D 71 71 C0 80-36 35 C0 80 43 68 61 74   oyo-qq¿Ä65¿ÄChat
'20 46 72 69 65 6E 64 73-C0 80                      Friends¿Ä
    WhoName = Split(Data, "7¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    FriendsList = Split(Data, "65¿Ä")(1)
    FriendsList = Split(FriendsList, "¿Ä")(0)
      If CheckGroups(FriendsList, FrmMessenger.TreeView) = True Then
        FrmMessenger.TreeView.Nodes.Add FriendsList, tvwChild, WhoName, WhoName, 1
      Else
        FrmMessenger.TreeView.Nodes.Add , FriendsList, FriendsList, FriendsList, 5
        FrmMessenger.TreeView.Nodes.Add FriendsList, tvwChild, WhoName, WhoName, 1
      End If
    Call AllNodes(FrmMessenger.TreeView, True)
  Case Chr(&H84)
'59 4D-53 47 00 00 00 00 00 3D                     YMSG.....=
'00 84 00 00 00 01 F2 C4-3F 13 31 C0 80 6A 61 79   .Ñ....Úƒ?.1¿Äjay
'5F 64 6F 67 5F 74 68 61-74 73 5F 77 68 61 74 73   _dog_thats_whats
'5F 75 70 C0 80 36 36 C0-80 30 C0 80 37 C0 80 6A   _up¿Ä66¿Ä0¿Ä7¿Äj
'61 79 73 5F 62 6F 74 34-C0 80 36 35 C0 80 46 72   ays_bot4¿Ä65¿ÄFr
'69 65 6E 64 73 C0 80                              iends¿Ä
    WhoName = Split(Data, "7¿Ä")(1)
    WhoName = Split(WhoName, "¿Ä")(0)
    Call RemoveUserNode(WhoName, FrmMessenger.TreeView)
End Select
End Function
