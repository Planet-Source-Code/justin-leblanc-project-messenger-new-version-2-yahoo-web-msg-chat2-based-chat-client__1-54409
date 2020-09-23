Attribute VB_Name = "MODGeneral"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'
Public YId(0 To 6) As String
Public YCurrentId As String
Public YPass As String
Public YCookie As String
Public SessionKey(1 To 2) As String
Public YGroups(1 To 100) As String
Public YServer(1 To 3) As String
Public YPort(1 To 3) As Long
Public YBuild(1 To 2) As String
Public AddName As String
Public RemoveName As String
Public AddUserForm As Boolean
Public InChat As Boolean
Public ChatRoom As String
Public ChatMessage As String
Public YVoiceToken As String
Public YRoomSpace As String

Public Function SockConnect(Server As String, Port As Long, Sock As Winsock)
  Sock.Close
  Sock.Connect Server, Port
End Function

Public Function LoadStatus(Combo As ImageCombo)
Combo.ComboItems.Add , , "Available", 1
Combo.ComboItems.Add , , "Invisible", 3
End Function

Public Function Pause(interval)
Dim X
 X = Timer
  Do While Timer - X < Val(interval)
    DoEvents
  Loop
End Function

Public Function ParseBuddiesAndGroups(Data As String, Tree As TreeView)
On Error Resume Next
Dim Groups() As String
Dim Names() As String
Dim Buffer As String
Dim GroupName As String
Dim i As Integer
Dim X As Integer
'87 - groups and friends (splitter between groups 0A)
'88 - ignore list
'89 - identitys
  Tree.Nodes.Clear
  Erase YGroups
  Tree.Nodes.Add , , "=:Groups:=", "Groups - " & YCurrentId, 3
  Tree.Nodes.Item(1).Bold = True
  Tree.Nodes.Add , , "=:Info:=", "Groups() - Friends()", 6
  Tree.Nodes.Item(2).Bold = True
  Buffer = Split(Data, "87À€")(1)
  Buffer = Split(Buffer, "À€88")(0)
    Groups = Split(Buffer, Chr(&HA))
      For i = 0 To UBound(Groups) - 1
        GroupName = Split(Groups(i), ":")(0)
        YGroups(i + 1) = GroupName
          If GroupName = "" Then GoTo None
        Tree.Nodes.Add , GroupName, GroupName, GroupName, 5
        Names = Split(Split(Groups(i), ":")(1), ",")
          For X = 0 To UBound(Names)
            Tree.Nodes.Add GroupName, tvwChild, LCase(Names(X)), Names(X), 1
          Next X
      Next i
None:
  Call AllNodes(Tree, True)
End Function

Public Function AllNodes(Tree As TreeView, Expand As Boolean)
Dim i As Integer
Dim FriendCount As Long
Dim GroupCount As Long
  For i = 1 To Tree.Nodes.Count
    If Expand = True Then
      Tree.Nodes.Item(i).Expanded = True
    Else
      Tree.Nodes.Item(i).Expanded = False
    End If
    If Tree.Nodes.Item(i).Image = 5 Or Tree.Nodes.Item(i).Image = 4 Then
      Tree.Nodes.Item(i).Bold = True
      GroupCount = GroupCount + 1
    End If
    If Tree.Nodes.Item(i).Image = 1 Or Tree.Nodes.Item(i).Image = 2 Then
      Tree.Nodes.Item(i).ForeColor = &H808080
      FriendCount = FriendCount + 1
    End If
  Next i
Tree.Nodes.Item(2).Text = "Groups(" & GroupCount & ") - Friends(" & FriendCount & ")"
Tree.Nodes.Item(2).Selected = True
End Function

Public Function ParseIdentitys(Data As String)
Dim Buffer As String
Dim Names() As String
Dim i As Integer
  Erase YId
  Buffer = Split(Data, "89À€")(1)
  Buffer = Split(Buffer, "À€59")(0)
    Names = Split(Buffer, ",")
      For i = 0 To UBound(Names)
        YId(i) = Names(i)
      Next i
End Function

Public Function LogOutMessenger()
With FrmLogin
    .SockAuthentication.Close
    .SockYahooChat2.Close
    .SockChat.Close
    .CMDGeneral(0).Caption = "Sign In"
End With
With FrmMessenger
    .lblLogin.Caption = "Login to Yahoo!"
    .FMHide.Visible = True
    .TreeView.Visible = False
    .Toolbar.Visible = False
    .CBBStatus.Visible = False
    .mnuSignin.Caption = "Sign In"
    .mnuTools.Visible = False
    .mnuView.Visible = False
End With
End Function

Public Function ExpandNodes(Tree As TreeView, Expand As Boolean)
Dim i As Integer
  For i = 1 To Tree.Nodes.Count
    If Expand = True Then
      Tree.Nodes.Item(i).Expanded = True
    Else
      Tree.Nodes.Item(i).Expanded = False
    End If
  Next i
Tree.Nodes.Item(2).Selected = True
End Function

Public Function SetPMText(WhoFrom As String, message As String)
Dim Frm As Form
Dim NewPm As New FrmPm
  For Each Frm In Forms
    If LCase(Mid(Frm.Caption, 1, Len(WhoFrom))) = LCase(WhoFrom) Then
      Frm.RTBPm.SelStart = Len(Frm.RTBPm.Text)
      Frm.RTBPm.SelText = WhoFrom & ": " & message & vbCrLf
      Frm.RTBPm.SelStart = Len(Frm.RTBPm.Text) - Len(WhoFrom & ": " & message) - 2
      Frm.RTBPm.SelLength = Len(WhoFrom) + 1
      Frm.RTBPm.SelColor = &HC0&
      Frm.RTBPm.SelBold = True
      Frm.RTBPm.SelFontSize = 10
      Frm.RTBPm.SelStart = Len(Frm.RTBPm.Text)
      Frm.StatusBar.SimpleText = "Last Message Received at " & Time
      Exit Function
    End If
  Next Frm
'
NewPm.Caption = LCase(WhoFrom) & " ~ Instant Message"
NewPm.txtWho = LCase(WhoFrom)
NewPm.txtWho.Locked = True
NewPm.RTBPm.SelStart = Len(NewPm.RTBPm.Text)
NewPm.RTBPm.SelText = WhoFrom & ": " & message & vbCrLf
NewPm.RTBPm.SelStart = Len(NewPm.RTBPm.Text) - Len(WhoFrom & ": " & message) - 2
NewPm.RTBPm.SelLength = Len(WhoFrom) + 1
NewPm.RTBPm.SelColor = &HC0&
NewPm.RTBPm.SelBold = True
NewPm.RTBPm.SelFontSize = 10
NewPm.RTBPm.SelStart = Len(NewPm.RTBPm.Text)
NewPm.StatusBar.SimpleText = "Last Message Received at " & Time
NewPm.Show
End Function

Public Function Typing(WhoFrom As String)
Dim Frm As Form
  For Each Frm In Forms
    If LCase(Mid(Frm.Caption, 1, Len(WhoFrom))) = LCase(WhoFrom) Then
      Frm.StatusBar.SimpleText = LCase(WhoFrom) & " is typing"
      Exit Function
    End If
  Next Frm
End Function

Public Function GetNodeIndex(UserName As String, Tree As TreeView) As Long
Dim i As Integer
  For i = 1 To Tree.Nodes.Count
    If UserName = Tree.Nodes.Item(i).Key Then
      GetNodeIndex = i
      Exit Function
    End If
  Next i
GetNodeIndex = 0
End Function

Public Function GetGroupName(UserName As String, Tree As TreeView) As String
Dim i As Integer
Dim PosStop As Long
Dim Group As String
  For i = 1 To Tree.Nodes.Count
    If LCase(UserName) = Tree.Nodes.Item(i).Key Then
      PosStop = i
      GoTo GetGroup
    End If
  Next i
GetGroupName = ""
Exit Function
GetGroup:
  For i = 1 To PosStop
    If Tree.Nodes.Item(i).Image = 5 Then
      Group = Tree.Nodes.Item(i).Text
    End If
  Next i
GetGroupName = Group
End Function

Public Function RemoveUserNode(UserName As String, Tree As TreeView) As Long
Dim i As Integer
  For i = 1 To Tree.Nodes.Count
    If LCase(UserName) = Tree.Nodes.Item(i).Key Then
      Tree.Nodes.Remove i
      Exit Function
    End If
  Next i
End Function

Public Function SplitPackets(Data As String, Sock As Winsock, TypeMsg As Boolean)
Dim Packets() As String
Dim i As Integer
  If InStr(6, Data, "YMSG" & Chr(&H0)) > 0 Then
    Packets = Split(Data, "YMSG" & Chr(&H0))
      For i = 0 To UBound(Packets)
        If TypeMsg = True Then
          Call WebMessengerHandle("YMSG" & Chr(&H0) & Packets(i), Sock)
        Else
          Call YahooChat2Handle("YMSG" & Chr(&H0) & Packets(i), Sock)
        End If
DoEvents
      Next i
  Else
    If TypeMsg = True Then
      Call WebMessengerHandle(Data, Sock)
    Else
      Call YahooChat2Handle(Data, Sock)
    End If
  End If
End Function

Public Function RemoveItem(Item As String, List As ListView)
Dim i As Integer
  For i = 1 To List.ListItems.Count
    If LCase(Item) = LCase(List.ListItems(i).Key) Then
      List.ListItems.Remove i
      Exit Function
    End If
DoEvents
  Next i
End Function

Public Function ParseChatUsers(Data As String, List As ListView, Chat As RichTextBox)
On Error Resume Next
Dim Names() As String
Dim NickName As String
Dim ActualName As String
Dim Age As String
Dim Location As String
Dim Sex As String
Dim i As Integer
  Names = Split(Data, "À€109À€")
'109 - actual name
'110 - age
'141 - nick name
'142 - location
'113 - sex (male=33792,female=66560,empty=1024)
'31 30 39 C0 80 6A                                 109À€j
'61 79 5F 64 6F 67 5F 74-68 61 74 73 5F 77 68 61   ay_dog_thats_wha
'74 73 5F 75 70 C0 80 31-31 30 C0 80 31 39 C0 80   ts_upÀ€110À€19À€
'31 34 32 C0 80 4D 61 73-73 61 63 68 75 73 65 74   142À€Massachuset
'74 73 C0 80 31 31 33 C0-80 33 33 37 39 32 C0 80   tsÀ€113À€33792À€
    For i = 1 To UBound(Names)
        NickName = ""
        ActualName = ""
        Age = ""
        Location = ""
        Sex = ""
        Names(i) = "À€109À€" & Names(i)
      If InStr(1, Names(i), "À€109À€") > 0 Then
        ActualName = Split(Names(i), "À€109À€")(1)
        ActualName = LCase(Split(ActualName, "À€")(0))
      End If
      If InStr(1, Names(i), "À€110À€") > 0 Then
        Age = Split(Names(i), "À€110À€")(1)
        Age = Split(Age, "À€")(0)
      End If
      If Not Len(Age) > 0 Then Age = "Empty"
      If InStr(1, Names(i), "À€141À€") > 0 Then
        NickName = Split(Names(i), "À€141À€")(1)
        NickName = Split(NickName, "À€")(0)
      End If
      If Not Len(NickName) > 0 Then NickName = ActualName
      If InStr(1, Names(i), "À€142À€") > 0 Then
        Location = Split(Names(i), "À€142À€")(1)
        Location = Split(Location, "À€")(0)
      End If
      If Not Len(Location) > 0 Then Location = "Empty"
      If InStr(1, Names(i), "À€113À€") > 0 Then
        Sex = Split(Names(i), "À€113À€")(1)
        Sex = Split(Sex, "À€")(0)
      End If
      If Sex = "33792" Or 34048 Then
      'male
        Sex = "Male"
      End If
      If Sex = "66560" Then
      'female
        Sex = "Female"
      End If
      If Sex = "1024" Then
      'empty
        Sex = "Empty"
      End If
      If Age = "0" Then
        Age = "Empty"
      End If
        Call RemoveItem(ActualName, FrmChat.ListView)
        List.ListItems.Add(, ActualName, NickName, , 6).Tag = "Name: " & ActualName & " " & "Age: " & Age & " " & "Sex: " & Sex & " " & "Location: " & Location
          If InChat = True Then
          If FrmMessenger.mnuFilterJoinLeave.Checked = True Then GoTo ResumeLoop
          If LCase(ActualName) = LCase(YCurrentId) Then GoTo ResumeLoop
          If ActualName = NickName Then
            Chat.SelStart = Len(Chat.Text)
            Chat.SelText = ActualName & " Enters Room" & vbCrLf
            Chat.SelStart = Len(Chat.Text) - Len(ActualName & " Enters Room") - 2
            Chat.SelLength = Len(ActualName)
            Chat.SelColor = &HFF&
            Chat.SelItalic = True
            Chat.SelFontSize = 10
            Chat.SelStart = Len(Chat.Text)
          Else
            Chat.SelStart = Len(Chat.Text)
            Chat.SelText = NickName & " (" & ActualName & ") Enters Room" & vbCrLf
            Chat.SelStart = Len(Chat.Text) - Len(NickName & " (" & ActualName & ") Enters Room") - 2
            Chat.SelLength = Len(NickName & " (" & ActualName & ")")
            Chat.SelColor = &HFF&
            Chat.SelItalic = True
            Chat.SelFontSize = 10
            Chat.SelStart = Len(Chat.Text)
          End If
          End If
ResumeLoop:
    Next i
End Function

Public Function SetChatText(UserName As String, message As String, Chat As RichTextBox)
    Chat.SelStart = Len(Chat.Text)
    Chat.SelText = UserName & ": " & message & vbCrLf
    Chat.SelStart = Len(Chat.Text) - Len(UserName & ": " & message) - 2
    Chat.SelLength = Len(UserName) + 1
    Chat.SelColor = &HC0&
    Chat.SelBold = True
    Chat.SelFontSize = 10
    Chat.SelStart = Len(Chat.Text)
End Function

Public Function SetEmoteChatText(UserName As String, message As String, Chat As RichTextBox)
    Chat.SelStart = Len(Chat.Text)
    Chat.SelText = UserName & " " & message & vbCrLf
    Chat.SelStart = Len(Chat.Text) - Len(UserName & " " & message) - 2
    Chat.SelLength = Len(UserName) + 1
    Chat.SelColor = &HC0&
    Chat.SelBold = True
    Chat.SelFontSize = 10
    Chat.SelStart = Len(Chat.Text) - Len(message) - 2
    Chat.SelLength = Len(message)
    Chat.SelColor = &HC000C0
    Chat.SelStart = Len(Chat.Text)
End Function

Public Function SetThinkChatText(UserName As String, message As String, Chat As RichTextBox)
    Chat.SelStart = Len(Chat.Text)
    Chat.SelText = UserName & " " & message & vbCrLf
    Chat.SelStart = Len(Chat.Text) - Len(UserName & " " & message) - 2
    Chat.SelLength = Len(UserName)
    Chat.SelColor = &HC0&
    Chat.SelBold = True
    Chat.SelFontSize = 10
    Chat.SelStart = Len(Chat.Text) - Len(message) - 2
    Chat.SelLength = Len(message)
    Chat.SelColor = &HC000C0
    Chat.SelStart = Len(Chat.Text)
End Function


Public Function CheckIfPMOpen(UserName As String) As Boolean
Dim Frm As Form
  For Each Frm In Forms
    If LCase(Mid(Frm.Caption, 1, Len(UserName))) = LCase(UserName) Then
      CheckIfPMOpen = True
      Exit Function
    End If
  Next Frm
CheckIfPMOpen = False
End Function

Public Function SetNickName(UserName As String, List As ListView) As String
Dim i As Integer
  For i = 1 To List.ListItems.Count
    If LCase(UserName) = LCase(List.ListItems(i).Key) Then
      SetNickName = List.ListItems(i).Text
      Exit Function
    End If
  Next i
SetNickName = UserName
End Function

Public Function DenyPmNonBuddy(UserName As String, Menu As Menu, List As TreeView) As Boolean
Dim i As Integer
  If Menu.Checked = True Then
    For i = 1 To List.Nodes.Count
      If LCase(UserName) = LCase(List.Nodes(i).Key) Then
        DenyPmNonBuddy = True
        Exit Function
      End If
    Next i
      DenyPmNonBuddy = False
  Else
    DenyPmNonBuddy = True
  End If
End Function

Public Function CheckGroups(Group As String, Tree As TreeView) As Boolean
Dim i As Integer
  For i = 1 To Tree.Nodes.Count
    If LCase(Group) = LCase(Tree.Nodes(i).Key) Then
      CheckGroups = True
      Exit Function
    End If
  Next i
CheckGroups = False
End Function

Public Function MessageFilter(message As String) As String
Dim Pos(1 To 2) As Long
Dim RemoveMessage As String
ReCheck:
DoEvents:
  If InStrRev(message, ">") > 0 Then
    Pos(1) = InStrRev(message, ">")
  If InStrRev(message, "<", Pos(1)) > 0 Then
    Pos(2) = InStrRev(message, "<", Pos(1))
  If Pos(1) > Pos(2) Then
    RemoveMessage = Mid(message, Pos(2), Pos(1) - Pos(2) + 1)
    message = Replace(message, RemoveMessage, "")
    GoTo ReCheck
  End If
  End If
  End If
''
  If InStr(1, message, "<") > 0 Then
    Pos(1) = InStr(1, message, "<")
  If InStr(Pos(1), message, ">") > 0 Then
    Pos(2) = InStr(1, message, ">")
  If Pos(2) > Pos(1) Then
    RemoveMessage = Mid(message, Pos(1), Pos(2) - Pos(1) + 1)
    message = Replace(message, RemoveMessage, "")
    GoTo ReCheck
  End If
  End If
  End If
  If InStr(1, message, Chr(&H5B)) > 0 Then
    Pos(1) = InStr(1, message, Chr(&H5B))
  If InStr(Pos(1), message, Chr(&H6D)) > 0 Then
    Pos(2) = InStr(Pos(1), message, Chr(&H6D))
  If Pos(2) > Pos(1) Then
    RemoveMessage = Mid(message, Pos(1) - 1, Pos(2) - Pos(1) + 2)
    message = Replace(message, RemoveMessage, "")
    GoTo ReCheck
  End If
  End If
  End If
  If InStr(1, message, Chr(&H1B)) > 0 Then
    Pos(1) = InStr(1, message, Chr(&H1B))
  If InStr(Pos(1), message, Chr(&H6D)) > 0 Then
    Pos(2) = InStr(Pos(1), message, Chr(&H6D))
  If Pos(2) > Pos(1) Then
    RemoveMessage = Mid(message, Pos(1) - 1, Pos(2) - Pos(1) + 2)
    message = Replace(message, RemoveMessage, "")
    GoTo ReCheck
  End If
  End If
  End If
MessageFilter = message
End Function

Public Function EnableVoice(UserName As String, RoomName As String, VoiceToken As String, RoomSpace As String, Host As String, YAcs As YAcs)
YAcs.appInfo = "mc(6, 0, 0, 1643)&u=" & UserName & "&ia=us"
YAcs.UserName = UserName
YAcs.HostName = Host
YAcs.confKey = VoiceToken
YAcs.confName = "ch/" & RoomName & "::" & RoomSpace
YAcs.inputGain = 100
YAcs.outputGain = 100
YAcs.inputAGC = 100
YAcs.loadSound UserName
YAcs.createAndJoinConference
YAcs.joinConference
YAcs.inputSource = 100
End Function

Public Function ResetIcons(List As ListView)
On Error GoTo Leave
Dim i As Integer
  For i = 1 To List.ListItems.Count
    If List.ListItems(i).SmallIcon = 8 Then GoTo ResumeNext
      List.ListItems(i).SmallIcon = 6
ResumeNext:
DoEvents
  Next i
Leave:
End Function

Public Function SetIcon(UserName As String, Icon As Integer, List As ListView)
On Error GoTo Leave
Dim i As Integer
  For i = 1 To List.ListItems.Count
    If LCase(List.ListItems(i).Key) = LCase(UserName) Then
    If List.ListItems(i).SmallIcon = 8 Then Exit Function
      List.ListItems(i).SmallIcon = Icon
      Exit Function
    End If
  Next i
Leave:
End Function

Public Function SetIcon2(UserName As String, Icon As Integer, List As ListView)
On Error GoTo Leave
Dim i As Integer
  For i = 1 To List.ListItems.Count
    If LCase(List.ListItems(i).Key) = LCase(UserName) Then
      List.ListItems(i).SmallIcon = Icon
      Exit Function
    End If
  Next i
Leave:
End Function
Public Function CheckIfIgnored(UserName As String, List As ListView) As Boolean
On Error GoTo Leave
Dim i As Integer
  For i = 1 To List.ListItems.Count
    If LCase(List.ListItems(i).Key) = LCase(UserName) Then
    If List.ListItems(i).SmallIcon = 8 Then
      CheckIfIgnored = True
      Exit Function
    End If
    End If
  Next i
Leave:
CheckIfIgnored = False
End Function
