VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{7D1E9C3C-BD6A-11D3-87A8-009027A35D73}#1.0#0"; "yacsui.dll"
Object = "{2B323CCC-50E3-11D3-9466-00A0C9700498}#1.0#0"; "yacscom.dll"
Begin VB.Form FrmChat 
   Caption         =   " ~ Chat"
   ClientHeight    =   4695
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6120
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.PictureBox VoiceBar 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   6015
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   6015
      Begin YACSUILibCtl.YSlider YSlider 
         Height          =   255
         Left            =   5160
         OleObjectBlob   =   "FrmChat.frx":0000
         TabIndex        =   13
         Top             =   75
         Width           =   735
      End
      Begin VB.CommandButton CMDTalk 
         Caption         =   "Talk"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   70
         Width           =   735
      End
      Begin VB.Image IMGVoice 
         Height          =   360
         Index           =   0
         Left            =   0
         MouseIcon       =   "FrmChat.frx":0024
         MousePointer    =   99  'Custom
         Picture         =   "FrmChat.frx":0176
         ToolTipText     =   "Disable Voice"
         Top             =   10
         Width           =   360
      End
      Begin VB.Label lblTalkingUser 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   14
         Top             =   105
         Width           =   1815
      End
      Begin YACSUILibCtl.YVuMeter YVuMeter 
         Height          =   255
         Index           =   1
         Left            =   2520
         OleObjectBlob   =   "FrmChat.frx":0386
         TabIndex        =   12
         Top             =   75
         Width           =   735
      End
      Begin YACSUILibCtl.YVuMeter YVuMeter 
         Height          =   255
         Index           =   0
         Left            =   1800
         OleObjectBlob   =   "FrmChat.frx":03AA
         TabIndex        =   11
         Top             =   75
         Width           =   735
      End
      Begin VB.Label lblUserVoiceSpecs 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0/0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   10
         Top             =   105
         Width           =   735
      End
   End
   Begin RichTextLib.RichTextBox RTBText 
      Height          =   615
      Left            =   0
      TabIndex        =   3
      Top             =   3720
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   1085
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"FrmChat.frx":03CE
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "IMGPm"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add Person"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ignore"
            Object.ToolTipText     =   "Ignore Person"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "profile"
            Object.ToolTipText     =   "View Profile"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pm"
            Object.ToolTipText     =   "Instant Message Person"
            ImageIndex      =   5
         EndProperty
      EndProperty
      Begin VB.ComboBox CBBRoom 
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Text            =   "QQ's Chat Room:1"
         Top             =   0
         Width           =   3735
      End
      Begin VB.CommandButton CMDJoin 
         Caption         =   "Join"
         Height          =   330
         Left            =   5280
         TabIndex        =   6
         Top             =   5
         Width           =   615
      End
      Begin MSComctlLib.ImageList IMGPm 
         Left            =   5040
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":0445
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":0997
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":0EE9
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":143B
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":166E
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":1D48
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":209A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmChat.frx":23EC
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton CMDSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   3720
      Width           =   615
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   3015
      Left            =   4200
      TabIndex        =   2
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "IMGPm"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Users: 0"
         Object.Width           =   2646
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTBChat 
      Height          =   3015
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   5318
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmChat.frx":273E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4440
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Voice Status: Voice Disabled"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin YACSCOMLibCtl.YAcs YAcs 
      Left            =   960
      OleObjectBlob   =   "FrmChat.frx":27B5
      Top             =   3480
   End
   Begin VB.Image IMGVoice 
      Height          =   360
      Index           =   1
      Left            =   0
      MouseIcon       =   "FrmChat.frx":27D9
      MousePointer    =   99  'Custom
      Picture         =   "FrmChat.frx":292B
      ToolTipText     =   "Enable Voice"
      Top             =   3360
      Width           =   360
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuMessage 
         Caption         =   "Send a message..."
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Person"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignore Person"
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfile 
         Caption         =   "Profile"
      End
      Begin VB.Menu mnuFiles 
         Caption         =   "Files"
      End
   End
   Begin VB.Menu mnuMenu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuSelectall 
         Caption         =   "Select All"
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "FrmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'yacscom.dll
'yacsui.dll
Private Sub CMDJoin_Click()
If LCase(ChatRoom) = LCase(CBBRoom.Text) Then Exit Sub
InChat = False
ListView.ListItems.Clear
If FrmLogin.SockChat.State = sckConnected Then FrmLogin.SockChat.SendData JoinRoom(YCurrentId, CBBRoom.Text, SessionKey(2))
End Sub

Private Sub CMDSend_Click()
If FrmLogin.SockChat.State = sckConnected Then
  RTBChat.SelStart = Len(RTBChat.Text)
  RTBChat.SelText = SetNickName(YCurrentId, FrmChat.ListView) & ": " & RTBText.Text & vbCrLf
  RTBChat.SelStart = Len(RTBChat.Text) - Len(SetNickName(YCurrentId, FrmChat.ListView) & ": " & RTBText.Text) - 2
  RTBChat.SelLength = Len(SetNickName(YCurrentId, FrmChat.ListView)) + 1
  RTBChat.SelColor = &HFF0000
  RTBChat.SelBold = True
  RTBChat.SelFontSize = 10
  RTBChat.SelStart = Len(RTBChat.Text)
  FrmLogin.SockChat.SendData Chat2ChatSend(YCurrentId, ChatRoom, RTBText.Text, SessionKey(2))
  RTBText.Text = ""
End If
End Sub

Private Sub CMDTalk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  YAcs.startTransmit
End If
End Sub

Private Sub CMDTalk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
  YAcs.stopTransmit
End If
End Sub

Private Sub Form_Load()
YVuMeter(0).direction = 1
YVuMeter(1).highlight = 1
YVuMeter(1).minValue = -40
YVuMeter(1).maxValue = 0
YVuMeter(1).Value = -40
YVuMeter(0).highlight = 1
YVuMeter(0).minValue = -40
YVuMeter(0).maxValue = 0
YVuMeter(0).Value = -40
Call Form_Resize
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 4900 Then Me.Width = 4900
If Me.Height < 4200 Then Me.Height = 4200
'room
CBBRoom.Width = Me.Width - 2250
'button room
CMDJoin.Left = Me.Width - CMDJoin.Width - 150
'chat window
RTBChat.Width = Me.Width - 2000
RTBChat.Height = Me.Height - 2150
'voice bar
VoiceBar.Top = RTBChat.Height + VoiceBar.Height
VoiceBar.Width = Me.Width - 100
'
IMGVoice(1).Top = RTBChat.Height + IMGVoice(1).Height + 30
''
YSlider.Left = VoiceBar.Width - YSlider.Width - 100
lblTalkingUser.Width = VoiceBar.Width - 4100
'
ListView.Left = Me.Width - ListView.Width - 150
ListView.Height = Me.Height - 2150
'users
lblUsers.Left = Me.Width - ListView.Width - 150
'message box
RTBText.Width = Me.Width - 900
RTBText.Top = RTBChat.Height + RTBText.Height + 150
'button send
CMDSend.Top = RTBChat.Height + CMDSend.Height + 150
CMDSend.Left = Me.Width - CMDSend.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
If VoiceBar.Visible = True Then
  Call IMGVoice_Click(0)
End If
  Cancel = True
  InChat = False
  ChatRoom = ""
  ListView.ListItems.Clear
If FrmLogin.SockChat.State = sckConnected Then FrmLogin.SockChat.SendData LogOutChat2(YCurrentId, ChatRoom, SessionKey(2))
  Me.Visible = False
End Sub

Private Sub IMGVoice_Click(Index As Integer)
Select Case Index
  Case 0
    IMGVoice(1).Visible = True
    VoiceBar.Visible = False
    '
    YAcs.leaveConference
  Case 1
    IMGVoice(1).Visible = False
    VoiceBar.Visible = True
    '
    Call EnableVoice(YCurrentId, ChatRoom, YVoiceToken, YRoomSpace, "v9.vc.dcn.yahoo.com", YAcs)
End Select
End Sub

Private Sub ListView_Click()
On Error GoTo Leave
  ListView.ToolTipText = ListView.SelectedItem.Tag
Leave:
End Sub

Private Sub ListView_DblClick()
On Error GoTo Leave
Dim NewPm As New FrmPm
If CheckIfPMOpen(ListView.SelectedItem.Key) = True Then Exit Sub
  NewPm.Caption = LCase(ListView.SelectedItem.Key) & " ~ Instant Message"
  NewPm.txtWho = LCase(ListView.SelectedItem.Key)
  NewPm.txtWho.Locked = True
  NewPm.Show
Leave:
End Sub

Private Sub ListView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Leave
If Button = 2 Then
If CheckIfIgnored(ListView.SelectedItem.Key, ListView) = True Then
  mnuIgnore.Caption = "UnIgnore Person"
Else
  mnuIgnore.Caption = "Ignore Person"
End If
  Call PopupMenu(mnuMenu)
End If
Leave:
End Sub

Private Sub ListView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Leave
  ListView.ToolTipText = ListView.SelectedItem.Tag
Leave:
End Sub

Private Sub mnuAdd_Click()
On Error GoTo Leave
AddName = LCase(ListView.SelectedItem.Key)
AddUserForm = True
  If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
Leave:
End Sub

Private Sub mnuClear_Click()
RTBChat.Text = ""
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText RTBChat.SelText
End Sub

Private Sub mnuFiles_Click()
On Error GoTo Leave
  Call ShellExecute(&O0, "Open", "http://f2.up.briefcase.yahoo.com/edit/" & LCase(ListView.SelectedItem.Key) & "/reg?.done=http%3a//briefcase.yahoo.com/bc/" & LCase(ListView.SelectedItem.Key), vbNullString, vbNullString, vbNormal)
Leave:
End Sub

Private Sub mnuIgnore_Click()
If mnuIgnore.Caption = "Ignore Person" Then
  YAcs.muteSource 0, ListView.SelectedItem.Key
  Call SetIcon2(ListView.SelectedItem.Key, 8, ListView)
Else
  YAcs.muteSource 1, ListView.SelectedItem.Key
  Call SetIcon2(ListView.SelectedItem.Key, 6, ListView)
End If
End Sub

Private Sub mnuMessage_Click()
On Error GoTo Leave
Dim NewPm As New FrmPm
If CheckIfPMOpen(ListView.SelectedItem.Key) = True Then Exit Sub
  NewPm.Caption = LCase(ListView.SelectedItem.Key) & " ~ Instant Message"
  NewPm.txtWho = LCase(ListView.SelectedItem.Key)
  NewPm.txtWho.Locked = True
  NewPm.Show
Leave:
End Sub

Private Sub mnuProfile_Click()
On Error GoTo Leave
  Call ShellExecute(&O0, "Open", "http://profiles.yahoo.com/" & LCase(ListView.SelectedItem.Key), vbNullString, vbNullString, vbNormal)
Leave:
End Sub

Private Sub mnuSelectall_Click()
RTBChat.SelStart = 0
RTBChat.SelLength = Len(RTBChat.Text)
End Sub

Private Sub RTBChat_Change()
'the rtb control can't handle over a certain amount of data so we'll take away so text when
'it's length exceeds a cetrain amount
If Len(RTBChat.Text) > 3000 Then
  RTBChat.SelStart = 0
  RTBChat.SelLength = 300
  RTBChat.SelText = ""
  RTBChat.SelStart = Len(RTBChat.Text)
End If
End Sub

Private Sub RTBChat_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
  Call PopupMenu(mnuMenu2)
End If
End Sub

Private Sub RTBText_Change()
If Len(RTBText.Text) = 0 Then
  CMDSend.Enabled = False
Else
  CMDSend.Enabled = True
End If
End Sub

Private Sub RTBText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CMDSend_Click
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Leave
Dim MsgC As String
Dim NewPm As New FrmPm
Select Case Button.Key
  Case "add"
    AddUserForm = True
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
  Case "ignore"
    YAcs.muteSource 0, ListView.SelectedItem.Key
    Call SetIcon2(ListView.SelectedItem.Key, 8, ListView)
  Case "profile"
    MsgC = InputBox("Enter persons Profile to view", "View Profile")
      If Len(MsgC) > 0 Then
        Call ShellExecute(&O0, "Open", "http://profiles.yahoo.com/" & LCase(MsgC), vbNullString, vbNullString, vbNormal)
      End If
  Case "pm"
    NewPm.Show
End Select
Leave:
End Sub

Private Sub YAcs_onAudioError(ByVal code As Long, ByVal message As String)
StatusBar.SimpleText = "Voice Status: Audio Error (" & message & ")"
End Sub

Private Sub YAcs_onConferenceNotReady()
StatusBar.SimpleText = "Voice Status: Conference Not Ready"
End Sub

Private Sub YAcs_onConferenceReady()
StatusBar.SimpleText = "Voice Status: Voice Enabled"
End Sub

Private Sub YAcs_onInputLevelChange(ByVal level As Integer)
YVuMeter(0).Value = level
End Sub

Private Sub YAcs_onOutputLevelChange(ByVal level As Integer)
YVuMeter(1).Value = level
End Sub

Private Sub YAcs_onRemoteSourceOffAir(ByVal sourceId As Long, ByVal sourceName As String)
lblTalkingUser.Caption = ""
lblUserVoiceSpecs.ForeColor = &H0&
End Sub

Private Sub YAcs_onRemoteSourceOnAir(ByVal sourceId As Long, ByVal sourceName As String)
lblTalkingUser.Caption = SetNickName(sourceName, ListView)
Call SetIcon(sourceName, 7, ListView)
End Sub

Private Sub YAcs_onSourceEntry(ByVal sourceId As Long, ByVal sourceName As String)
Call SetIcon(sourceName, 7, ListView)
If CheckIfIgnored(sourceName, ListView) = True Then
  YAcs.muteSource 0, sourceName
End If
End Sub

Private Sub YAcs_onSourceExit(ByVal sourceId As Long, ByVal sourceName As String)
Call SetIcon(sourceName, 6, ListView)
End Sub

Private Sub YAcs_onSystemConnect()
StatusBar.SimpleText = "Voice Status: System Connect"
End Sub

Private Sub YAcs_onSystemConnectFailure(ByVal code As Long, ByVal message As String)
StatusBar.SimpleText = "Voice Status: System Connect Failure"
End Sub

Private Sub YAcs_onSystemDisconnect()
StatusBar.SimpleText = "Voice Status: System Disconnect"
Call ResetIcons(ListView)
End Sub

Private Sub YAcs_onTransmitReport(ByVal numReceiving As Long, ByVal numTotal As Long)
If numReceiving = numTotal Then lblUserVoiceSpecs.ForeColor = &HC000&
If numReceiving < numTotal And numReceiving > 0 Then lblUserVoiceSpecs.ForeColor = &HC0C0&
If numReceiving = 0 Then lblUserVoiceSpecs.ForeColor = &HC0&
lblUserVoiceSpecs.Caption = numReceiving & "/" & numTotal
End Sub

Private Sub YSlider_onValueChanged(ByVal newVal As Integer)
YAcs.outputGain = YSlider.Value
End Sub
