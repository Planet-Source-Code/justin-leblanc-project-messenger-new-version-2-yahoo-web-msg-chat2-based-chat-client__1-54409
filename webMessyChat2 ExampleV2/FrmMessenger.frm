VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmMessenger 
   Caption         =   "Project! Messenger"
   ClientHeight    =   4785
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   3960
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
   ScaleHeight     =   4785
   ScaleWidth      =   3960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList IMGLTree 
      Left            =   3240
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":0261
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":0494
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":05A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":06B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":07CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo CBBStatus 
      Height          =   345
      Left            =   0
      TabIndex        =   7
      Top             =   400
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Text            =   "Status"
      ImageList       =   "IMGLSmileys"
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   2415
      Left            =   0
      TabIndex        =   6
      Top             =   780
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4260
      _Version        =   393217
      Indentation     =   353
      Style           =   1
      ImageList       =   "IMGLTree"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   1920
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   21
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":08DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":0E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMessenger.frx":1508
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   741
      ButtonWidth     =   741
      ButtonHeight    =   688
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "adduser"
            Object.ToolTipText     =   "Add People"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pmuser"
            Object.ToolTipText     =   "Instant Message People"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "chat"
            Object.ToolTipText     =   "Join chat"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDSearch 
      Caption         =   "Y! Search"
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtSearch 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Frame FMHide 
      Height          =   3975
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   3135
      Begin MSComctlLib.ImageList IMGLSmileys 
         Left            =   2400
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMessenger.frx":1D9A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMessenger.frx":229C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmMessenger.frx":279E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image IMGLogin 
         Height          =   1455
         Left            =   720
         Picture         =   "FrmMessenger.frx":2CA0
         Top             =   480
         Width           =   1635
      End
      Begin VB.Label lblLogin 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login to Yahoo!"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   690
         Left            =   840
         MouseIcon       =   "FrmMessenger.frx":A92A
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   2040
         Width           =   1545
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4410
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuLogin 
      Caption         =   "&Login"
      Begin VB.Menu mnuSignin 
         Caption         =   "Sign In"
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCLose 
         Caption         =   "Close"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Visible         =   0   'False
      Begin VB.Menu mnuGroups 
         Caption         =   "Groups"
         Begin VB.Menu mnuExpand 
            Caption         =   "Expand All"
         End
         Begin VB.Menu mnuCollapse 
            Caption         =   "Collapse"
         End
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Send a message..."
      End
      Begin VB.Menu space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuManagefriends 
         Caption         =   "Manage friends list"
         Begin VB.Menu mnuAddfriend 
            Caption         =   "Add a friend"
         End
         Begin VB.Menu mnudeletefriend 
            Caption         =   "Delete a friend"
         End
      End
      Begin VB.Menu mnuViewfriends 
         Caption         =   "View friends"
         Begin VB.Menu mnuProfile 
            Caption         =   "Profile"
         End
         Begin VB.Menu mnuFiles 
            Caption         =   "Files"
         End
      End
      Begin VB.Menu space6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilters 
         Caption         =   "Filters"
         Begin VB.Menu mnublockusers 
            Caption         =   "Block Messages from"
            Begin VB.Menu mnuAll 
               Caption         =   "All"
            End
            Begin VB.Menu mnuNonfriends 
               Caption         =   "Non Friends"
            End
         End
         Begin VB.Menu mnuBlockAdd 
            Caption         =   "Block Add requests"
         End
         Begin VB.Menu space7 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFilterJoinLeave 
            Caption         =   "Filter Join And Leave Chat Messages"
         End
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "Yahoo! Chat"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuYahooHelp 
         Caption         =   "Yahoo! Help Center"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuUser 
      Caption         =   "User"
      Visible         =   0   'False
      Begin VB.Menu mnuSendamessage 
         Caption         =   "Send a message..."
      End
      Begin VB.Menu space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProfilesa 
         Caption         =   "Profile"
      End
      Begin VB.Menu mnuFilesa 
         Caption         =   "Files"
      End
      Begin VB.Menu space5 
         Caption         =   "-"
      End
      Begin VB.Menu mnudeleteafriend 
         Caption         =   "Delete friend"
      End
   End
End
Attribute VB_Name = "FrmMessenger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CBBStatus_Click()
Select Case CBBStatus.SelectedItem.Index
  Case 1
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgVisible(SessionKey(1))
  Case 2
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgInvisible(SessionKey(1))
End Select
End Sub

Private Sub CMDSearch_Click()
Dim Url As String
If txtSearch.Text = "" Then Exit Sub
  Url = Replace(txtSearch.Text, Chr(&H20), "%20")
  txtSearch.Text = ""
  Call ShellExecute(&O0, "Open", "http://search.yahoo.com/search?p=" & Url & "&fr=msgr-buddy&ei=UTF-8", vbNullString, vbNullString, vbNormal)
End Sub

Private Sub FMHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogin.FontUnderline = False
End Sub

Private Sub Form_Load()
Call Form_Resize
Call LoadStatus(CBBStatus)
End Sub

Private Sub Form_Resize()
On Error Resume Next
  If Me.Width < 2600 Then Me.Width = 2600
  If Me.Height < 3700 Then Me.Height = 3700
'frame and label login and image login
    FMHide.Width = Me.Width - 330
    FMHide.Height = Me.Height - 1250
    '
    lblLogin.Top = FMHide.Height / 2 + 600
    lblLogin.Left = 0
    lblLogin.Width = FMHide.Width
    '
    IMGLogin.Top = FMHide.Height / 2 - 1000
    IMGLogin.Left = FMHide.Width / 2 - 800
    'status box and toolbox
    CBBStatus.Width = Me.Width - 100
    '
    TreeView.Width = Me.Width - 110
    TreeView.Height = Me.Height - 1980
'search box
txtSearch.Top = Me.Height - txtSearch.Height - 820
txtSearch.Width = Me.Width - CMDSearch.Width - 450
'search button
CMDSearch.Top = Me.Height - txtSearch.Height - 800
CMDSearch.Left = Me.Width - CMDSearch.Width - 350

End Sub

Private Sub Form_Unload(Cancel As Integer)
With FrmLogin
  .SockAuthentication.Close
  .SockYahooChat2.Close
  .SockChat.Close
End With
End
End Sub

Private Sub lblLogin_Click()
'If Mid(lblLogin.Caption, 1, 24) = "Connecting to Yahoo! as " Then Exit Sub
With FrmLogin
  YCurrentId = LCase(.txtName.Text)
  YPass = .txtPassword.Text
  '
  lblLogin.Caption = "Connecting to Yahoo! as " & YCurrentId
  '
  .Visible = False
  Call SockConnect(YServer(1), YPort(1), .SockAuthentication)
End With
End Sub

Private Sub lblLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblLogin.FontUnderline = True
End Sub

Private Sub mnuAbout_Click()
FrmAbout.Show
End Sub

Private Sub mnuAddfriend_Click()
AddUserForm = True
If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
End Sub

Private Sub mnuAll_Click()
If mnuAll.Checked = True Then
  mnuAll.Checked = False
Else
  mnuAll.Checked = True
  mnuNonfriends.Checked = False
End If
End Sub

Private Sub mnuBlockAdd_Click()
If mnuBlockAdd.Checked = True Then
  mnuBlockAdd.Checked = False
Else
  mnuBlockAdd.Checked = True
End If
End Sub

Private Sub mnuChat_Click()
FrmChat.Visible = True
End Sub

Private Sub mnuCLose_Click()
Unload Me
End Sub

Private Sub mnuCollapse_Click()
Call ExpandNodes(TreeView, False)
End Sub

Private Sub mnudeleteafriend_Click()
RemoveName = LCase(TreeView.SelectedItem.Text)
AddUserForm = False
If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
End Sub

Private Sub mnuDeletefriend_Click()
Dim MsgC As String
  MsgC = InputBox("Enter persons Name to Remove", "Remove User")
    If Len(MsgC) > 0 Then
      RemoveName = LCase(MsgC)
      AddUserForm = False
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
    End If
End Sub

Private Sub mnuExpand_Click()
Call ExpandNodes(TreeView, True)
End Sub

Private Sub mnuFiles_Click()
Dim MsgC As String
  MsgC = InputBox("Enter persons Name", "View Files")
    If Len(MsgC) > 0 Then
      Call ShellExecute(&O0, "Open", "http://f2.up.briefcase.yahoo.com/edit/" & LCase(MsgC) & "/reg?.done=http%3a//briefcase.yahoo.com/bc/" & LCase(MsgC), vbNullString, vbNullString, vbNormal)
    End If
End Sub

Private Sub mnuFilesa_Click()
Call ShellExecute(&O0, "Open", "http://f2.up.briefcase.yahoo.com/edit/" & LCase(TreeView.SelectedItem.Text) & "/reg?.done=http%3a//briefcase.yahoo.com/bc/" & LCase(TreeView.SelectedItem.Text), vbNullString, vbNullString, vbNormal)
End Sub

Private Sub mnuFilterJoinLeave_Click()
If mnuFilterJoinLeave.Checked = True Then
  mnuFilterJoinLeave.Checked = False
Else
  mnuFilterJoinLeave.Checked = True
End If
End Sub

Private Sub mnuNonfriends_Click()
If mnuNonfriends.Checked = True Then
  mnuNonfriends.Checked = False
Else
  mnuNonfriends.Checked = True
  mnuAll.Checked = False
End If
End Sub

Private Sub mnuProfile_Click()
Dim MsgC As String
  MsgC = InputBox("Enter persons Profile to view", "View Profile")
    If Len(MsgC) > 0 Then
      Call ShellExecute(&O0, "Open", "http://profiles.yahoo.com/" & LCase(MsgC), vbNullString, vbNullString, vbNormal)
    End If
End Sub

Private Sub mnuProfilesa_Click()
Call ShellExecute(&O0, "Open", "http://profiles.yahoo.com/" & LCase(TreeView.SelectedItem.Text), vbNullString, vbNullString, vbNormal)
End Sub

Private Sub mnuRefresh_Click()
Call lblLogin_Click
End Sub

Private Sub mnuSendamessage_Click()
Dim NewPm As New FrmPm
  NewPm.Caption = LCase(TreeView.SelectedItem.Text) & " ~ Instant Message"
  NewPm.txtWho = LCase(WhoFrom)
  NewPm.txtWho.Locked = True
  NewPm.Show
End Sub

Private Sub mnuSendMessage_Click()
Dim NewPm As New FrmPm
  NewPm.Show
End Sub

Private Sub mnuSignin_Click()
FrmLogin.Visible = True
End Sub

Private Sub mnuYahooHelp_Click()
Call ShellExecute(&O0, "Open", "http://help.yahoo.com/help/us/edit/", vbNullString, vbNullString, vbNormal)
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim NewPm As New FrmPm
Select Case Button.Key
  Case "adduser"
    AddUserForm = True
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
  Case "pmuser"
    NewPm.Show
  Case "chat"
    FrmChat.Visible = True
End Select
End Sub

Private Sub TreeView_Collapse(ByVal Node As MSComctlLib.Node)
Node.Image = 4
End Sub

Private Sub TreeView_DblClick()
On Error GoTo Leave
Dim NewPm As New FrmPm
If TreeView.SelectedItem.Image = 1 Or TreeView.SelectedItem.Image = 2 Then
  NewPm.Caption = LCase(TreeView.SelectedItem.Text) & " ~ Instant Message"
  NewPm.txtWho = LCase(TreeView.SelectedItem.Text)
  NewPm.txtWho.Locked = True
  NewPm.Show
End If
Leave:
End Sub

Private Sub TreeView_Expand(ByVal Node As MSComctlLib.Node)
Node.Image = 5
End Sub

Private Sub TreeView_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Leave
If Button = 2 And TreeView.SelectedItem.Image = 1 Or TreeView.SelectedItem.Image = 2 Then
  Call PopupMenu(mnuUser)
End If
Leave:
End Sub
