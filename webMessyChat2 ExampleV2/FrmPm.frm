VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmPm 
   Caption         =   " ~ Instant Message"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4905
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton CMDSend 
      Caption         =   "Send"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   615
      Left            =   3960
      TabIndex        =   4
      Top             =   2280
      Width           =   615
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   2985
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox txtText 
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   2280
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1085
      _Version        =   393217
      MultiLine       =   0   'False
      TextRTF         =   $"FrmPm.frx":0000
   End
   Begin RichTextLib.RichTextBox RTBPm 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   3201
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"FrmPm.frx":0077
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "IMGPm"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add Person"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "ignore"
            Object.ToolTipText     =   "Ignore Person"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "profile"
            Object.ToolTipText     =   "View Profile"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin VB.TextBox txtWho 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   0
         Width           =   3615
      End
      Begin MSComctlLib.ImageList IMGPm 
         Left            =   3600
         Top             =   0
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
               Picture         =   "FrmPm.frx":00EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPm.frx":0640
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmPm.frx":0B92
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
   End
End
Attribute VB_Name = "FrmPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDSend_Click()
'Dim MsgC As String
If Len(txtWho.Text) = 0 Then Exit Sub
If txtWho.Locked = False Then
  txtWho.Text = LCase(txtWho.Text)
  txtWho.Locked = True
  Me.Caption = txtWho.Text & " ~ Instant Message"
If CheckIfPMOpen(txtWho.Text) = True Then Unload Me
End If
'I disabled this feature because when using chat2 proto with web messenger proto it allows you to still
'send pms even when the user is no on your buddy list
'If GetNodeIndex(LCase(txtWho.Text), FrmMessenger.TreeView) = 0 Then
'  MsgC = MsgBox("To have the ability to send messages to this user[" & LCase(txtWho.Text) & "] you MUST add them or your message will not be received. If you want to add this user select 'Yes'.", vbYesNo, "User Not In List")
'    Select Case MsgC
'      Case vbYes
'        AddName = LCase(txtWho.Text)
'        If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
'        Exit Sub
'      Case vbNo
'        Exit Sub
'    End Select
'End If
'If FrmLogin.SockYahooChat2.State = sckConnected Then
  'FrmLogin.SockYahooChat2.SendData WebMsgPmSend(YCurrentId, txtWho.Text, txtText.Text, SessionKey(1))
If FrmLogin.SockChat.State = sckConnected Then
  FrmLogin.SockChat.SendData Chat2PmSend(YCurrentId, txtWho.Text, txtText.Text, SessionKey(2))
  RTBPm.SelStart = Len(RTBPm.Text)
  RTBPm.SelText = YCurrentId & ": " & txtText.Text & vbCrLf
  RTBPm.SelStart = Len(RTBPm.Text) - Len(YCurrentId & ": " & txtText.Text) - 2
  RTBPm.SelLength = Len(YCurrentId) + 1
  RTBPm.SelColor = &HFF0000
  RTBPm.SelBold = True
  RTBPm.SelFontSize = 10
  RTBPm.SelStart = Len(RTBPm.Text)
End If
  txtText.Text = ""
End Sub

Private Sub Form_Resize()
On Error Resume Next
If Me.Width < 2500 Then Me.Width = 2500
If Me.Height < 2700 Then Me.Height = 2700
  RTBPm.Width = Me.Width - 140
  RTBPm.Height = Me.Height - 2000
  '
  txtText.Top = RTBPm.Top + RTBPm.Height + 200
  txtText.Width = Me.Width - 940
  '
  CMDSend.Top = txtText.Top
  CMDSend.Left = txtText.Width + 100
  '
  txtWho.Width = Me.Width - 1200
End Sub

Private Sub mnuClear_Click()
RTBPm.Text = ""
End Sub

Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText RTBPm.SelText
End Sub

Private Sub mnuSelectall_Click()
RTBPm.SelStart = 0
RTBPm.SelLength = Len(RTBPm.Text)
End Sub

Private Sub RTBPm_Change()
If Len(RTBPm.Text) > 3000 Then
  RTBPm.SelStart = 0
  RTBPm.SelLength = 300
  RTBPm.SelText = ""
  RTBPm.SelStart = Len(RTBPm.Text)
End If
End Sub

Private Sub RTBPm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
  Call PopupMenu(mnuMenu)
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "add"
    AddName = LCase(txtWho.Text)
    AddUserForm = True
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
  Case "ignore"
  
  Case "profile"
    If Len(Me.txtWho.Text) = 0 Then Exit Sub
    Call ShellExecute(&O0, "Open", "http://profiles.yahoo.com/" & LCase(txtWho.Text), vbNullString, vbNullString, vbNormal)
End Select
End Sub

Private Sub txtText_Change()
If Len(txtText.Text) = 0 Then
  CMDSend.Enabled = False
Else
  CMDSend.Enabled = True
End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Call CMDSend_Click
End If
End Sub
