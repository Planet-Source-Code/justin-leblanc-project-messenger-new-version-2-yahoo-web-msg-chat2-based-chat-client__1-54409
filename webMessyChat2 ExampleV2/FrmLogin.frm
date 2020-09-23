VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sign In"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3135
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock SockChat 
      Left            =   1080
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock SockAuthentication 
      Left            =   120
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Help"
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   11
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Sign In"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   3360
      Width           =   855
   End
   Begin VB.Frame FmAlreadyHaveYId 
      Caption         =   "Already have a Yahoo! ID?"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   2895
      Begin VB.CheckBox CKBGeneral 
         Caption         =   "Sign in as invisible"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox CKBGeneral 
         Caption         =   "Automatically sign in"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox CKBGeneral 
         Caption         =   "Remember My ID/Password"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Label lblGeneral 
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblGeneral 
         BackStyle       =   0  'Transparent
         Caption         =   "Yahoo! ID:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame FmNewUser 
      Caption         =   "New User?"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin MSWinsockLib.Winsock SockYahooChat2 
         Left            =   480
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.CommandButton CMDGetId 
         Caption         =   "Get a Yahoo! ID"
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is the start of a Yahoo! Messenger Client based on Yahoo!s "Web Messenger Protocol" and Yahoo!s "Chat 2.0" Protocol
'I threw together a rough draft for learning purposes of basic TCP protocol transactions and to teach how users can create their
'Yahoo! based client to chat with. Later I'll add a simple chat room interface to it.
'
'**Project V2
'more options like: chat, block non friends,block all, block add requests, chat voice,afew more chat2 packets like chatsend,emotechatsend, thinkchatsend and more!
'for voice I used these 2 .dll's
'yacscom.dll
'yacsui.dll
'if you dont have them and need them goto http://www.yahoo.com and download the new yahoo messenger
'
'By: Justin LeBlanc
'Email: visual_basic_software@yahoo.com
'
Private Sub CMDGeneral_Click(Index As Integer)
Select Case Index
  Case 0
    If CMDGeneral(0).Caption = "Sign In" Then
      YCurrentId = LCase(txtName.Text)
      YPass = txtPassword.Text
      '
      FrmMessenger.lblLogin.Caption = "Connecting to Yahoo! as " & YCurrentId
      '
      Me.Visible = False
      Call SockConnect(YServer(1), YPort(1), SockAuthentication)
    Else
      Call LogOutMessenger
    End If
  Case 1
    Me.Visible = False
  Case 2
    Call ShellExecute(&O0, "Open", "http://help.yahoo.com/help/us/edit/", vbNullString, vbNullString, vbNormal)
End Select
End Sub

Private Sub CMDGetId_Click()
Call ShellExecute(&O0, "Open", "http://edit.my.yahoo.com/config/eval_register?.src=chat&.lg=us&.intl=us&.done=http%3a//chat.yahoo.com", vbNullString, vbNullString, vbNormal)
End Sub

Private Sub Form_Load()
YBuild(1) = Chr(&H0) & Chr(&HA)
YBuild(2) = Chr(&H0) & Chr(&H65)
YServer(1) = "mail.yahoo.com"
YPort(1) = "80"
YServer(2) = "wcs1.msg.dcn.yahoo.com"
YPort(2) = "5050"
YServer(3) = "dcs1.chat.dcn.yahoo.com"
YPort(3) = "5050"
FrmMessenger.Show
End Sub

Private Sub SockAuthentication_Connect()
SockAuthentication.SendData Authentication(YCurrentId, YPass)
End Sub

Private Sub SockAuthentication_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
  SockAuthentication.GetData Buffer
Debug.Print "Authentication Socket: " & Buffer
  Call AuthenticationHandle(Buffer, YCookie, SockAuthentication, SockYahooChat2)
End Sub

Private Sub SockAuthentication_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print "Authentication Socket: An Error Occured: " & Description & " [" & Number & "]"
  Call LogOutMessenger
End Sub

Private Sub SockChat_Connect()
SockChat.SendData LoginChat2(YCurrentId, YCookie)
End Sub

Private Sub SockChat_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
  SockChat.GetData Buffer
Debug.Print "Chat Socket: " & Buffer
  Call SplitPackets(Buffer, SockChat, False)
End Sub

Private Sub SockChat_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print "Chat Socket: An Error Occured: " & Description & " [" & Number & "]"
  Call LogOutMessenger
If Number = 10049 Then
  Select Case YServer(3)
    Case "dcs3.chat.dcn.yahoo.com"
      YServer(3) = "dcs1.chat.dcn.yahoo.com"
      Call CMDGeneral_Click(0)
    Case "dcs2.chat.dcn.yahoo.com"
      YServer(3) = "dcs3.chat.dcn.yahoo.com"
      Call CMDGeneral_Click(0)
    Case "dcs1.chat.dcn.yahoo.com"
      YServer(3) = "dcs2.chat.dcn.yahoo.com"
      Call CMDGeneral_Click(0)
  End Select
ElseIf Number = 10053 Then
  Call LogOutMessenger
End If
End Sub

Private Sub SockYahooChat2_Connect()
SockYahooChat2.SendData LoginWebMessenger(YCurrentId, YCookie)
End Sub

Private Sub SockYahooChat2_DataArrival(ByVal bytesTotal As Long)
Dim Buffer As String
  SockYahooChat2.GetData Buffer
Debug.Print "Messenger Socket: " & Buffer
  Call SplitPackets(Buffer, FrmLogin.SockYahooChat2, True)
End Sub

Private Sub SockYahooChat2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Debug.Print "Messenger Socket: An Error Occured: " & Description & " [" & Number & "]"
  Call LogOutMessenger
If Number = 10049 Then
  Select Case YServer(2)
    Case "wcs3.msg.dcn.yahoo.com"
      YServer(2) = "wcs1.msg.dcn.yahoo.com"
      Call CMDGeneral_Click(0)
    Case "wcs2.msg.dcn.yahoo.com"
      YServer(2) = "wcs3.msg.dcn.yahoo.com"
      Call CMDGeneral_Click(0)
    Case "wcs1.msg.dcn.yahoo.com"
      YServer(2) = "wcs2.msg.dcn.yahoo.com"
      Call CMDGeneral_Click(0)
  End Select
ElseIf Number = 10053 Then
  Call LogOutMessenger
End If
End Sub
