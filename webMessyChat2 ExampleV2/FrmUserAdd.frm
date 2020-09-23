VERSION 5.00
Begin VB.Form FrmUserAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Request"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5295
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Deny"
      Height          =   375
      Index           =   3
      Left            =   4320
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Accept And Add"
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   5
      Top             =   1680
      Width           =   1575
   End
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Accept"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   4
      Top             =   1680
      Width           =   855
   End
   Begin VB.TextBox txtWho 
      Appearance      =   0  'Flat
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox txtMessage 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   960
      Width           =   4455
   End
   Begin VB.CommandButton CMDGeneral 
      Caption         =   "Profile"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "The Following is the Users message to you."
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
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4575
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "The Following User would like to add you to his/her buddylist."
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "FrmUserAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDGeneral_Click(Index As Integer)
Select Case Index
  Case 0
    Call ShellExecute(&O0, "Open", "http://profiles.yahoo.com/" & LCase(txtWho.Text), vbNullString, vbNullString, vbNormal)
  Case 1
    Unload Me
  Case 2
    AddName = LCase(txtWho.Text)
    AddUserForm = True
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddPrompt(SessionKey(1))
    Unload Me
  Case 3
    If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgDenyAdd(YCurrentId, txtWho.Text, "Thanks, but no thanks.", SessionKey(1))
    Unload Me
End Select
End Sub
