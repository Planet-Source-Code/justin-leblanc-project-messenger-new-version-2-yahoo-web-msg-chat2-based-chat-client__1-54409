VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmRemove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remove User"
   ClientHeight    =   4125
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton CMDCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton CMDRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   3600
      Width           =   855
   End
   Begin VB.TextBox txtAuth 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
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
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2655
   End
   Begin VB.CheckBox CKBRemoveName 
      Caption         =   "Also Remove my Name from their Buddylist"
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
      TabIndex        =   0
      Top             =   720
      Width           =   4095
   End
   Begin SHDocVwCtl.WebBrowser WBImg 
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   5055
      ExtentX         =   8916
      ExtentY         =   2778
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.TextBox txtToken 
      Height          =   315
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "(Case Sensitive)"
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
      Index           =   2
      Left            =   2280
      TabIndex        =   6
      Top             =   3240
      Width           =   1575
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "For Authentication purposes please type the Word/Letters from the following image:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "User to Remove"
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
      TabIndex        =   2
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "FrmRemove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDCancel_Click()
Unload Me
End Sub

Private Sub CMDRemove_Click()
Dim Group As String
  Group = GetGroupName(txtWho.Text, FrmMessenger.TreeView)
If Len(Group) = 0 Then
  Unload Me
  Exit Sub
End If
DoEvents
If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgRemoveUser(YCurrentId, LCase(txtWho.Text), Group, txtAuth.Text, txtToken.Text, SessionKey(1))
DoEvents
If CKBRemoveName.Value = 1 Then
If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgDenyAdd(YCurrentId, LCase(txtWho.Text), "Thanks, but no thanks.", SessionKey(1))
End If
  Unload Me
End Sub
