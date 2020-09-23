VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FrmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
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
   ScaleHeight     =   5625
   ScaleWidth      =   5310
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox txtMessage 
      Height          =   615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   4440
      Width           =   3615
   End
   Begin VB.CommandButton CMDCencel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   11
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton CMDAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   5160
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
      TabIndex        =   9
      Top             =   3840
      Width           =   2055
   End
   Begin SHDocVwCtl.WebBrowser WBImg 
      Height          =   1575
      Left            =   120
      TabIndex        =   8
      Top             =   2160
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
      Location        =   "http:///"
   End
   Begin VB.ComboBox CBBIdentity 
      Height          =   330
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2655
   End
   Begin VB.ComboBox CBBGroup 
      Height          =   330
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtWho 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtToken 
      Height          =   315
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter a brief introduction to this person (optional):"
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
      Index           =   5
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   4215
   End
   Begin VB.Label lblGeneral 
      Alignment       =   2  'Center
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
      Index           =   4
      Left            =   2280
      TabIndex        =   7
      Top             =   3840
      Width           =   1455
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
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   4575
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose the identity you would like this person to see:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4935
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose or enter a group for this person:"
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
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "User you wish to Add:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDAdd_Click()
If FrmLogin.SockYahooChat2.State = sckConnected Then FrmLogin.SockYahooChat2.SendData WebMsgAddUser(CBBIdentity.Text, txtWho.Text, txtMessage.Text, CBBGroup.Text, txtAuth.Text, txtToken.Text, SessionKey(1))
  Call AllNodes(FrmMessenger.TreeView, True)
  Unload Me
End Sub

Private Sub CMDCencel_Click()
Unload Me
End Sub

