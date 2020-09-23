VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3390
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
   ScaleHeight     =   2745
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton CMDOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Protocol Type: Tcp - Yahoo! Web based Messenger and Chat2 Protocol"
      Height          =   495
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Language: Visual Basic"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   2655
   End
   Begin VB.Image IMGG 
      Height          =   240
      Left            =   360
      Picture         =   "FrmAbout.frx":0000
      Top             =   240
      Width           =   240
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Project! Messenger"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Email: visual_basic_software@yahoo.com"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Justin LeBlanc"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label lblGeneral 
      BackStyle       =   0  'Transparent
      Caption         =   "Project! Messenger Build (1,0,1)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDOk_Click()
Me.Hide
End Sub
