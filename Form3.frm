VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "KlanScape"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7395
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7605
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "Form3.frx":0000
      Top             =   1440
      Width           =   6615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close Window"
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   6960
      Width           =   5055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go to the official project website"
      Height          =   375
      Left            =   960
      TabIndex        =   4
      Top             =   6480
      Width           =   5055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "I liked this code and would like to vote for it at Planet Source Code"
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   6000
      Width           =   5055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   $"Form3.frx":044B
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   360
      TabIndex        =   7
      Top             =   4800
      Width           =   6615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "IF YOU LIKE THIS PROJECT I HAVE SPENT THIS MUCH TIME ON..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   4560
      Width           =   5895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "I hope you are learning from this project!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please support development of KlanScape Brower w/Source Code"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   840
      Width           =   5895
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "THIS SOURCE CODE AND PROJECT ARE COPYRIGHTED"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF8080&
      Height          =   5535
      Left            =   120
      Top             =   240
      Width           =   6975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Form5.Show
Form5.WebBrowser1.Navigate ("http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?lngWId=1&blnAuthorSearch=TRUE&lngAuthorId=2284543731&strAuthorName=Josh%20Doucette&txtMaxNumberOfEntriesPerPage=25")


End Sub

Private Sub Command2_Click()
On Error Resume Next
Form5.Show
Form5.WebBrowser1.Navigate ("http://www.klansoft.com/klanscape")

End Sub

Private Sub Command3_Click()
Me.Hide

End Sub
