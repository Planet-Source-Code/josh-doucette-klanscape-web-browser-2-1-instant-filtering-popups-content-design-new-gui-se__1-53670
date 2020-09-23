VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form4 
   Caption         =   "KlanScape Text Editor"
   ClientHeight    =   6960
   ClientLeft      =   150
   ClientTop       =   825
   ClientWidth     =   8115
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   6960
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   1200
      ScaleHeight     =   2175
      ScaleWidth      =   4695
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command4 
         Caption         =   "Close"
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Find"
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   480
         Width           =   3855
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   4215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Does String Exist?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   2175
         Left            =   0
         Top             =   0
         Width           =   4695
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         Height          =   255
         Left            =   1920
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   1200
      ScaleHeight     =   2295
      ScaleWidth      =   4695
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton Command3 
         Caption         =   "Close"
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2760
         TabIndex        =   6
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   360
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Find"
         Height          =   375
         Left            =   1080
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   4215
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "and"
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Find a string between:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   4095
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFC0C0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF8080&
         Height          =   2295
         Left            =   0
         Top             =   0
         Width           =   4695
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   6705
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13811
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   6975
   End
   Begin VB.Menu filemnu 
      Caption         =   "&File"
      Begin VB.Menu fileexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu gfdxsgf 
      Caption         =   "Tools"
      Begin VB.Menu fhdxhdxh 
         Caption         =   "Does String Exist?"
      End
      Begin VB.Menu gdszgsdgz 
         Caption         =   "Find String Between"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
Label3.Caption = StringBetween(Text1.Text, Text2.Text, Text3.Text)
End Sub

Private Sub Command2_Click()
On Error Resume Next
a = Exists_In_String(Text1.Text, Text5.Text)
Label4.Caption = a


End Sub

Private Sub Command3_Click()
Picture1.Visible = False


End Sub

Private Sub Command4_Click()
Picture2.Visible = False

End Sub

Private Sub fhdxhdxh_Click()
Picture2.Visible = True

End Sub

Private Sub fileexit_Click()
Me.Hide
End Sub

Private Sub Form_Resize()
On Error Resume Next
Text1.Width = Me.ScaleWidth
Text1.Height = Me.ScaleHeight - StatusBar1.Height
End Sub

Private Sub gdszgsdgz_Click()
Picture1.Visible = True

End Sub

