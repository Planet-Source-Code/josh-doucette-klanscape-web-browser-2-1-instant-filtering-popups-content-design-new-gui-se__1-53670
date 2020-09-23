VERSION 5.00
Begin VB.Form Form2 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Filter is now active."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9855
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      FillColor       =   &H00C0FFFF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   10815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   10980
      TabIndex        =   0
      Top             =   0
      Width           =   90
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos& _
                Lib "user32" ( _
                ByVal hWnd As Long, _
                ByVal hWndInsertAfter As Long, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal cX As Long, _
                ByVal cY As Long, _
                ByVal wFlags As Long)
                Private Const SWP_NOMOVE = &H2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOSIZE = &H1
Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0
Private Const HWND_TOPMOST = -1
Private Const HWND_BOTTOM = 1
                

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
On Error Resume Next
Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOACTIVATE)

End Sub

Private Sub Form_Click()
Me.Hide
End Sub

Private Sub Form_Resize()
On Error Resume Next
Me.Top = 0
Me.Left = 0
Me.Width = Screen.Width
Shape1.Width = Screen.Width


End Sub

Private Sub Label1_Click()
Me.Hide

End Sub

Private Sub Label2_Click()
Label1.Caption = "By Thomas Robb of KKK | eMail : webmaster@klansoft.com"
End Sub
