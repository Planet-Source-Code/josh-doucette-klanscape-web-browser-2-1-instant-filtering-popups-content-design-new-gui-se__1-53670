VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form5 
   Caption         =   "KlanScape"
   ClientHeight    =   7965
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   9480
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   7965
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   741
      ButtonWidth     =   2143
      Appearance      =   1
      TextAlignment   =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "Redo"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Home Page"
            Key             =   "Home Page"
            Object.ToolTipText     =   "Official Homepage"
            ImageKey        =   "Macro"
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7575
      ExtentX         =   13361
      ExtentY         =   13150
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
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   4140
      Top             =   3735
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":000C
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":011E
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0230
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0342
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form5.frx":0454
            Key             =   "Macro"
         EndProperty
      EndProperty
   End
   Begin VB.Menu jtfj 
      Caption         =   "&File"
      Begin VB.Menu gdsgdfs 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu hfhdeshdf 
      Caption         =   "&Navigate"
      Begin VB.Menu go2 
         Caption         =   "Go To.."
      End
      Begin VB.Menu fafs 
         Caption         =   "-"
      End
      Begin VB.Menu fnbxdbf 
         Caption         =   "Back"
      End
      Begin VB.Menu bdxfbgd 
         Caption         =   "Next"
      End
      Begin VB.Menu dzgsdzg 
         Caption         =   "Refresh"
      End
      Begin VB.Menu fdxx 
         Caption         =   "-"
      End
      Begin VB.Menu gdsxfxg 
         Caption         =   "Stop"
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "Back"
WebBrowser1.GoBack

        Case "Forward"
WebBrowser1.GoForward
Case "Stop"
          WebBrowser1.Stop
          

        Case "Refresh"
          WebBrowser1.refreh
          

        Case "Home Page"
          WebBrowser1.Navigate ("http://www.klansoft.com/klanscape")
          

    End Select
End Sub

Private Sub bdxfbgd_Click()
On Error Resume Next
WebBrowser1.GoForward

End Sub

Private Sub dzgsdzg_Click()
On Error Resume Next
WebBrowser1.Refresh

End Sub

Private Sub fnbxdbf_Click()
On Error Resume Next
WebBrowser1.GoBack

End Sub

Private Sub Form_Resize()
On Error Resume Next
WebBrowser1.Width = Me.ScaleWidth
WebBrowser1.Height = Me.ScaleHeight - Toolbar1.Height

End Sub

Private Sub gdsgdfs_Click()
Me.Hide

End Sub

Private Sub gdsxfxg_Click()
On Error Resume Next
WebBrowser1.Stop

End Sub

Private Sub go2_Click()
On Error Resume Next
fa = InputBox("URL to go to?", "Address", "http://www.klansoft.com")
WebBrowser1.Navigate fa

End Sub
