VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "KlanScape : The Ultimate Internet Browser"
   ClientHeight    =   8640
   ClientLeft      =   1635
   ClientTop       =   1245
   ClientWidth     =   10410
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8640
   ScaleWidth      =   10410
   Begin MSComDlg.CommonDialog thedialog 
      Left            =   3960
      Top             =   8040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer7 
      Interval        =   1
      Left            =   3960
      Top             =   7800
   End
   Begin VB.TextBox Text25 
      Height          =   285
      Left            =   5280
      TabIndex        =   139
      Text            =   "yes"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox quickpgsource 
      Height          =   1335
      Left            =   8760
      MultiLine       =   -1  'True
      TabIndex        =   138
      Top             =   6960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer6 
      Interval        =   15
      Left            =   3360
      Top             =   7680
   End
   Begin VB.PictureBox Picture10 
      BorderStyle     =   0  'None
      Height          =   3015
      Left            =   3600
      ScaleHeight     =   3015
      ScaleWidth      =   5535
      TabIndex        =   111
      Top             =   2760
      Visible         =   0   'False
      Width           =   5535
      Begin VB.Frame Frame5 
         Caption         =   "Information Options"
         Height          =   3015
         Left            =   120
         TabIndex        =   112
         Top             =   120
         Width           =   5535
         Begin VB.CommandButton Command16 
            Caption         =   "Close"
            Height          =   375
            Left            =   1920
            TabIndex        =   140
            Top             =   2280
            Width           =   1335
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Get Site Name and URL"
            Height          =   255
            Left            =   240
            TabIndex        =   119
            Top             =   360
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Detect Scripts"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   600
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Detect Swear Words"
            Height          =   255
            Left            =   240
            TabIndex        =   117
            Top             =   840
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Detect <!-- Hidden Messages -->"
            Height          =   255
            Left            =   240
            TabIndex        =   116
            Top             =   1080
            Value           =   1  'Checked
            Width           =   4815
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Detect Possible Viruses"
            Height          =   255
            Left            =   240
            TabIndex        =   115
            Top             =   1320
            Width           =   4815
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Scan For:"
            Height          =   255
            Left            =   240
            TabIndex        =   114
            Top             =   1560
            Width           =   1095
         End
         Begin VB.TextBox Text10 
            Height          =   285
            Left            =   1320
            TabIndex        =   113
            Top             =   1560
            Width           =   3015
         End
      End
   End
   Begin VB.Timer Timer5 
      Interval        =   1
      Left            =   5280
      Top             =   7680
   End
   Begin VB.PictureBox Picture3 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   10410
      TabIndex        =   77
      Top             =   570
      Width           =   10410
      Begin VB.ComboBox addressbar1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   79
         Text            =   "www.google.com"
         Top             =   120
         Width           =   6255
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6480
         TabIndex        =   78
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Timer Timer4 
      Interval        =   15
      Left            =   4440
      Top             =   7800
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   6500
      Left            =   6720
      Top             =   7800
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8760
      Top             =   7320
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   3135
      TabIndex        =   3
      Top             =   1200
      Width           =   3135
      Begin TabDlg.SSTab SSTab1 
         Height          =   7095
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   12515
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Custom"
         TabPicture(0)   =   "Form1.frx":08CA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Check1"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Frame4"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame15"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Privacy"
         TabPicture(1)   =   "Form1.frx":08E6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame6"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame7"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame12"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "Frame13"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).Control(4)=   "Frame11"
         Tab(1).Control(4).Enabled=   0   'False
         Tab(1).Control(5)=   "Frame9"
         Tab(1).Control(5).Enabled=   0   'False
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Edit 1"
         TabPicture(2)   =   "Form1.frx":0902
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame10"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "Frame8"
         Tab(2).Control(1).Enabled=   0   'False
         Tab(2).Control(2)=   "Frame14"
         Tab(2).Control(2).Enabled=   0   'False
         Tab(2).ControlCount=   3
         TabCaption(3)   =   "Edit 2"
         TabPicture(3)   =   "Form1.frx":091E
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Frame16"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).Control(1)=   "Frame17"
         Tab(3).Control(1).Enabled=   0   'False
         Tab(3).Control(2)=   "Frame18"
         Tab(3).Control(2).Enabled=   0   'False
         Tab(3).ControlCount=   3
         Begin VB.Frame Frame9 
            Caption         =   "Inspect Source Code"
            Height          =   855
            Left            =   -74880
            TabIndex        =   173
            Top             =   4440
            Width           =   2775
            Begin VB.CommandButton Command9 
               Caption         =   "URL"
               Height          =   375
               Left            =   1440
               TabIndex        =   175
               Top             =   360
               Width           =   1215
            End
            Begin VB.CommandButton Command17 
               Caption         =   "Page"
               Height          =   375
               Left            =   120
               TabIndex        =   174
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame18 
            Caption         =   "Delete / Replace element"
            Height          =   1335
            Left            =   -74880
            TabIndex        =   167
            Top             =   5400
            Width           =   2655
            Begin VB.TextBox Text35 
               Height          =   495
               Left            =   120
               MultiLine       =   -1  'True
               TabIndex        =   170
               Text            =   "Form1.frx":093A
               ToolTipText     =   "If replacing, what to replace it with."
               Top             =   720
               Width           =   2415
            End
            Begin VB.CommandButton Command14 
               Caption         =   "Replace"
               Height          =   375
               Left            =   1440
               TabIndex        =   169
               Top             =   240
               Width           =   1095
            End
            Begin VB.CommandButton Command26 
               Caption         =   "Delete"
               Height          =   375
               Left            =   120
               TabIndex        =   168
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame17 
            Caption         =   "Edit Active Elements"
            Height          =   2415
            Left            =   -74880
            TabIndex        =   149
            Top             =   2880
            Width           =   2775
            Begin VB.TextBox Text34 
               Height          =   285
               Left            =   1440
               TabIndex        =   166
               Text            =   "90"
               Top             =   2040
               Width           =   1095
            End
            Begin VB.CheckBox Check21 
               Caption         =   "Height:"
               Height          =   255
               Left            =   120
               TabIndex        =   165
               Top             =   2040
               Width           =   1335
            End
            Begin VB.TextBox Text33 
               Height          =   285
               Left            =   1440
               TabIndex        =   164
               Text            =   "90"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.CheckBox Check20 
               Caption         =   "Width:"
               Height          =   255
               Left            =   120
               TabIndex        =   163
               Top             =   1800
               Width           =   1335
            End
            Begin VB.TextBox Text32 
               Height          =   285
               Left            =   1440
               TabIndex        =   162
               Text            =   "bold"
               Top             =   1560
               Width           =   1095
            End
            Begin VB.CheckBox Check18 
               Caption         =   "Style:"
               Height          =   255
               Left            =   120
               TabIndex        =   161
               Top             =   1560
               Width           =   1335
            End
            Begin VB.TextBox Text31 
               Height          =   285
               Left            =   1440
               TabIndex        =   160
               Text            =   "Tahoma"
               Top             =   1320
               Width           =   1095
            End
            Begin VB.CheckBox Check17 
               Caption         =   "Font:"
               Height          =   255
               Left            =   120
               TabIndex        =   159
               Top             =   1320
               Width           =   1335
            End
            Begin VB.TextBox Text30 
               Height          =   285
               Left            =   1440
               TabIndex        =   158
               Text            =   "red"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.CheckBox Check16 
               Caption         =   "Fore Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   157
               Top             =   1080
               Width           =   1335
            End
            Begin VB.TextBox Text29 
               Height          =   285
               Left            =   1440
               TabIndex        =   156
               Text            =   "red"
               Top             =   840
               Width           =   1095
            End
            Begin VB.CheckBox Check15 
               Caption         =   "Border Color:"
               Height          =   255
               Left            =   120
               TabIndex        =   155
               Top             =   840
               Width           =   1335
            End
            Begin VB.TextBox Text28 
               Height          =   285
               Left            =   1440
               TabIndex        =   154
               Text            =   "1"
               Top             =   600
               Width           =   1095
            End
            Begin VB.CheckBox Check14 
               Caption         =   "Border Width:"
               Height          =   255
               Left            =   120
               TabIndex        =   153
               Top             =   600
               Width           =   1335
            End
            Begin VB.TextBox Text27 
               Height          =   285
               Left            =   1440
               TabIndex        =   152
               Text            =   "black"
               Top             =   360
               Width           =   1095
            End
            Begin VB.CheckBox Check13 
               Caption         =   "BGColor:"
               Height          =   255
               Left            =   120
               TabIndex        =   151
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.Frame Frame16 
            Caption         =   "Active Elements"
            Height          =   2415
            Left            =   -74880
            TabIndex        =   143
            Top             =   360
            Width           =   2775
            Begin VB.TextBox Label51 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               Height          =   615
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   150
               Text            =   "Form1.frx":0955
               Top             =   1680
               Width           =   2535
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               Caption         =   "Type:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   172
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label36 
               Caption         =   "---"
               Height          =   255
               Left            =   720
               TabIndex        =   171
               Top             =   1080
               Width           =   1575
            End
            Begin VB.Label Label50 
               Caption         =   "Attribute HTML:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   148
               Top             =   1440
               Width           =   2535
            End
            Begin VB.Label Label49 
               Caption         =   "---"
               Height          =   255
               Left            =   720
               TabIndex        =   147
               Top             =   840
               Width           =   1695
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               Caption         =   " ID:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   146
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label23 
               Caption         =   "Active Attribute:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   145
               Top             =   360
               Width           =   2535
            End
            Begin VB.Label Label28 
               Caption         =   "---"
               Height          =   255
               Left            =   120
               TabIndex        =   144
               Top             =   600
               Width           =   2535
            End
         End
         Begin VB.Frame Frame15 
            Height          =   975
            Left            =   120
            TabIndex        =   135
            Top             =   5520
            Width           =   2655
            Begin VB.CheckBox Check12 
               Caption         =   "Add HTML to each page"
               Height          =   255
               Left            =   120
               TabIndex        =   137
               Top             =   0
               Width           =   2175
            End
            Begin VB.TextBox Text12 
               Height          =   375
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   136
               Text            =   "Form1.frx":095B
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame Frame14 
            Caption         =   "Design Mode"
            Height          =   735
            Left            =   -74880
            TabIndex        =   131
            Top             =   480
            Width           =   2775
            Begin VB.CommandButton Command25 
               Caption         =   "OFF"
               Height          =   375
               Left            =   600
               TabIndex        =   134
               Top             =   240
               Width           =   495
            End
            Begin VB.CheckBox Check11 
               Caption         =   "Drag-And-Drop"
               Height          =   255
               Left            =   1200
               TabIndex        =   133
               Top             =   240
               Width           =   1455
            End
            Begin VB.CommandButton Command24 
               Caption         =   "ON"
               Height          =   375
               Left            =   120
               TabIndex        =   132
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "SSL"
            Height          =   735
            Left            =   -74880
            TabIndex        =   109
            Top             =   6120
            Width           =   2775
            Begin VB.Label Label42 
               Caption         =   "---"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   120
               TabIndex        =   110
               Top             =   240
               Width           =   2535
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Active Cookies"
            Height          =   615
            Left            =   -74880
            TabIndex        =   105
            Top             =   5400
            Width           =   2775
            Begin VB.TextBox Text14 
               BackColor       =   &H8000000F&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   120
               Locked          =   -1  'True
               TabIndex        =   106
               Text            =   "---"
               Top             =   240
               Width           =   2415
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Add HTML"
            Height          =   2775
            Left            =   -74880
            TabIndex        =   102
            Top             =   3120
            Width           =   2775
            Begin VB.TextBox Text26 
               Height          =   285
               Left            =   120
               TabIndex        =   142
               Text            =   "Where"
               ToolTipText     =   "Where to add the HTML"
               Top             =   600
               Width           =   2535
            End
            Begin VB.ComboBox Combo3 
               Height          =   315
               ItemData        =   "Form1.frx":0D86
               Left            =   120
               List            =   "Form1.frx":0D93
               Style           =   2  'Dropdown List
               TabIndex        =   141
               Top             =   240
               Width           =   2535
            End
            Begin VB.CommandButton Command15 
               Caption         =   "Write HTML to page"
               Height          =   375
               Left            =   120
               TabIndex        =   104
               Top             =   2280
               Width           =   2535
            End
            Begin VB.TextBox Text13 
               Height          =   1245
               Left            =   120
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   103
               Top             =   960
               Width           =   2535
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Script Errors"
            Height          =   975
            Left            =   -74880
            TabIndex        =   69
            Top             =   3360
            Width           =   2775
            Begin VB.PictureBox Picture12 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   120
               ScaleHeight     =   585
               ScaleWidth      =   585
               TabIndex        =   177
               Top             =   240
               Width           =   615
               Begin VB.Image Image3 
                  Height          =   480
                  Left            =   43
                  Picture         =   "Form1.frx":0DE6
                  Top             =   30
                  Width           =   480
               End
            End
            Begin VB.OptionButton Option11 
               Caption         =   "Ignore Script Errors"
               Height          =   255
               Left            =   840
               TabIndex        =   71
               Top             =   480
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton Option9 
               Caption         =   "Show Script Errors"
               Height          =   255
               Left            =   840
               TabIndex        =   70
               Top             =   240
               Width           =   1815
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Page Colors"
            Height          =   1695
            Left            =   -74880
            TabIndex        =   68
            Top             =   1320
            Width           =   2775
            Begin VB.CommandButton Command23 
               Caption         =   "write"
               Height          =   255
               Left            =   2040
               TabIndex        =   101
               Top             =   1320
               Width           =   495
            End
            Begin VB.TextBox Text20 
               Height          =   285
               Left            =   960
               TabIndex        =   100
               Text            =   "white"
               Top             =   1320
               Width           =   975
            End
            Begin VB.CommandButton Command22 
               Caption         =   "write"
               Height          =   255
               Left            =   2040
               TabIndex        =   98
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox Text19 
               Height          =   285
               Left            =   960
               TabIndex        =   97
               Text            =   "white"
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton Command21 
               Caption         =   "write"
               Height          =   255
               Left            =   2040
               TabIndex        =   95
               Top             =   840
               Width           =   495
            End
            Begin VB.TextBox Text18 
               Height          =   285
               Left            =   960
               TabIndex        =   94
               Text            =   "white"
               Top             =   840
               Width           =   975
            End
            Begin VB.CommandButton Command20 
               Caption         =   "write"
               Height          =   255
               Left            =   2040
               TabIndex        =   92
               Top             =   600
               Width           =   495
            End
            Begin VB.TextBox Text17 
               Height          =   285
               Left            =   960
               TabIndex        =   91
               Text            =   "white"
               Top             =   600
               Width           =   975
            End
            Begin VB.CommandButton Command18 
               Caption         =   "write"
               Height          =   255
               Left            =   2040
               TabIndex        =   89
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox Text16 
               Height          =   285
               Left            =   960
               TabIndex        =   88
               Text            =   "black"
               Top             =   360
               Width           =   975
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               Caption         =   "A Link:"
               Height          =   255
               Left            =   120
               TabIndex        =   99
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               Caption         =   "Old Links:"
               Height          =   255
               Left            =   120
               TabIndex        =   96
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               Caption         =   "Links:"
               Height          =   255
               Left            =   120
               TabIndex        =   93
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "ForeColor:"
               Height          =   255
               Left            =   120
               TabIndex        =   90
               Top             =   600
               Width           =   735
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "bgColor:"
               Height          =   255
               Left            =   120
               TabIndex        =   87
               Top             =   360
               Width           =   735
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "File Downloads"
            Height          =   1095
            Left            =   -74880
            TabIndex        =   53
            Top             =   2160
            Width           =   2775
            Begin VB.PictureBox Picture11 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   120
               ScaleHeight     =   585
               ScaleWidth      =   585
               TabIndex        =   179
               Top             =   240
               Width           =   615
               Begin VB.Image Image5 
                  Height          =   480
                  Left            =   45
                  Picture         =   "Form1.frx":16B0
                  Top             =   30
                  Visible         =   0   'False
                  Width           =   480
               End
               Begin VB.Image Image2 
                  Height          =   480
                  Left            =   43
                  Picture         =   "Form1.frx":1F7A
                  Top             =   30
                  Width           =   480
               End
            End
            Begin VB.OptionButton Option6 
               Caption         =   "Block Downloads"
               Height          =   255
               Left            =   960
               TabIndex        =   67
               Top             =   720
               Width           =   1695
            End
            Begin VB.OptionButton Option5 
               Caption         =   "Block and Alert Me"
               Height          =   255
               Left            =   960
               TabIndex        =   66
               Top             =   480
               Width           =   1695
            End
            Begin VB.OptionButton Option4 
               Caption         =   "Allow Downloads"
               Height          =   255
               Left            =   960
               TabIndex        =   65
               Top             =   240
               Value           =   -1  'True
               Width           =   1695
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Popup Windows"
            Height          =   1695
            Left            =   -74880
            TabIndex        =   45
            Top             =   360
            Width           =   2775
            Begin VB.PictureBox Picture7 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   120
               ScaleHeight     =   585
               ScaleWidth      =   585
               TabIndex        =   178
               Top             =   240
               Width           =   615
               Begin VB.Image Image1 
                  Height          =   480
                  Left            =   43
                  Picture         =   "Form1.frx":2844
                  Top             =   30
                  Width           =   480
               End
            End
            Begin VB.ComboBox Combo2 
               Height          =   315
               ItemData        =   "Form1.frx":310E
               Left            =   840
               List            =   "Form1.frx":3118
               Style           =   2  'Dropdown List
               TabIndex        =   73
               Top             =   1200
               Width           =   1515
            End
            Begin VB.OptionButton Option3 
               Caption         =   "Block Popups"
               Height          =   255
               Left            =   840
               TabIndex        =   72
               Top             =   960
               Width           =   1455
            End
            Begin VB.OptionButton Option30 
               Caption         =   "Block and Alert"
               Height          =   255
               Left            =   840
               TabIndex        =   54
               Top             =   720
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Allow Popups"
               Height          =   255
               Left            =   840
               TabIndex        =   49
               Top             =   240
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton Option1 
               Caption         =   "Ask Me"
               Height          =   255
               Left            =   840
               TabIndex        =   48
               Top             =   480
               Width           =   1455
            End
            Begin VB.Label Label25 
               BackStyle       =   0  'Transparent
               Caption         =   "0"
               Height          =   255
               Left            =   120
               TabIndex        =   51
               Top             =   1200
               Width           =   615
            End
            Begin VB.Label Label24 
               BackStyle       =   0  'Transparent
               Caption         =   "Blocked:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   50
               Top             =   960
               Width           =   615
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Filter Settings"
            Height          =   1935
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   2775
            Begin VB.PictureBox Picture8 
               BorderStyle     =   0  'None
               Height          =   1570
               Left            =   120
               ScaleHeight     =   1575
               ScaleWidth      =   2535
               TabIndex        =   56
               Top             =   240
               Visible         =   0   'False
               Width           =   2535
               Begin VB.ComboBox Combo1 
                  Height          =   315
                  ItemData        =   "Form1.frx":313A
                  Left            =   120
                  List            =   "Form1.frx":3159
                  Style           =   2  'Dropdown List
                  TabIndex        =   60
                  Top             =   840
                  Width           =   2295
               End
               Begin VB.CommandButton Command12 
                  Caption         =   "Return to Standard Filter"
                  Height          =   375
                  Left            =   120
                  TabIndex        =   59
                  Top             =   1200
                  Width           =   2175
               End
               Begin VB.Label Label27 
                  Alignment       =   2  'Center
                  Caption         =   "Select how live you would like to be. No 'filtering' message."
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   120
                  TabIndex        =   58
                  Top             =   360
                  Width           =   2295
               End
               Begin VB.Label Label26 
                  Alignment       =   2  'Center
                  Caption         =   "Live Filter is Enabled"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   240
                  TabIndex        =   57
                  Top             =   120
                  Width           =   1935
               End
            End
            Begin VB.PictureBox Picture13 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFC0C0&
               ForeColor       =   &H80000008&
               Height          =   615
               Left            =   120
               ScaleHeight     =   585
               ScaleWidth      =   585
               TabIndex        =   180
               Top             =   240
               Width           =   615
               Begin VB.Image Image4 
                  Height          =   480
                  Left            =   43
                  Picture         =   "Form1.frx":31F7
                  Top             =   30
                  Width           =   480
               End
            End
            Begin MSComctlLib.Slider Slider1 
               Height          =   255
               Left            =   840
               TabIndex        =   27
               Top             =   480
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               LargeChange     =   1
               Min             =   1
               Max             =   4
               SelectRange     =   -1  'True
               SelStart        =   1
               Value           =   1
               TextPosition    =   1
            End
            Begin VB.Label Label13 
               BackStyle       =   0  'Transparent
               Caption         =   "Option 1 - Filter after navigation complete. This is one of the fastest filters and will work for most purposes."
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   975
               Left            =   120
               TabIndex        =   28
               Top             =   960
               Width           =   2535
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enable custom filters"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Value           =   1  'Checked
            Width           =   2415
         End
         Begin VB.Frame Frame3 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   15
            Top             =   4560
            Width           =   2775
            Begin VB.CheckBox Check10 
               Caption         =   "Block"
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   0
               Value           =   1  'Checked
               Width           =   735
            End
            Begin VB.TextBox Text7 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   17
               Text            =   "at least 18"
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox Text6 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   16
               Text            =   "Adult Content"
               Top             =   480
               Width           =   1335
            End
            Begin VB.Label Label10 
               Caption         =   "If site contains:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   19
               Top             =   240
               Width           =   1215
            End
            Begin VB.Label Label9 
               Caption         =   "Reason:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   18
               Top             =   480
               Width           =   1215
            End
         End
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   10
            Top             =   3600
            Width           =   2775
            Begin VB.CheckBox Check8 
               Caption         =   "Redirect"
               Height          =   255
               Left            =   120
               TabIndex        =   55
               Top             =   0
               Value           =   1  'Checked
               Width           =   975
            End
            Begin VB.TextBox Text4 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   12
               Text            =   "microsoft.com"
               Top             =   480
               Width           =   1335
            End
            Begin VB.TextBox Text3 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1320
               TabIndex        =   11
               Text            =   "macintosh"
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label6 
               Caption         =   "Then goto:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   14
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "If site contains:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   13
               Top             =   240
               Width           =   1215
            End
         End
         Begin VB.Frame Frame2 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   5
            Top             =   2640
            Width           =   2775
            Begin VB.CheckBox Check9 
               Caption         =   "Alert"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   0
               Value           =   1  'Checked
               Width           =   615
            End
            Begin VB.TextBox Text5 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   7
               Text            =   "'Apple' found!!"
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1200
               TabIndex        =   6
               Text            =   "apple"
               Top             =   240
               Width           =   1455
            End
            Begin VB.Label Label4 
               Caption         =   "Then Popup:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   9
               Top             =   480
               Width           =   975
            End
            Begin VB.Label Label1 
               Caption         =   "If site contains:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   1095
            End
         End
      End
   End
   Begin TabDlg.SSTab sstab2 
      Height          =   6495
      Left            =   3120
      TabIndex        =   1
      Top             =   1080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   11456
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Website"
      TabPicture(0)   =   "Form1.frx":3AC1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "WebBrowser1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Picture1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Site Information"
      TabPicture(1)   =   "Form1.frx":3ADD
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "Label14"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "Label17"
      Tab(1).Control(5)=   "Label18"
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(7)=   "Label41"
      Tab(1).Control(8)=   "Label43"
      Tab(1).Control(9)=   "Label44"
      Tab(1).Control(10)=   "Label45"
      Tab(1).Control(11)=   "Label46"
      Tab(1).Control(12)=   "Label47"
      Tab(1).Control(13)=   "Text8"
      Tab(1).Control(14)=   "Text9"
      Tab(1).Control(15)=   "Command10"
      Tab(1).Control(16)=   "Command11"
      Tab(1).Control(17)=   "Picture6"
      Tab(1).Control(18)=   "Picture9"
      Tab(1).Control(19)=   "Text15"
      Tab(1).Control(20)=   "Text21"
      Tab(1).Control(21)=   "Text22"
      Tab(1).Control(22)=   "Text23"
      Tab(1).Control(23)=   "Text24"
      Tab(1).ControlCount=   24
      TabCaption(2)   =   "Site Contents"
      TabPicture(2)   =   "Form1.frx":3AF9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label29"
      Tab(2).Control(1)=   "Label30"
      Tab(2).Control(2)=   "tv"
      Tab(2).Control(3)=   "newtext2"
      Tab(2).Control(4)=   "list1"
      Tab(2).Control(5)=   "Command19"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "Text-Only"
      TabPicture(3)   =   "Form1.frx":3B15
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label31"
      Tab(3).Control(1)=   "text1"
      Tab(3).ControlCount=   2
      Begin VB.TextBox Text24 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   130
         Text            =   "---"
         Top             =   5280
         Width           =   4695
      End
      Begin VB.TextBox Text23 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   129
         Text            =   "---"
         Top             =   5040
         Width           =   4695
      End
      Begin VB.TextBox Text22 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   128
         Text            =   "---"
         Top             =   4800
         Width           =   4695
      End
      Begin VB.TextBox Text21 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   127
         Text            =   "---"
         Top             =   4560
         Width           =   4695
      End
      Begin VB.TextBox Text15 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   -73680
         Locked          =   -1  'True
         TabIndex        =   126
         Text            =   "---"
         Top             =   4320
         Width           =   4575
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Refresh"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -70200
         TabIndex        =   82
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox list1 
         Height          =   315
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   81
         Top             =   840
         Width           =   5535
      End
      Begin VB.TextBox newtext2 
         Height          =   285
         Left            =   -70440
         TabIndex        =   74
         Text            =   "empty"
         Top             =   5640
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   -74760
         ScaleHeight     =   735
         ScaleWidth      =   4575
         TabIndex        =   62
         Top             =   3120
         Visible         =   0   'False
         Width           =   4575
         Begin VB.CommandButton Command13 
            Caption         =   "Save"
            Height          =   375
            Left            =   3240
            TabIndex        =   64
            Top             =   240
            Width           =   1095
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Left            =   360
            TabIndex        =   63
            Text            =   "c:\log.txt"
            Top             =   240
            Width           =   2655
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00808080&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H00808080&
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   735
            Left            =   240
            Top             =   0
            Width           =   4335
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   4935
         Left            =   0
         ScaleHeight     =   4875
         ScaleWidth      =   4995
         TabIndex        =   43
         Top             =   120
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton Command7 
            Caption         =   "Go to your Home Page"
            Height          =   375
            Left            =   240
            TabIndex        =   52
            Top             =   3360
            Width           =   2175
         End
         Begin VB.CommandButton Command8 
            Caption         =   "<- Return to previous page"
            Height          =   375
            Left            =   240
            TabIndex        =   44
            Top             =   2880
            Width           =   2175
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "with the KlanScape Filter"
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
            Left            =   2160
            TabIndex        =   86
            Top             =   840
            Width           =   4455
         End
         Begin VB.Label Label33 
            BackStyle       =   0  'Transparent
            Caption         =   "You have the following choices:"
            Height          =   255
            Left            =   240
            TabIndex        =   85
            Top             =   2520
            Width           =   3495
         End
         Begin VB.Label Label32 
            BackStyle       =   0  'Transparent
            Caption         =   "Reason Given:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   1200
            Width           =   4455
         End
         Begin VB.Label Label21 
            BackColor       =   &H00C0E0FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Site is BLOCKED."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   26.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   615
            Left            =   240
            TabIndex        =   83
            Top             =   240
            Width           =   4815
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Can't load"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   855
            Left            =   240
            TabIndex        =   61
            Top             =   1560
            Width           =   4575
         End
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   2895
         Left            =   -74880
         ScaleHeight     =   2895
         ScaleWidth      =   5055
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         Begin VB.Label Label20 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":3B31
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   240
            TabIndex        =   42
            Top             =   960
            Width           =   3975
         End
         Begin VB.Label Label22 
            BackStyle       =   0  'Transparent
            Caption         =   "Gathering Information..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   41
            Top             =   360
            Width           =   5535
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderColor     =   &H000080FF&
            Height          =   1815
            Left            =   0
            Top             =   120
            Width           =   4815
         End
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Save to file as a log"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74760
         TabIndex        =   39
         Top             =   3360
         Width           =   3375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Get Information"
         Height          =   375
         Left            =   -74760
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "Site URL goes Here."
         Top             =   1200
         Width           =   5655
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   -74760
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "Site Name Goes Here."
         Top             =   960
         Width           =   5655
      End
      Begin VB.PictureBox Picture4 
         Height          =   4815
         Left            =   0
         ScaleHeight     =   4755
         ScaleWidth      =   7395
         TabIndex        =   20
         Top             =   120
         Width           =   7455
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "and inspected using the filter."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   3360
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "KlanScape"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   0
            TabIndex        =   23
            Top             =   240
            Width           =   8175
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Your site is being loaded"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   22
            Top             =   840
            Width           =   2805
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Click here to cancel"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   1560
            Width           =   7095
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   5535
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   4935
         ExtentX         =   8705
         ExtentY         =   9763
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
      Begin MSComctlLib.TreeView tv 
         Height          =   2625
         Left            =   -74880
         TabIndex        =   80
         Top             =   1560
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   4630
         _Version        =   393217
         Indentation     =   471
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
      End
      Begin RichTextLib.RichTextBox text1 
         Height          =   4815
         Left            =   -75000
         TabIndex        =   107
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   8493
         _Version        =   393217
         TextRTF         =   $"Form1.frx":3BCA
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "Refer:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   125
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         Caption         =   "Created On:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   124
         Top             =   5040
         Width           =   1215
      End
      Begin VB.Label Label45 
         Alignment       =   2  'Center
         Caption         =   "File Size:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   123
         Top             =   4800
         Width           =   1215
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         Caption         =   "Directory:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   122
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label Label43 
         Alignment       =   2  'Center
         Caption         =   "Domain Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   121
         Top             =   4320
         Width           =   1215
      End
      Begin VB.Label Label41 
         Caption         =   "General Contents:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   120
         Top             =   3960
         Width           =   6015
      End
      Begin VB.Label Label31 
         Caption         =   "The following was extracted from the site:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -75000
         TabIndex        =   108
         Top             =   120
         Width           =   6015
      End
      Begin VB.Label Label30 
         Caption         =   "This site contains the following links (URLs)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   76
         Top             =   480
         Width           =   6015
      End
      Begin VB.Label Label29 
         Caption         =   "This site contains the following elements:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74880
         TabIndex        =   75
         Top             =   1320
         Width           =   6015
      End
      Begin VB.Label Label19 
         Caption         =   "Does your custom search come up?"
         Height          =   255
         Left            =   -74760
         TabIndex        =   38
         Top             =   2760
         Width           =   5655
      End
      Begin VB.Label Label18 
         Caption         =   "Does this site contain possible viruses."
         Height          =   255
         Left            =   -74760
         TabIndex        =   37
         Top             =   2520
         Width           =   5655
      End
      Begin VB.Label Label17 
         Caption         =   "What would you like to do with this information?"
         Height          =   255
         Left            =   -74760
         TabIndex        =   35
         Top             =   3120
         Width           =   5535
      End
      Begin VB.Label Label16 
         Caption         =   "Does this site contain hidden messages."
         Height          =   255
         Left            =   -74760
         TabIndex        =   32
         Top             =   2280
         Width           =   5655
      End
      Begin VB.Label Label15 
         Caption         =   "Does this site contain swear words."
         Height          =   255
         Left            =   -74760
         TabIndex        =   31
         Top             =   2040
         Width           =   5655
      End
      Begin VB.Label Label14 
         Caption         =   "Does this site contain scripts."
         Height          =   255
         Left            =   -74760
         TabIndex        =   30
         Top             =   1800
         Width           =   5655
      End
      Begin VB.Label Label12 
         Caption         =   "Site Information"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   29
         Top             =   1440
         Width           =   5655
      End
   End
   Begin MSComctlLib.StatusBar picture2 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8385
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17859
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8160
      Top             =   7680
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7200
      Top             =   7680
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
            Picture         =   "Form1.frx":3C55
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":41EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4789
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":52BD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   0
      Left            =   4665
      Top             =   4110
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
            Picture         =   "Form1.frx":5857
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5969
            Key             =   "FauxS-X (Green) Entire Network"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5B43
            Key             =   "Forward"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   1
      Left            =   4665
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5C55
            Key             =   "FauxS-X (Green) Log Off"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5E2F
            Key             =   "FauxS-X (Green) Control Panel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Index           =   2
      Left            =   4665
      Top             =   4110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6009
            Key             =   "Back"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":611B
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":622D
            Key             =   "Default Icon"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6407
            Key             =   "Scheduled Tasks"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":65E1
            Key             =   "Network Drive Connected"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":67BB
            Key             =   "Internet Document"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6995
            Key             =   "Floppy Drive"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6B6F
            Key             =   "Find2"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6D49
            Key             =   "Fonts"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6F23
            Key             =   "Text Document"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":70FD
            Key             =   "Settings"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":72D7
            Key             =   "Help1"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   176
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
      _ExtentY        =   1005
      ButtonWidth     =   1455
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons(2)"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Back"
            Key             =   "Back"
            Description     =   "Back"
            Object.ToolTipText     =   "Back"
            ImageKey        =   "Back"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Forward"
            Key             =   "Forward"
            Description     =   "Forward"
            Object.ToolTipText     =   "Forward"
            ImageKey        =   "Forward"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Stop"
            Key             =   "Stop"
            Description     =   "Stop"
            Object.ToolTipText     =   "Stop"
            ImageKey        =   "Default Icon"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Refresh"
            Key             =   "Refresh"
            Description     =   "Refresh"
            Object.ToolTipText     =   "Refresh"
            ImageKey        =   "Scheduled Tasks"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open URL"
            Key             =   "Open Online"
            Description     =   "Open URL"
            Object.ToolTipText     =   "Open Page for Online"
            ImageKey        =   "Network Drive Connected"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open File"
            Key             =   "Open from File"
            Description     =   "Open File"
            Object.ToolTipText     =   "Open Page from File"
            ImageKey        =   "Internet Document"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save"
            Key             =   "Save"
            Description     =   "Save Page"
            Object.ToolTipText     =   "Save Website"
            ImageKey        =   "Floppy Drive"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Find"
            Key             =   "Find"
            Description     =   "Find"
            Object.ToolTipText     =   "Find"
            ImageKey        =   "Find2"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Size"
            Key             =   "Font Size"
            Description     =   "Font Size"
            Object.ToolTipText     =   "Change the Font Size"
            ImageKey        =   "Fonts"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Source"
            Key             =   "Edit"
            Description     =   "Source"
            Object.ToolTipText     =   "Inspect Source Code"
            ImageKey        =   "Text Document"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Setup"
            Key             =   "Settings"
            Description     =   "Setup"
            Object.ToolTipText     =   "Settings"
            ImageKey        =   "Settings"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About"
            Key             =   "About"
            Description     =   "About KlanScape"
            Object.ToolTipText     =   "About"
            ImageKey        =   "Help1"
         EndProperty
      EndProperty
   End
   Begin VB.Menu filemnu 
      Caption         =   "&File"
      Begin VB.Menu newwindow1 
         Caption         =   "&New Window"
      End
      Begin VB.Menu gdsg 
         Caption         =   "-"
      End
      Begin VB.Menu gadgx 
         Caption         =   "Open"
      End
      Begin VB.Menu savepage 
         Caption         =   "Save As..."
      End
      Begin VB.Menu fzdszf 
         Caption         =   "-"
      End
      Begin VB.Menu filex 
         Caption         =   "Close"
      End
   End
   Begin VB.Menu gagsagd 
      Caption         =   "&V&iew"
      Begin VB.Menu htfdxhredsh 
         Caption         =   "Show Panel"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu weband11 
      Caption         =   "&Website"
      Begin VB.Menu findmnu1 
         Caption         =   "Find (on site)..."
      End
      Begin VB.Menu textsz1 
         Caption         =   "Text Size"
         Begin VB.Menu smallestz 
            Caption         =   "Smallest"
         End
         Begin VB.Menu smalle1 
            Caption         =   "Small"
         End
         Begin VB.Menu smallestz1 
            Caption         =   "Medium"
         End
         Begin VB.Menu smallest3 
            Caption         =   "Large"
         End
         Begin VB.Menu smallest2 
            Caption         =   "Largest"
         End
      End
      Begin VB.Menu man11 
         Caption         =   "Manipulate"
         Begin VB.Menu on11 
            Caption         =   "ON"
         End
         Begin VB.Menu off11 
            Caption         =   "OFF"
         End
      End
      Begin VB.Menu dash11 
         Caption         =   "-"
      End
      Begin VB.Menu inspectsourcex 
         Caption         =   "Inspect Source"
      End
      Begin VB.Menu prop11 
         Caption         =   "Properties"
      End
   End
   Begin VB.Menu setup11 
      Caption         =   "Setup"
      Begin VB.Menu gszgdzsg 
         Caption         =   "Search Info : Options"
      End
      Begin VB.Menu inetproperties 
         Caption         =   "Internet Properties"
      End
   End
   Begin VB.Menu fggshdsfh 
      Caption         =   "Help"
      Begin VB.Menu htfdhdh 
         Caption         =   "Official Home Page (Updates)"
      End
      Begin VB.Menu fileabout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim webdoc  As HTMLDocument
Dim texbody As HTMLBody
Dim Texob As IHTMLTxtRange
Dim J As Integer

Private Sub htfdhdh_Click()
On Error Resume Next
Form5.Show
Form5.WebBrowser1.Navigate ("http://www.klansoft.com/klanscape")

End Sub

Private Sub Option4_Click()
On Error Resume Next
Image2.Visible = True
Image5.Visible = False

End Sub

Private Sub Option5_Click()
On Error Resume Next
Image2.Visible = False
Image5.Visible = True

End Sub

Private Sub Option6_Click()
On Error Resume Next
Image2.Visible = False
Image5.Visible = True

End Sub

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
           WebBrowser1.Refresh
        Case "Open Online"
g = InputBox("URL to Open?", "Open URL", "http://www.google.com")
If Len(g) < 1 Then Exit Sub
WebBrowser1.Navigate (g)

        Case "Open from File"
           gadgx_Click
        Case "Save"
          savepage_Click
        Case "Find"
           WebBrowser1.SetFocus
    SendKeys "^f"
        Case "Font Size"
    PopupMenu textsz1
        Case "Edit"
 Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
Form4.Show
Form4.text1.Text = doc.body.innerHTML


        Case "Settings"
            MsgBox "Coming Soon"
        Case "About"
Form3.Show

    End Select
End Sub


Private Sub addressbar1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then WebBrowser1.Navigate addressbar1.Text

End Sub

Private Sub Check1_Click()
Form2.Show
If Check1.Value = 1 Then Form2.Label1.Caption = "Filter is now ENABLED!"
If Check1.Value = 0 Then Form2.Label1.Caption = "Filter is now DISABLED!"
End Sub

Private Sub Check11_Click()
On Error Resume Next
If Check11.Value = 0 Then WebBrowser1.RegisterAsDropTarget = False
If Check11.Value = 1 Then WebBrowser1.RegisterAsDropTarget = True

End Sub

Private Sub Check12_Click()
If Check12.Value = 0 Then Exit Sub
If Check12.Value = 1 Then x = MsgBox("This is a new feature introduced in this version. You must navigate to a new URL for changes to take effect. Please notice that this is for example and experiment only and is still under development; it has known problems with some webpages including Planet-Source-Code which accidently trick it into a loop. With your ongoing support this will be fixed by the next version.")

End Sub

Private Sub Combo1_Click()
Dim bs
If Combo1.Text = "15,000 - Fastest" Then Timer3.Interval = 15000
If Combo1.Text = "10,000 - Fastest" Then Timer3.Interval = 10000
If Combo1.Text = "8,000 -  Fast" Then Timer3.Interval = 8000
If Combo1.Text = "6,500 - Sticky" Then Timer3.Interval = 6500
If Combo1.Text = "4,000 - Slowish" Then Timer3.Interval = 4000
If Combo1.Text = "3,000 - Slower" Then Timer3.Interval = 3000
If Combo1.Text = "2,000 - Slow" Then Timer3.Interval = 3000
If Combo1.Text = "600 - Slowest" Then GoTo slowxx2000
If Combo1.Text = "1 - Small or Local ONLY" Then GoTo slowxx1
Exit Sub
slowxx2000:
bs = MsgBox("600 Milliseconds can cause websites to feel sticky or frozen at times. It is very live. Continue?", vbYesNo, "Live Filtering : KlanScape")
If bs = vbNo Then Combo1.Text = "6,500 - Sticky"
If bs = vbNo Then Timer3.Interval = 6500
If bs = vbNo Then Exit Sub
Timer3.Interval = 600
Exit Sub
slowxx1:
bs = MsgBox("1 Millisecond is as live as you can get. It can cause large or medium size websites to feel 'sticky' and a little frozen at times. Everything you do is instant, for example: the second you type something on a form on the site it should go through the filter (instant filter). Continue?", vbYesNo, "Live Filtering : KlanScape")
If bs = vbNo Then Combo1.Text = "6,500 - Sticky"
If bs = vbNo Then Timer3.Interval = 6500
If bs = vbNo Then Exit Sub
Timer3.Interval = 1
Exit Sub

End Sub

Private Sub Combo3_Change()
On Error Resume Next
Text26.Enabled = False
If Combo3.Text = "Add In Between HTML (Custom)" Then Text26.Enabled = True

End Sub

Private Sub Combo3_Click()
On Error Resume Next
Text26.Enabled = False
If Combo3.Text = "Add In Between HTML (Custom)" Then Text26.Enabled = True
End Sub

Private Sub Command1_Click()
WebBrowser1.Navigate (addressbar1.Text)

End Sub

Private Function WebPageContains(ByVal S As String) As Boolean
    Dim i As Long, HTMLElement
    
On Error Resume Next
    For i = 1 To WebBrowser1.Document.All.Length
        Set HTMLElement = WebBrowser1.Document.All.Item(i)
                If Not (HTMLElement Is Nothing) Then
            If InStr(1, HTMLElement.innerHTML, S, vbTextCompare) > 0 Then
                WebPageContains = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Sub Command10_Click()
On Error Resume Next
If Picture1.Visible = True Then GoTo blockedsite
Picture6.Visible = True
Timer2.Enabled = True
Exit Sub
blockedsite:
bx = MsgBox("Can not gather information on this site because it is blocked.", vbInformation, WebBrowser1.LocationURL)
Exit Sub

End Sub

Private Sub Command11_Click()
Picture9.Visible = True
vz$ = Time
Text11.Text = "c:\log" + "_" + Date$ + ".txt"

End Sub

Private Sub Command12_Click()
Timer3.Enabled = False
Timer1.Enabled = False
Picture8.Visible = False
Slider1.Value = 1


End Sub

Private Sub Command13_Click()
On Error GoTo err0x
Open Text11.Text For Output As 1
Print #1, "-- LOG FILE SAVED BY KLANSCAPE EXAMPLE : --"
Print #1, "          -- www.KlanSoft.com --           "
Print #1, ""
Print #1, Text8.Text
Print #1, Text9.Text
Print #1, Label14
Print #1, Label15
Print #1, Label16
Print #1, Label18
Print #1, Label19
Print #1, ""
Print #1, "            --  End of Log --"
Close #1
bz = MsgBox("File Saved!", vbInformation, "Log Saved to Disk")
Picture9.Visible = False
Exit Sub
err0x:
bz = MsgBox("Error saving file. Please check that the file name is vaild.", vbInformation, "Error Saving Log!")
Picture9.Visible = False
Exit Sub
End Sub

Private Sub Command14_Click()
 On Error Resume Next
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
    quickpgsource.Text = doc.body.innerHTML
    doc.activeElement.innerHTML = Text35.Text
    
    
    
End Sub

Private Sub Command15_Click()
 On Error Resume Next
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
    quickpgsource.Text = doc.body.innerHTML
If Combo3.Text = "Add AFTER current HTML" Then doc.body.innerHTML = quickpgsource.Text + Text13.Text
If Combo3.Text = "Add BEFORE current HTML" Then doc.body.innerHTML = Text13.Text + quickpgsource.Text
If Combo3.Text = "Add In Between HTML (Custom)" Then doc.body.insertAdjacentHTML Text26.Text, Text13.Text
End Sub

Private Sub Command16_Click()
On Error Resume Next
Picture10.Visible = False



End Sub

Private Sub Command17_Click()
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
Form4.Show
Form4.text1.Text = doc.body.innerHTML


End Sub

Private Sub Command18_Click()
On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
   doc.bgColor = Text16.Text
     Set Web = Nothing
End Sub

Private Sub Command19_Click()
getelement (WebBrowser1.LocationURL)
End Sub

Private Sub Command2_Click()
On Error Resume Next
WebBrowser1.GoHome


End Sub

Private Sub Command20_Click()
 On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
   doc.bgColor = Text16.Text
     Set Web = Nothing
     doc.fgColor = Text17.Text
     
End Sub

Private Sub Command21_Click()
 On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
doc.linkColor = Text18.Text

     Set Web = Nothing
     
End Sub

Private Sub Command22_Click()
 On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
doc.linkColor = Text18.Text

     Set Web = Nothing
 doc.vlinkColor = Text19.Text
 

End Sub

Private Sub Command23_Click()
 On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
doc.linkColor = Text18.Text

     Set Web = Nothing
doc.alinkColor = Text20.Text

 
End Sub

Private Sub Command24_Click()
On Error Resume Next
WebBrowser1.Document.designMode = "on"
End Sub

Private Sub Command25_Click()
On Error Resume Next
WebBrowser1.Document.designMode = "off"

End Sub

Private Sub Command26_Click()
 On Error Resume Next
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
    quickpgsource.Text = doc.body.innerHTML
    doc.activeElement.innerHTML = ""
    
End Sub

Private Sub Command27_Click()
Dim S As MSHTML.HTMLInputElement

End Sub

Private Sub Command3_Click()
On Error Resume Next
WebBrowser1.GoBack


End Sub

Private Sub Command4_Click()
On Error Resume Next
WebBrowser1.GoForward


End Sub

Private Sub Command5_Click()
On Error Resume Next
WebBrowser1.Stop

End Sub

Private Sub Command6_Click()
On Error Resume Next
WebBrowser1.Refresh2 1
End Sub

Private Sub Command7_Click()
On Error Resume Next
On Error Resume Next
WebBrowser1.GoHome

WebBrowser1.Visible = True
Picture1.Visible = False

End Sub

Private Sub Command8_Click()
On Error Resume Next
WebBrowser1.GoBack
WebBrowser1.GoBack
WebBrowser1.Visible = True
Picture1.Visible = False

End Sub

Private Sub Command9_Click()
On Error Resume Next
Form4.Show
Form4.text1.Text = GetUrlSource(WebBrowser1.LocationURL)

End Sub

Private Sub fileabout_Click()
Form3.Show

End Sub

Private Sub filex_Click()
On Error Resume Next
WebBrowser1.Stop

WebBrowser1.Visible = False
Form3.Show
Me.Hide

End Sub

Private Sub findmnu1_Click()
WebBrowser1.SetFocus
    SendKeys "^f"
End Sub

Private Sub Form_Load()
On Error Resume Next
WebBrowser1.Navigate ("http://www.google.com")
Combo1.Text = "6,500 - Sticky"
Combo2.Text = "KlanScape"
Combo3.Text = ("Add AFTER current HTML")

Call ResetTreeView
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form3.Show

End Sub

Private Sub gadgx_Click()
 On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
   thedialog.Filter = "Compatible Webpages *.html,*.asp,*.shtml,*.js,*.dhtml) | *.htm;*.html;*.asp;*.shtml;*.js;*.dhtml"
    thedialog.ShowOpen
If thedialog.FileName = "" Then Exit Sub
    WebBrowser1.Navigate (thedialog.FileName)

End Sub

Private Sub gszgdzsg_Click()
On Error Resume Next
Picture10.Visible = True

End Sub

Private Sub htfdxhredsh_Click()
If htfdxhredsh.Checked = True Then GoTo hidepanel1
If htfdxhredsh.Checked = False Then GoTo showpanel
Exit Sub
hidepanel1:
Picture5.Visible = False
htfdxhredsh.Checked = False
sstab2.Left = 0
Form_Resize
Exit Sub
showpanel:
Picture5.Visible = True
htfdxhredsh.Checked = True
sstab2.Left = 3240
Form_Resize
Exit Sub
End Sub

Private Sub inetproperties_Click()
Dim RetVal
    RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)
End Sub

Private Sub inspectsourcex_Click()
On Error Resume Next
Form4.Show
Form4.text1.Text = GetUrlSource(WebBrowser1.LocationURL)

End Sub

Private Sub newwindow1_Click()
On Error Resume Next
    Dim sTargetURL As String
    sTargetURL = WebBrowser1.LocationURL
        Call MakeNewKlanBrowser(sTargetURL, WebBrowser1.Silent)
        If Err Then Err.Clear
    DoEvents
End Sub

Private Sub off11_Click()
On Error Resume Next
WebBrowser1.Document.designMode = "Off"
End Sub

Private Sub on11_Click()
On Error Resume Next
WebBrowser1.Document.designMode = "On"

End Sub

Private Sub Option11_Click()
On Error Resume Next
If Option11.Value = True Then WebBrowser1.Silent = True
End Sub

Private Sub Option9_Click()
On Error Resume Next
If Option9.Value = True Then WebBrowser1.Silent = False
End Sub

Private Sub prop11_Click()
On Error Resume Next
      Call WebBrowser1.ExecWB(OLECMDID_PROPERTIES, OLECMDEXECOPT_DODEFAULT)
    If Err Then Err.Clear
End Sub

Private Sub savepage_Click()
 On Error Resume Next
    Set Web = WebBrowser1
     Set doc = WebBrowser1.Document
    thedialog.Filter = "htm (*.htm) | *.htm"
    thedialog.ShowSave
If Com.FileName = "" Then
    Exit Sub
Else
    Open thedialog.FileName For Output As #1
     Print #1, Web.Document
   Close #1
End If
End Sub

Private Sub Slider1_Change()
On Error Resume Next
If Slider1.Value = 1 Then GoTo xxval1
If Slider1.Value = 2 Then GoTo xxval2
If Slider1.Value = 3 Then GoTo xxval3
If Slider1.Value = 4 Then GoTo xxval4
Exit Sub
xxval1:
Label13.Caption = "Option 1 - Filter after navigation complete. This is one of the fastest filters and will work for most purposes."

Exit Sub
xxval2:
Label13.Caption = "Option 2 - Filter after document complete. This is a little slower and is typically the same results as option 1."
Exit Sub
xxval3:
Label13.Caption = "Option 3 - Filter after progress change. This is the best live filter but also the slowest for large websites."
Exit Sub
xxval4:
Label13.Caption = "Option 4 - Live Filtering (Slowest). This will filter any website in real-time, including entering something on a form."
b = MsgBox("Warning: Live filtering can slow down the program. It depends on your internet connection, amount of filter, and size of site you are visiting. Are you SURE you want to enable Live Filtering?", vbYesNo, "Live Filtering?")

If b = vbNo Then GoTo cancellive1
Timer3.Enabled = True
Picture8.Visible = True
Exit Sub
cancellive1:
Slider1.Value = 1
Label13.Caption = "Option 1 - Filter after navigation complete. This is one of the fastest filters and will work for most purposes."
Exit Sub
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = vbKeyReturn Then WebBrowser1.Navigate text1.Text
        If Err Then Err.Clear
End Sub

Private Sub smalle1_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1)
    If Err Then Err.Clear
    DoEvents
End Sub

Private Sub smallest2_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4)
    If Err Then Err.Clear
    DoEvents
End Sub

Private Sub smallest3_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3)
    If Err Then Err.Clear
    DoEvents
End Sub

Private Sub smallestz_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0)
    If Err Then Err.Clear
    DoEvents
End Sub

Private Sub smallestz1_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2)
    If Err Then Err.Clear
    DoEvents
End Sub

Private Sub Text2_Change()
On Error Resume Next
Timer1.Interval = text1

End Sub




Private Sub Form_Resize()
Dim a
On Error Resume Next

If Picture5.Visible = True Then sstab2.Width = Me.ScaleWidth - Picture5.Width
If Picture5.Visible = True Then sstab2.Height = Me.Height - a - 800
If Picture5.Visible = False Then sstab2.Width = Me.ScaleWidth
If Picture5.Visible = False Then sstab2.Height = Me.Height - a - 800
If Picture5.Visible = True Then WebBrowser1.Width = Me.ScaleWidth - Picture5.Width - 200
If Picture5.Visible = False Then WebBrowser1.Width = Me.ScaleWidth - 200
sstab2.Height = Me.Height - 2200
WebBrowser1.Height = Me.Height - 2670
Picture4.Width = WebBrowser1.Width

Picture4.Height = WebBrowser1.Height
Picture1.Height = WebBrowser1.Height
Picture1.Width = WebBrowser1.Width


SSTab1.Height = Me.ScaleHeight - 1500
Picture5.Height = Me.ScaleHeight - 1500

text1.Height = Me.Height - 3000
text1.Width = Me.Width - 3800
tv.Width = Me.Width - 3800
tv.Height = Me.Height - 4200
list1.Width = Me.Width - 3800

Frame15.Height = Me.Height - 8300
Text12.Height = Me.Height - 8600
End Sub

Private Sub Label8_Click()
WebBrowser1.Stop
Picture4.Visible = False
WebBrowser1.Visible = True

End Sub

Private Sub Text27_Change()
On Error Resume Next

End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Dim custfilt
Timer1.Enabled = False
If Check1.Value = 0 Then Exit Sub

If LCase((WebPageContains(Text7.Text))) = True Then GoTo fix000

custfilt = LCase(Text3.Text)
If WebPageContains(custfilt) = True Then WebBrowser1.Navigate (Text4.Text)
If Check9.Value = 1 Then GoTo mmmmm1
GoTo authxx
mmmmm1:
If WebPageContains(Text2) = True Then GoTo makepopup1
GoTo authxx
makepopup1:
Form2.Show
Form2.Label1.Caption = Text5.Text


GoTo authxx


authxx:
Picture4.Visible = False
WebBrowser1.Visible = True
Exit Sub
fix000:
WebBrowser1.Navigate "about:blank"

Label3.Caption = Text6.Text
Picture4.Visible = False
Picture1.Visible = True
WebBrowser1.Visible = False

End Sub

Private Sub Timer2_Timer()
On Error Resume Next
Timer2.Enabled = False

If Check2.Value = 1 Then GoTo getinfo1
Text8.Text = "Location Name : Skipped Search"
Text9.Text = "Location URL  : Skipped Search"
GoTo getinfo2
getinfo1:
Text8.Text = WebBrowser1.LocationName
Text9.Text = WebBrowser1.LocationURL
getinfo2:
If Check3.Value = 0 Then Label14.Caption = "Contains Scripts : Skipped Search"
If Check3.Value = 0 Then GoTo getinfo3
Label14.Caption = "This site does not contain scripts."
If WebPageContains("script") = True Then Label14.Caption = "This site contains scripts."
getinfo3:
If Check4.Value = 0 Then Label15.Caption = "Contains Swearing : Skipped Search"
If Check4.Value = 0 Then GoTo getinfo4
Label15.Caption = "No known swearing was found."
If WebPageContains("fuck") = True Then Label15.Caption = "This site contains swear words."
If WebPageContains("fag") = True Then Label15.Caption = "This site contains swear words."
If WebPageContains("shit") = True Then Label15.Caption = "This site contains swear words."
If WebPageContains("bitch") = True Then Label15.Caption = "This site contains swear words."
getinfo4:
If Check3.Value = 0 Then Label16.Caption = "Hidden Messages : Skipped Search"
If Check5.Value = 0 Then GoTo getinfo5
Label16.Caption = "Hidden Messages : None Found!"
If WebPageContains("<!--") = True Then Label16.Caption = "Contains hidden messages or comments."
getinfo5:
If Check6.Value = 0 Then Label18.Caption = "Did not search for Viruses."
If Check6.Value = 0 Then GoTo getinfo6
Label18.Caption = "No virus detected."
totalvirusmessage$ = ""
If LCase((WebPageContains("vbs"))) = True Then totalvirusmessages$ = totalvirusmessage$ + " VBS "
If LCase((WebPageContains("exe"))) = True Then totalvirusmessage$ = totalvirusmessage$ + " EXE "
If LCase((WebPageContains("aim:"))) = True Then totalvirusmessage$ = totalvirusmessage$ + " AIM: "
If LCase((WebPageContains("virii"))) = True Then totalvirusmessage$ = totalvirusmessage$ + " VIRII "
If Len(totalvirusmessage) < 2 Then Label18.Caption = "Contains: " + totalvirusmessage$
getinfo6:
Picture6.Visible = False
If Check7.Value = 0 Then Label19.Caption = "Didn't search for custom string."
If Check7.Value = 0 Then Exit Sub
Label19.Caption = "Your custom search was not found!"
If LCase((WebPageContains(Text10.Text))) = True Then Label19.Caption = "Found your custom search!"


End Sub

Private Sub Timer3_Timer()
Timer1.Enabled = True
End Sub

Private Sub Timer5_Timer()
 On Error Resume Next
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
    quickpgsource.Text = doc.body.innerHTML
    
x = doc.cookie
Text14.Text = x
b = doc.security
Label42.Caption = b
Label28.Caption = doc.activeElement
Label49.Caption = doc.activeElement.Id
Label51.Text = doc.activeElement.innerHTML
Text15.Text = doc.domain
Text21.Text = doc.Dir
Text22.Text = doc.fileSize
Text23.Text = doc.fileCreatedDate
Text24.Text = doc.referrer
If Check14.Value = 1 Then doc.activeElement.Style.BorderWidth = Text28.Text
If Check15.Value = 1 Then doc.activeElement.Style.BorderColor = Text29.Text
If Check13.Value = 1 Then doc.activeElement.Style.backgroundColor = Text27.Text
'doc.activeElement.Style.backgroundImage
If Check16.Value = 1 Then doc.activeElement.Style.color = Text30.Text
If Check17.Value = 1 Then doc.activeElement.Style.fontFamily = Text31.Text
If Check18.Value = 1 Then doc.activeElement.Style.fontStyle = Text32.Text
If Check20.Value = 1 Then doc.activeElement.Style.Width = Text33.Text
If Check21.Value = 1 Then doc.activeElement.Style.Height = Text34.Text
Label36.Caption = doc.activeElement.tagName







End Sub

Private Sub Timer6_Timer()
Timer6.Enabled = False
On Error Resume Next
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
    quickpgsource.Text = doc.body.innerHTML

a = Exists_In_String("<!--This page has been filtered by KlanScape-->", quickpgsource)

If a = True Then Exit Sub
If a = False Then GoTo nextpart

'An alternitive method to find out if we have already edited the page
'If WebPageContains("<!--This page has been filtered by KlanScape-->") = False Then WebBrowser1.Navigate (Text4.Text)

Exit Sub
nextpart:
doc.body.innerHTML = Text12.Text & "<!--This page has been filtered by KlanScape--><!-- That message is strictly to prevent a loop bug, by giving the browser a way of knowing we have already edited it-->" & quickpgsource.Text

End Sub

Private Sub Timer7_Timer()
 On Error Resume Next
Dim objLink As HTMLHeaderElement
Dim doc As MSHTML.HTMLDocument
Dim objDocument As MSHTML.HTMLDocument
    Set doc = WebBrowser1.Document
    quickpgsource.Text = doc.body.innerHTML


End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
If Check1.Value = 0 Then Exit Sub
If Slider1.Value = 4 Then Exit Sub
WebBrowser1.Visible = False
Picture4.Visible = True
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next

Me.Caption = "KlanScape Browser : " + WebBrowser1.LocationName
If Slider1.Value = 2 Then Timer1.Enabled = True

Set webdoc = WebBrowser1.Document
Dim Acollection As IHTMLElementCollection
Set Acollection = webdoc.All.tags("a")
For i = 0 To Acollection.Length - 1
    list1.AddItem Acollection.Item(i).toString
Next
Label30.Caption = "This site contains the following" + Str(Acollection.Length) + " links (URLs):"
Set texbody = webdoc.body
Set Texob = texbody.createTextRange()
text1.Text = Texob.Text
Texob.moveToElementText Acollection.Item(3)
newtext2.Text = Texob.Text
Texob.Select
getelement (WebBrowser1.LocationURL)

If Check12.Value = 1 Then Timer6.Enabled = True
Text25.Text = "done"

End Sub

Private Sub WebBrowser1_FileDownload(Cancel As Boolean)
On Error Resume Next
If Option6.Value = True Then Cancel = True
If Option5.Value = True Then GoTo downloadblockal
Exit Sub
downloadblockal:
Form2.Show
Form2.Label1.Caption = "File Download was BLOCKED. To allow it, change your Privacy Settings."
Cancel = True
Exit Sub
End Sub

Private Sub WebBrowser1_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
If Slider1.Value = 1 Then Timer1.Enabled = True
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
On Error Resume Next
Dim bx
If Option3.Value = True Then Cancel = True
If Option3.Value = True Then Label25 = Label25 + 1
If Option3.Value = True Then Exit Sub
If Option1.Value = True Then GoTo asktopopup
If Option2.Value = True Then GoTo popupisok
If Option30.Value = True Then GoTo alertwindow
GoTo openklanwindow
Exit Sub
asktopopup:
bx = MsgBox("Popup Requested, Allow?", vbYesNo, WebBrowser1.LocationURL)
If bx = vbYes Then GoTo popupisok
Cancel = True
Label25 = Label25 + 1
Exit Sub
popupisok:
If Combo2.Text = "KlanScape" Then GoTo openklanwindow
Exit Sub
openklanwindow:
    DoEvents
Dim frmB As New Form1
With frmB
Set ppDisp = .WebBrowser1.object
.WebBrowser1.RegisterAsBrowser = True
.WebBrowser1.Silent = WebBrowser1.Silent
.Show
End With
Set frmB = Nothing
Cancel = False
Exit Sub
alertwindow:
Label25 = Label25 + 1
Form2.Show
Form2.Label1.Caption = "Blocked Popup (To allow, change Privacy Settings)."
Cancel = True

End Sub

Private Sub WebBrowser1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
If Slider1.Value = 3 Then Timer1.Enabled = True
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
On Error Resume Next
picture2.SimpleText = Text
End Sub
Public Sub getelement(URL As String)
On Error Resume Next
' This example uses the built-in KlanScape browser,
' You could also use a new, hidden Internet Explorer to
' get information about a website using the following:
'
 'Dim web As New SHDocVw.InternetExplorer
    
    Set Web = WebBrowser1
   Dim doc As New MSHTML.HTMLDocument
   Dim e As MSHTML.HTMLGenericElement
   Dim a As MSHTML.HTMLAnchorElement
   Dim i As MSHTML.HTMLImg
   Dim t As MSHTML.HTMLTitleElement
   Dim S As MSHTML.HTMLInputElement
   Call ResetTreeView
   Do While Web.Busy
   DoEvents
   Loop
   Set doc = Web.Document
        For Each e In doc.All
      If e.tagName = "A" Then
         Set a = e
         If a.href <> "" Then Call AddToTreeView(a.href, "A", 2)
      ElseIf e.tagName = "IMG" Then
         Set i = e
         If i.src <> "" Then Call AddToTreeView(i.src, "IMG", 3)
      ElseIf e.tagName = "TITLE" Then
         Set t = e
         If t.Text <> "" Then Call AddToTreeView("Page Title: " & t.Text, "Doc", 4)
      ElseIf e.tagName = "INPUT" Then
         Set S = e
         If S.Name <> "" Then Call AddToTreeView("Name (" & S.Name & ")   Size (" & S.Size & ")   Value(" & S.Value & ")", "INPUT", 5)
      End If
   Next
ErrPoint:
   Call CountThem
   Set Web = Nothing
End Sub


Sub CountThem()
    Dim x As Integer
    For x = 1 To 4
        tv.Nodes(x).Text = tv.Nodes(x).Text & " (" & tv.Nodes(x).children & ")"
    Next
End Sub

Sub AddToTreeView(mText As String, mParent As String, Optional mImage As Integer)
    
    On Error GoTo ErrPoint
    Dim tvNode As Node
    If mImage = 0 Then
        Set tvNode = tv.Nodes.Add(mParent, tvwChild, Right(mText, 20), mText)
    Else
        Set tvNode = tv.Nodes.Add(mParent, tvwChild, Right(mText, 20), mText, mImage)
    End If

ErrPoint:

End Sub

Sub ResetTreeView()
    tv.Nodes.Clear
    Dim tvNode As Node
    Set tvNode = tv.Nodes.Add(, tvwparent, "Doc", "Document Elements", 1)
    Set tvNode = tv.Nodes.Add(, tvwparent, "A", "Links", 1)
    Set tvNode = tv.Nodes.Add(, tvwparent, "IMG", "Images", 1)
    Set tvNode = tv.Nodes.Add(, tvwparent, "INPUT", "Inputs", 1)
End Sub
