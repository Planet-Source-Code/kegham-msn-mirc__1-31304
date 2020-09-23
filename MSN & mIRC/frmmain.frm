VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Random Irc MSN V 1.0"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   ForeColor       =   &H00000000&
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   5295
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   48
      Text            =   "frmmain.frx":628A
      Top             =   6000
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmmain.frx":6342
      Top             =   5985
      Width           =   1695
   End
   Begin VB.TextBox txtchan 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "/join #"
      Top             =   5640
      Width           =   1215
   End
   Begin VB.TextBox Txtnick 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "/nick"
      Top             =   5280
      Width           =   1215
   End
   Begin VB.TextBox Txtserver 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "/server"
      Top             =   4920
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4095
      Left            =   -15
      TabIndex        =   2
      Top             =   0
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   397
      TabCaption(0)   =   "Servers !"
      TabPicture(0)   =   "frmmain.frx":65A8
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label8"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "List1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "List3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdadserv"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command4"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command5"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Nicknames !"
      TabPicture(1)   =   "frmmain.frx":65C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Image6"
      Tab(1).Control(1)=   "Image2"
      Tab(1).Control(2)=   "Txtscroll"
      Tab(1).Control(3)=   "List2"
      Tab(1).Control(4)=   "cmdadd"
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(7)=   "tmrscroll"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Channels !"
      TabPicture(2)   =   "frmmain.frx":65E0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image7"
      Tab(2).Control(1)=   "Label10"
      Tab(2).Control(2)=   "Label11"
      Tab(2).Control(3)=   "Label12"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "MSN Messenger !"
      TabPicture(3)   =   "frmmain.frx":65FC
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1"
      Tab(3).Control(1)=   "Image3"
      Tab(3).Control(2)=   "Shape1"
      Tab(3).Control(3)=   "Label2"
      Tab(3).Control(4)=   "lblEmail"
      Tab(3).Control(5)=   "Label4"
      Tab(3).Control(6)=   "lblName"
      Tab(3).Control(7)=   "Label6"
      Tab(3).Control(8)=   "lblStatus"
      Tab(3).Control(9)=   "Image4"
      Tab(3).Control(10)=   "SSTab2"
      Tab(3).Control(11)=   "ProgressBar1"
      Tab(3).Control(12)=   "Timer4"
      Tab(3).Control(13)=   "tmrcheck"
      Tab(3).ControlCount=   14
      TabCaption(4)   =   "Credits !"
      TabPicture(4)   =   "frmmain.frx":6618
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Lb3"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Lb4"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Lb5"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Lb6"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "lb7"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "lbi"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "Shape2"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "Text3"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).ControlCount=   8
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   -73740
         Locked          =   -1  'True
         TabIndex        =   53
         Text            =   "http://www.geocities.com/vbdotlb"
         Top             =   3540
         Width           =   2670
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Save server list"
         Height          =   360
         Left            =   225
         Picture         =   "frmmain.frx":6634
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   3315
         Width           =   2145
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&LoadServer"
         Height          =   345
         Left            =   3855
         Picture         =   "frmmain.frx":6DBD
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   3330
         Width           =   975
      End
      Begin VB.CommandButton cmdadserv 
         Caption         =   "&AddServer "
         Height          =   345
         Left            =   2520
         Picture         =   "frmmain.frx":7546
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3330
         Width           =   915
      End
      Begin VB.ListBox List3 
         BackColor       =   &H00404040&
         ForeColor       =   &H00FFFFFF&
         Height          =   2010
         ItemData        =   "frmmain.frx":7CCF
         Left            =   2520
         List            =   "frmmain.frx":7CD1
         TabIndex        =   44
         Top             =   1320
         Width           =   2310
      End
      Begin VB.Timer tmrscroll 
         Interval        =   100
         Left            =   -73440
         Top             =   3120
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Save Nicklist"
         Height          =   375
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1680
         Width           =   3015
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Load saved nicks"
         Height          =   375
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   1080
         Width           =   3015
      End
      Begin VB.CommandButton cmdadd 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Add nick"
         Height          =   375
         Left            =   -72960
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   480
         Width           =   3015
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2760
         ItemData        =   "frmmain.frx":7CD3
         Left            =   -74880
         List            =   "frmmain.frx":7CD5
         TabIndex        =   36
         Top             =   480
         Width           =   1815
      End
      Begin VB.Timer tmrcheck 
         Interval        =   1
         Left            =   -74655
         Top             =   1680
      End
      Begin VB.Timer Timer4 
         Interval        =   75
         Left            =   -70800
         Top             =   1920
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   -73335
         TabIndex        =   33
         Top             =   1860
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1725
         Left            =   -74475
         TabIndex        =   8
         Top             =   2040
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   3043
         _Version        =   393216
         Style           =   1
         TabHeight       =   520
         WordWrap        =   0   'False
         ForeColor       =   10120214
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "My Status !"
         TabPicture(0)   =   "frmmain.frx":7CD7
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdonline"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Cmdaway"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdbuzy"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmphone"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdrightback"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmdlunch"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmdoffline"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "Command13"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Command14"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmdsignout"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "My Inbox !"
         TabPicture(1)   =   "frmmain.frx":7CF3
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblhotmail"
         Tab(1).Control(1)=   "Command8"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "MSN Window !"
         TabPicture(2)   =   "frmmain.frx":7D0F
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdoptions"
         Tab(2).Control(1)=   "cmdaudio"
         Tab(2).ControlCount=   2
         Begin VB.CommandButton cmdsignout 
            Caption         =   "Sign Out"
            Height          =   1005
            Left            =   3375
            TabIndex        =   41
            Top             =   450
            Width           =   495
         End
         Begin VB.CommandButton Command14 
            Caption         =   "Auto Signin"
            Height          =   345
            Left            =   2265
            TabIndex        =   27
            Top             =   1110
            Width           =   1110
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Sign In"
            Height          =   345
            Left            =   2265
            TabIndex        =   26
            Top             =   780
            Width           =   1110
         End
         Begin VB.CommandButton cmdoffline 
            Caption         =   "App Offline"
            Height          =   345
            Left            =   120
            TabIndex        =   25
            Top             =   1110
            Width           =   975
         End
         Begin VB.CommandButton cmdaudio 
            Caption         =   "Audio Tuning "
            Height          =   375
            Left            =   -74880
            Picture         =   "frmmain.frx":7D2B
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   960
            Width           =   3495
         End
         Begin VB.CommandButton cmdoptions 
            Caption         =   "Options"
            Height          =   375
            Left            =   -74880
            Picture         =   "frmmain.frx":19F25
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   480
            Width           =   3495
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Go Inbox !"
            Height          =   735
            Left            =   -74880
            TabIndex        =   15
            Top             =   720
            Width           =   3375
         End
         Begin VB.CommandButton cmdlunch 
            Caption         =   "Out lunch"
            Height          =   345
            Left            =   2265
            TabIndex        =   14
            Top             =   450
            Width           =   1110
         End
         Begin VB.CommandButton cmdrightback 
            Caption         =   "Be right back"
            Height          =   345
            Left            =   1080
            TabIndex        =   13
            Top             =   1110
            Width           =   1200
         End
         Begin VB.CommandButton cmphone 
            Caption         =   "On  phone"
            Height          =   345
            Left            =   1080
            TabIndex        =   12
            Top             =   780
            Width           =   1200
         End
         Begin VB.CommandButton cmdbuzy 
            Caption         =   "Buzy"
            Height          =   345
            Left            =   1080
            TabIndex        =   11
            Top             =   450
            Width           =   1200
         End
         Begin VB.CommandButton Cmdaway 
            Caption         =   "Away"
            Height          =   345
            Left            =   120
            TabIndex        =   10
            Top             =   780
            Width           =   975
         End
         Begin VB.CommandButton cmdonline 
            Caption         =   "Online"
            Height          =   345
            Left            =   120
            TabIndex        =   9
            Top             =   450
            Width           =   975
         End
         Begin VB.Label lblhotmail 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   -74880
            TabIndex        =   16
            Top             =   360
            Width           =   3375
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2010
         ItemData        =   "frmmain.frx":2C11F
         Left            =   240
         List            =   "frmmain.frx":2C171
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   1320
         Width           =   2175
      End
      Begin VB.TextBox Txtscroll 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   -74880
         Locked          =   -1  'True
         TabIndex        =   40
         Top             =   3480
         Width           =   5055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Always save your custom created lists !"
         Height          =   225
         Left            =   1215
         TabIndex        =   52
         Top             =   885
         Width           =   2715
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "kegham_d@hotmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73560
         MouseIcon       =   "frmmain.frx":2C3B0
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   3720
         Width           =   2055
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   $"frmmain.frx":2C6BA
         ForeColor       =   &H00FFFFFF&
         Height          =   2175
         Left            =   -74760
         TabIndex        =   50
         Top             =   1320
         Width           =   4815
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "UNREGISTERED !"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   -74760
         TabIndex        =   49
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Custom servers !"
         Height          =   255
         Left            =   2880
         TabIndex        =   43
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Default Servers !"
         Height          =   255
         Left            =   600
         TabIndex        =   42
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Image Image2 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   840
         Left            =   -71970
         MouseIcon       =   "frmmain.frx":2C979
         Picture         =   "frmmain.frx":2CC83
         Stretch         =   -1  'True
         Top             =   2310
         Width           =   1035
      End
      Begin VB.Image Image6 
         Height          =   3780
         Left            =   -75000
         Picture         =   "frmmain.frx":2CF8D
         Stretch         =   -1  'True
         Top             =   240
         Width           =   5325
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         FillColor       =   &H00008000&
         Height          =   3615
         Left            =   -74730
         Top             =   345
         Width           =   4815
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   75
         TabIndex        =   35
         Top             =   3690
         Width           =   4785
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2040
         TabIndex        =   34
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Image lbi 
         Height          =   480
         Left            =   -72915
         Picture         =   "frmmain.frx":2DCFD
         Stretch         =   -1  'True
         Top             =   2970
         Width           =   795
      End
      Begin VB.Label lb7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Designed and coded "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   615
         Left            =   -74445
         TabIndex        =   32
         Top             =   405
         Width           =   4245
      End
      Begin VB.Label Lb6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "We always do what"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   615
         Left            =   -74655
         TabIndex        =   31
         Top             =   1800
         Width           =   4500
      End
      Begin VB.Label Lb5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "We want !"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   -73650
         TabIndex        =   30
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Lb4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   """ Crackme """
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   570
         Left            =   -73920
         TabIndex        =   29
         Top             =   1320
         Width           =   2730
      End
      Begin VB.Label Lb3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "By"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   510
         Left            =   -73185
         TabIndex        =   28
         Top             =   870
         Width           =   1335
      End
      Begin VB.Image Image4 
         Height          =   435
         Left            =   -70920
         Picture         =   "frmmain.frx":2E02D
         Stretch         =   -1  'True
         Top             =   480
         Width           =   465
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73320
         TabIndex        =   24
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Your  Status  :"
         Height          =   255
         Left            =   -74400
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73320
         TabIndex        =   22
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "My nickname:"
         Height          =   255
         Left            =   -74400
         TabIndex        =   21
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   -73320
         TabIndex        =   20
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Email adress :"
         Height          =   255
         Left            =   -74400
         TabIndex        =   19
         Top             =   840
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Height          =   375
         Left            =   -74400
         Top             =   480
         Width           =   375
      End
      Begin VB.Image Image3 
         Height          =   330
         Left            =   -74400
         Picture         =   "frmmain.frx":342B7
         Stretch         =   -1  'True
         Top             =   480
         Width           =   330
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3855
         Left            =   30
         Picture         =   "frmmain.frx":34EF9
         Stretch         =   -1  'True
         Top             =   225
         Width           =   5310
      End
      Begin VB.Image Image7 
         Height          =   3810
         Left            =   -74955
         Picture         =   "frmmain.frx":394DE
         Stretch         =   -1  'True
         Top             =   270
         Width           =   5280
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "MSN Messenger center commands !"
         Height          =   3525
         Left            =   -74790
         TabIndex        =   7
         Top             =   345
         Width           =   4875
      End
   End
   Begin VB.TextBox txtOutput 
      Height          =   285
      Left            =   120
      LinkTopic       =   "mIRC|command"
      TabIndex        =   0
      Top             =   4560
      Width           =   3015
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
                          
                             ''''''''''''''''''''''''''''
Private Sub cmdadd_Click()   'Do not forget to vote plz '
Dim i As String              ''''''''''''''''''''''''''''
i = InputBox("Enter nickname please")

List2.AddItem i
On Error Resume Next
    Open "c:\nick.ini" For Output As #1
    Print #1,
    
    Close #1
    Exit Sub
End Sub
Private Sub cmdadd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdadd.BackColor = &HFF&

End Sub

Private Sub cmdadd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cmdadd.BackColor = &HFFFFFF
End Sub

Private Sub cmdadserv_Click()
Dim s As String
s = InputBox("Enter server please")

List3.AddItem s
On Error Resume Next
    Open "c:\servers.ini" For Output As #1
    Print #1,
    
Close #1
    Exit Sub
End Sub

Private Sub cmdaudio_Click()
On Error Resume Next
Messenger.MediaWizard 0
End Sub

Private Sub Cmdaway_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_AWAY
lblStatus = "Away"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0 "
End If
End Sub
Private Sub cmdbuzy_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_BUSY
lblStatus = "Buzy"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If
End Sub

Private Sub cmdlunch_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_OUT_TO_LUNCH
lblStatus = "Out to lunch"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If
End Sub

Private Sub cmdoffline_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_INVISIBLE
lblStatus = "Appear Offline"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If
End Sub

Private Sub cmdonline_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_ONLINE
lblStatus = "Online"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If
End Sub

Private Sub cmdoptions_Click()
On Error Resume Next
Messenger.OptionsPages 0, MOPT_GENERAL_PAGE
End Sub

Private Sub cmdrightback_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_BE_RIGHT_BACK
lblStatus = "Be right back"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If
End Sub

Private Sub cmdsignout_Click()
On Error Resume Next
messengerapi.Signout
End Sub

Private Sub cmphone_Click()
On Error GoTo oh:
Timer4.Enabled = True
ProgressBar1.Visible = True
messengerapi.MyStatus = MISTATUS_ON_THE_PHONE
lblStatus = "On Phone"
oh:
If Err.Number = -2147467259 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If

End Sub

Private Sub Command1_Click()
OpenList List2
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFF&
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &HFFFFFF
End Sub

Private Sub Command13_Click()
On Error Resume Next
messengerapi.Signin 0, "", ""
End Sub

Private Sub Command14_Click()
On Error Resume Next
messengerapi.AutoSignin
End Sub

Private Sub Command2_Click()

SaveList List2
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HFF&
End Sub

Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.BackColor = &HFFFFFF

End Sub
Private Sub Command4_Click()
OpenLiss List3
End Sub

Private Sub Command5_Click()
  SaveLiss List3
  
End Sub

Private Sub Command8_Click()
On Error GoTo oh:
messengerapi.OpenInbox
oh:
If Err.Number = -2130705634 Then
MsgBox "Make sure that you are connected", vbInformation, "Random Irc MSN V 1.0"
End If

End Sub

Private Sub Form_Load()
Label5.Caption = App.Path
tmrcheck.Enabled = True
tmrscroll.Enabled = True

End Sub

Private Sub Image2_Click()
MsgBox "Mirc is copyright khaled mardam bey", vbInformation

End Sub

Private Sub Label12_Click()
Dim sendmail
        sendmail = Shell("start.exe mailto:kegham_d@hotmail.com", vbNormalFocus)

End Sub

Private Sub List1_Click()
txtOutput.Text = Txtserver.Text & " " & List1.Text
txtOutput.LinkMode = vbLinkManual

If txtOutput <> "" Then
     
        txtOutput.LinkMode = vbLinkNone
        frmmain.txtOutput.LinkItem = txtOutput.Text
        txtOutput.LinkMode = vbLinkManual
        
        txtOutput.LinkPoke
      Else
MsgBox "Please choose  a command to run!", vbCritical + vbOKOnly, "Error - Input Required!"
    End If

End Sub

Private Sub List2_DblClick()
txtOutput.Text = Txtnick.Text & " " & List2.Text
If List2.Text <> "" Then
     
        txtOutput.LinkMode = vbLinkNone

If txtOutput <> "" Then
     
        txtOutput.LinkMode = vbLinkNone
        frmmain.txtOutput.LinkItem = txtOutput.Text
        txtOutput.LinkMode = vbLinkManual
        txtOutput.LinkPoke
      
    Else
MsgBox "Please choose a command to run!", vbCritical + vbOKOnly, "Error - Input Required!"
    End If
 End If
End Sub

Private Sub List3_Click()
txtOutput.Text = Txtserver.Text & " " & List3.Text
txtOutput.LinkMode = vbLinkManual

If txtOutput <> "" Then
     
        txtOutput.LinkMode = vbLinkNone
        frmmain.txtOutput.LinkItem = txtOutput.Text
        txtOutput.LinkMode = vbLinkManual
        
        txtOutput.LinkPoke
      Else
MsgBox "Please choose  a command to run!", vbCritical + vbOKOnly, "Error - Input Required!"
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
GotoVal = Me.Height / 2      'Form unload animation

For Gointo = 1 To GotoVal
DoEvents
Me.Height = Me.Height - 10
      
If Me.Height <= 11 Then GoTo horiz
    Next Gointo
horiz:
Me.Height = 30
GotoVal = Me.Width / 2
For Gointo = 1 To GotoVal
DoEvents
    Me.Width = Me.Width - 10
    
        If Me.Width <= 11 Then End
        Next Gointo
        
        End
End Sub

Private Sub Timer4_Timer()
On Error Resume Next     'Progress bar timer

ProgressBar1.Value = ProgressBar1.Value + 10
If ProgressBar1.Value = 45 Then

End If
 If ProgressBar1.Value = 100 Then

ProgressBar1.Visible = False
Timer4.Enabled = False

  End If
End Sub

Public Function CheckStatus()
If Messenger.MyStatus = MISTATUS_ONLINE Then
lblStatus.Caption = "Online"
End If
If Messenger.MyStatus = MISTATUS_AWAY Then
lblStatus.Caption = "Away"
End If
If Messenger.MyStatus = MISTATUS_BUSY Then
lblStatus.Caption = "Busy"
End If
If Messenger.MyStatus = MISTATUS_INVISIBLE Then
lblStatus.Caption = "Appear Offline"
End If
If Messenger.MyStatus = MISTATUS_BE_RIGHT_BACK Then
lblStatus.Caption = "Be Right Back"
End If
If Messenger.MyStatus = MISTATUS_ON_THE_PHONE Then
lblStatus.Caption = "On The Phone"
End If
If Messenger.MyStatus = MISTATUS_OUT_TO_LUNCH Then
lblStatus.Caption = "Out To Lunch"
End If

End Function

Private Sub tmrcheck_Timer()
On Error Resume Next
lblhotmail.Caption = "You have (" & Messenger.UnreadEmailCount(MUAFOLDER_INBOX) & " unread emails)"
lblEmail.Caption = messengerapi.MySigninName
lblName.Caption = messengerapi.MyFriendlyName
CheckStatus

End Sub
Private Sub tmrscroll_Timer()
i = i + 1
Select Case i
Case 1
Txtscroll.Text = "D"
Case 2
Txtscroll.Text = "Do"
Case 3
Txtscroll.Text = "Do "
Case 4
Txtscroll.Text = "Do N"
Case 5
Txtscroll.Text = "Do No"
Case 6
Txtscroll.Text = "Do Not"
Case 7
Txtscroll.Text = "Do Not "
Case 8
Txtscroll.Text = "Do Not F"
Case 9
Txtscroll.Text = "Do Not Fo"
Case 10
Txtscroll.Text = "Do Not For"
Case 11
Txtscroll.Text = "Do Not Forg"
Case 12
Txtscroll.Text = "Do Not Forge"
Case 13
Txtscroll.Text = "Do Not Forget"
Case 14
Txtscroll.Text = "Do Not Forget T"
Case 15
Txtscroll.Text = "Do Not Forget To"
Case 16
Txtscroll.Text = "Do Not Forget To "
Case 17
Txtscroll.Text = "Do Not Forget To S"
Case 18
Txtscroll.Text = "Do Not Forget To Sa"
Case 19
Txtscroll.Text = "Do Not Forget To Sav"
Case 20
Txtscroll.Text = "Do Not Forget To Save"
Case 21
Txtscroll.Text = "Do Not Forget To Save "
Case 22
Txtscroll.Text = "Do Not Forget To Save Y"
Case 23
Txtscroll.Text = "Do Not Forget To Save Yo"
Case 24
Txtscroll.Text = "Do Not Forget To Save You"
Case 25
Txtscroll.Text = "Do Not Forget To Save Your"
Case 26
Txtscroll.Text = "Do Not Forget To Save Your "
Case 27
Txtscroll.Text = "Do Not Forget To Save Your N"
Case 28
Txtscroll.Text = "Do Not Forget To Save Your Ni"
Case 29
Txtscroll.Text = "Do Not Forget To Save Your Nic"
Case 30
Txtscroll.Text = "Do Not Forget To Save Your Nick"
Case 31
Txtscroll.Text = "Do Not Forget To Save Your Nickl"
Case 32
Txtscroll.Text = "Do Not Forget To Save Your Nickli"
Case 33
Txtscroll.Text = "Do Not Forget To Save Your Nicklis"
 Case 33
Txtscroll.Text = "Do Not Forget To Save Your Nicklist"
     i = 0
    Case Else
    i = 0
End Select
End Sub
Private Sub Txtscroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MsgBox " Hey ", vbCritical
End Sub
Private Sub SaveList(lst As ListBox)
Dim a As Integer
Dim ff As Integer
Dim Header As String
Dim LCount As Long
Dim numbyte As Byte
ff = FreeFile
Open "C:\nick.ini" For Binary As #1
Header = "SavedList"
Put ff, 1, Header
LCount = lst.ListCount
Put ff, , LCount
For a = 0 To lst.ListCount - 1
    numbyte = Len(lst.List(a))
    Put ff, , numbyte
    Put ff, , lst.List(a)
Next a
Close ff
End Sub
Private Sub SaveLiss(ls As ListBox)
Dim a As Integer
Dim ff As Integer
Dim Header As String
Dim LCount As Long
Dim numbyte As Byte
ff = FreeFile
Open "C:\servers.ini" For Binary As #1
Header = "SavedList"
Put ff, 1, Header
LCount = ls.ListCount
Put ff, , LCount
For a = 0 To ls.ListCount - 1
    numbyte = Len(ls.List(a))
    Put ff, , numbyte
    Put ff, , ls.List(a)
Next a
Close ff
End Sub

Private Sub OpenLiss(lst As ListBox)
Dim a As Integer
Dim ff As Integer
Dim LCount As Long
Dim data As String
Dim numbyte As Byte
ff = FreeFile

Open "C:\Servers.ini" For Binary As #1
    data = "SavedList"
    Get ff, 1, data
    Get ff, , LCount
    If data = "SavedList" Then
        For a = 0 To LCount - 1
            Get ff, , numbyte
            data = String(numbyte, " ")
            Get ff, , data
            lst.AddItem (data)
        Next a
    Else
        MsgBox "File has been modified or not right extention or no servers saved", vbInformation, "Error loading servers list"
        
        
    End If
Close ff
End Sub

Private Sub OpenList(lst As ListBox)
Dim a As Integer
Dim ff As Integer
Dim LCount As Long
Dim data As String
Dim numbyte As Byte

ff = FreeFile

Open "C:\Nick.ini" For Binary As #1
    data = "SavedList"
    Get ff, 1, data
    Get ff, , LCount
    If data = "SavedList" Then
        For a = 0 To LCount - 1
            Get ff, , numbyte
            data = String(numbyte, " ")
            Get ff, , data
            lst.AddItem (data)
        Next a
    Else
        MsgBox "File has been modified or not right extention or no nicknames saved", vbInformation, "Error loading nicklist"
        
        End If
Close ff
End Sub

