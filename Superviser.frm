VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master [Superviser Area]"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "Superviser.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   315
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1995
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   420
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3255
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   525
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame f 
      Caption         =   "File Summary:"
      Height          =   3270
      Index           =   1
      Left            =   4515
      TabIndex        =   31
      Top             =   4620
      Width           =   7365
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2955
         Left            =   3465
         OleObjectBlob   =   "Superviser.frx":030A
         TabIndex        =   47
         Top             =   210
         Width           =   3795
      End
      Begin MSComCtl2.DTPicker dt 
         Height          =   330
         Left            =   1470
         TabIndex        =   11
         Top             =   2835
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36823
      End
      Begin VB.Label l 
         Caption         =   "Return Date:"
         Height          =   225
         Index           =   58
         Left            =   105
         TabIndex        =   46
         Top             =   2940
         Width           =   960
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   3360
         X2              =   3360
         Y1              =   210
         Y2              =   3150
      End
      Begin VB.Label Label2 
         Caption         =   "Investigated By:"
         Height          =   225
         Left            =   105
         TabIndex        =   45
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1470
         TabIndex        =   44
         Top             =   315
         Width           =   1800
      End
      Begin VB.Label Label16 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1470
         TabIndex        =   43
         Top             =   2205
         Width           =   1800
      End
      Begin VB.Label Label13 
         Caption         =   "Last access:"
         Height          =   225
         Left            =   105
         TabIndex        =   42
         Top             =   2205
         Width           =   1275
      End
      Begin VB.Label Label25 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2205
         TabIndex        =   41
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label24 
         Caption         =   "Job ID:"
         Height          =   225
         Left            =   105
         TabIndex        =   40
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label23 
         Caption         =   "Inactivating Date:"
         Height          =   225
         Left            =   105
         TabIndex        =   39
         Top             =   2520
         Width           =   1275
      End
      Begin VB.Label Label22 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1470
         TabIndex        =   38
         Top             =   2520
         Width           =   1800
      End
      Begin VB.Label Label21 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1470
         TabIndex        =   37
         Top             =   1890
         Width           =   1800
      End
      Begin VB.Label Label20 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1470
         TabIndex        =   36
         Top             =   1575
         Width           =   1800
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1470
         TabIndex        =   35
         Top             =   1050
         Width           =   1800
      End
      Begin VB.Label Label15 
         Caption         =   "Registered to:"
         Height          =   225
         Left            =   105
         TabIndex        =   34
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label14 
         Caption         =   "Last Modified:"
         Height          =   225
         Left            =   105
         TabIndex        =   33
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label Label12 
         Caption         =   "Date and Time Created:"
         Height          =   435
         Left            =   105
         TabIndex        =   32
         Top             =   1365
         Width           =   1065
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3900
      Left            =   4515
      TabIndex        =   30
      Top             =   630
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   6879
      _Version        =   393216
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "Users"
      TabPicture(0)   =   "Superviser.frx":2D29
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label19"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label26"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label27"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Text9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Text10"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Command12"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Text11"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Text12"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Command13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Check1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "New Users"
      TabPicture(1)   =   "Superviser.frx":2D45
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text5"
      Tab(1).Control(1)=   "Command11"
      Tab(1).Control(2)=   "Command10"
      Tab(1).Control(3)=   "Text7"
      Tab(1).Control(4)=   "Text6"
      Tab(1).Control(5)=   "Text4"
      Tab(1).Control(6)=   "Text3"
      Tab(1).Control(7)=   "Text2"
      Tab(1).Control(8)=   "Text8"
      Tab(1).Control(9)=   "Label9"
      Tab(1).Control(10)=   "Line2"
      Tab(1).Control(11)=   "Label11"
      Tab(1).Control(12)=   "Label10"
      Tab(1).Control(13)=   "Label8"
      Tab(1).Control(14)=   "Label7"
      Tab(1).Control(15)=   "Label32"
      Tab(1).Control(16)=   "Label33"
      Tab(1).ControlCount=   17
      TabCaption(2)   =   "Transfered Jobs"
      TabPicture(2)   =   "Superviser.frx":2D61
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3"
      Tab(2).Control(1)=   "Command2"
      Tab(2).Control(2)=   "Label37"
      Tab(2).Control(3)=   "Label36"
      Tab(2).Control(4)=   "Label35"
      Tab(2).Control(5)=   "Label34"
      Tab(2).Control(6)=   "Label31"
      Tab(2).Control(7)=   "Label30"
      Tab(2).Control(8)=   "Label29"
      Tab(2).Control(9)=   "Label28"
      Tab(2).ControlCount=   10
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Reject"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -71850
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   3360
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Accept"
         Enabled         =   0   'False
         Height          =   330
         Left            =   -74790
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   3360
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Change Superviser Password?"
         Height          =   330
         Left            =   525
         TabIndex        =   3
         Top             =   525
         Width           =   3165
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancel"
         Enabled         =   0   'False
         Height          =   330
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3255
         Width           =   1485
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Continue..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3255
         Width           =   1590
      End
      Begin VB.TextBox Text12 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2415
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2835
         Width           =   1590
      End
      Begin VB.TextBox Text11 
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2415
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2415
         Width           =   1590
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H0080C0FF&
         Caption         =   "Change Password!"
         Height          =   330
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1890
         Width           =   3480
      End
      Begin VB.TextBox Text10 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1050
         Width           =   1590
      End
      Begin VB.TextBox Text9 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2415
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1470
         Width           =   1590
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H80000014&
         ForeColor       =   &H00FF0000&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73635
         TabIndex        =   18
         Top             =   3360
         Width           =   2115
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Reject"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -71325
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1155
         Width           =   750
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H0080C0FF&
         Caption         =   "&Accept"
         Enabled         =   0   'False
         Height          =   435
         Left            =   -71325
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   630
         Width           =   750
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   -73110
         TabIndex        =   13
         Top             =   1050
         Width           =   1590
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   -73110
         TabIndex        =   12
         Top             =   630
         Width           =   1590
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73110
         TabIndex        =   14
         Top             =   1470
         Width           =   1590
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74055
         TabIndex        =   15
         Top             =   1890
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00C0FFFF&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   -74055
         MultiLine       =   -1  'True
         PasswordChar    =   "*"
         ScrollBars      =   1  'Horizontal
         TabIndex        =   16
         Top             =   2310
         Width           =   2535
      End
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74055
         TabIndex        =   17
         Top             =   2940
         Width           =   2535
      End
      Begin VB.Label Label37 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73530
         TabIndex        =   24
         Top             =   2415
         Width           =   2640
      End
      Begin VB.Label Label36 
         Caption         =   "In Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74370
         TabIndex        =   66
         Top             =   2520
         Width           =   750
      End
      Begin VB.Label Label35 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73530
         TabIndex        =   23
         Top             =   1995
         Width           =   2640
      End
      Begin VB.Label Label34 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73530
         TabIndex        =   22
         Top             =   1575
         Width           =   2640
      End
      Begin VB.Label Label31 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73530
         TabIndex        =   21
         Top             =   1155
         Width           =   2640
      End
      Begin VB.Label Label30 
         Caption         =   "To:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74370
         TabIndex        =   65
         Top             =   2100
         Width           =   540
      End
      Begin VB.Label Label29 
         Caption         =   "From:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74370
         TabIndex        =   64
         Top             =   1680
         Width           =   540
      End
      Begin VB.Label Label28 
         Caption         =   "Job ID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74370
         TabIndex        =   63
         Top             =   1260
         Width           =   750
      End
      Begin VB.Label Label27 
         Caption         =   "Confirm Password:"
         Height          =   285
         Left            =   525
         TabIndex        =   62
         Top             =   2835
         Width           =   1590
      End
      Begin VB.Label Label26 
         Caption         =   "Password:"
         Height          =   285
         Left            =   525
         TabIndex        =   61
         Top             =   2415
         Width           =   960
      End
      Begin VB.Label Label19 
         Caption         =   "User Name:"
         Height          =   225
         Left            =   525
         TabIndex        =   60
         Top             =   1050
         Width           =   1905
      End
      Begin VB.Label Label18 
         Caption         =   "Password:"
         Height          =   285
         Left            =   525
         TabIndex        =   59
         Top             =   1470
         Width           =   1905
      End
      Begin VB.Label Label9 
         Caption         =   "Request Date:"
         Height          =   285
         Left            =   -74790
         TabIndex        =   58
         Top             =   3360
         Width           =   1065
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   -71430
         X2              =   -71430
         Y1              =   630
         Y2              =   3570
      End
      Begin VB.Label Label11 
         Caption         =   "Last Name:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   57
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label10 
         Caption         =   "First Name:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   56
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label Label8 
         Caption         =   "E-Mail:"
         Height          =   180
         Left            =   -74790
         TabIndex        =   55
         Top             =   1995
         Width           =   540
      End
      Begin VB.Label Label7 
         Caption         =   "Password:"
         Height          =   285
         Left            =   -74790
         TabIndex        =   54
         Top             =   1470
         Width           =   960
      End
      Begin VB.Label Label32 
         Caption         =   "Address:"
         Height          =   180
         Left            =   -74790
         TabIndex        =   53
         Top             =   2520
         Width           =   645
      End
      Begin VB.Label Label33 
         Caption         =   "Tel #"
         Height          =   180
         Left            =   -74790
         TabIndex        =   52
         Top             =   3045
         Width           =   540
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Login as User!"
      Height          =   435
      Left            =   1995
      TabIndex        =   1
      Top             =   105
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Logout!"
      Height          =   435
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1800
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   9135
      TabIndex        =   28
      Top             =   2100
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24969217
      CurrentDate     =   36827
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   6840
      Left            =   105
      TabIndex        =   2
      Top             =   1050
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   12065
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   29
      Top             =   8295
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16960
            Text            =   "No actions detected!"
            TextSave        =   "No actions detected!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "www.MixofTix.net"
            TextSave        =   "www.MixofTix.net"
            Object.ToolTipText     =   "Visit our site for complete support ..."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1402
            MinWidth        =   1411
            TextSave        =   "10:00 AM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   1110
      Left            =   9135
      Stretch         =   -1  'True
      ToolTipText     =   "Visit our site 'www.4LawSupport.com'"
      Top             =   630
      Width           =   2685
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Inactivation_Date"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   2415
      TabIndex        =   51
      Top             =   840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label5 
      DataField       =   "Return_Date"
      DataSource      =   "Data1"
      Height          =   120
      Left            =   420
      TabIndex        =   50
      Top             =   630
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label4 
      DataField       =   "Last_Name"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   1365
      TabIndex        =   49
      Top             =   525
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label3 
      DataField       =   "First_Name"
      DataSource      =   "Data1"
      Height          =   225
      Left            =   1890
      TabIndex        =   48
      Top             =   525
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label l 
      Caption         =   "User Diagram:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   27
      Top             =   840
      Width           =   2010
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attorney Master, By_ Shahin Noursalehi - copy 2000-2001
'Contact: admin@MixofTix.net
'Terms of Agreement:
'By using this code, you agree to the following terms...
'1) You may use this code in your own programs (and may compile it into a program and distribute it in compiled format for languages that allow it) freely and with no charge.
'2) You MAY NOT redistribute this code (for example to a web site) without written permission from the original author. Failure to do so is a violation of copyright laws.
'3) You may link to this code from another website, but ONLY if it is not wrapped in a frame.
'4) You will abide by any additional copyright restrictions which the author may have placed in the code or code's description.
'5) Not for commercial use!

Private Sub Check1_Click()
If Check1.Value = 1 Then
Text9.PasswordChar = "*"
Label19.Caption = "Superviser Name:"
Label18.Caption = "Superviser Password:"
Dim db As Database
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data3.Recordset = db.OpenRecordset("select * from Fixed_IDS", dbOpenDynaset)
Text10.Text = Data3.Recordset("Superviser_Name")
Text9.Text = Data3.Recordset("Superviser_Password")
Data3.Recordset.Close
db.Close
Else
Text9.PasswordChar = ""
Text10.Text = ""
Text9.Text = ""
Label19.Caption = "User Name:"
Label18.Caption = "Password:"
End If
End Sub

Private Sub Command1_Click()
Unload Me
Splash.Show
End Sub



Private Sub Command10_Click()
Dim db As Database
'Dim t As Recordset
Dim a As String
Dim bb As String
Dim fs As Object
'*****************
a = Text6.Text & " " & Text7.Text
MkDir (App.Path & "\Users\" & a)
MkDir (App.Path & "\Users\" & a & "\Active Jobs")
MkDir (App.Path & "\Users\" & a & "\Inactive Jobs")
MkDir (App.Path & "\Users\" & a & "\Messages")
MkDir (App.Path & "\Users\" & a & "\Modified")
MkDir (App.Path & "\Users\" & a & "\Notes")
MkDir (App.Path & "\Users\" & a & "\Transfered Job")
bb = App.Path & "\New Users\" & tv.SelectedItem.Text
Set fs = CreateObject("Scripting.FileSystemObject")
a = App.Path & "\Users\" & a & "\"
fs.CopyFile bb, a
Kill bb
'*****************
'    Set Data3.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
'*****************
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data2.Recordset = db.OpenRecordset("select * from Users", dbOpenDynaset)
Data2.Recordset.AddNew
Data2.Recordset("Name") = Text6.Text & " " & Text7.Text
Data2.Recordset("Password") = Text4.Text
Data2.Recordset.Update
Data2.Recordset.Close
db.Close
tv.Nodes.Remove (tv.SelectedItem.Index)
Command10.Enabled = False
Command11.Enabled = False
Text6.Text = ""
Text7.Text = ""
Text4.Text = ""
Text3.Text = ""
Text2.Text = ""
Text8.Text = ""
Text5.Text = ""
Call Form_Load
End Sub

Private Sub Command11_Click()
Kill App.Path & "\New Users\" & tv.SelectedItem.Text
tv.Nodes.Remove (tv.SelectedItem.Index)
Command10.Enabled = False
Command11.Enabled = False
Text6.Text = ""
Text7.Text = ""
Text4.Text = ""
Text3.Text = ""
Text2.Text = ""
Text8.Text = ""
Text5.Text = ""
End Sub

Private Sub Command12_Click()
If Check1.Value = 1 Then
Text11.Enabled = True
Text11.SetFocus
Text12.Enabled = True
Command13.Enabled = True
Command14.Enabled = True
Else
'***
End If
If tv.SelectedItem.Text = "Password" Then
Text11.Enabled = True
Text11.SetFocus
Text12.Enabled = True
Command13.Enabled = True
Command14.Enabled = True
End If
End Sub

Private Sub Command13_Click()
Dim db As Database
If Text11.Text = Text12.Text Then
If Check1.Value = 1 Then
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data3.Recordset = db.OpenRecordset("select * from Fixed_IDS", dbOpenDynaset)
Data3.Recordset.Edit
Data3.Recordset("Superviser_Password") = Text11.Text
Data3.Recordset.Update
Data3.Recordset.Close
db.Close
Check1.Value = 0
Text10.Text = ""
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Text11.Enabled = False
Text12.Enabled = False
Command13.Enabled = False
Command14.Enabled = False
Exit Sub
End If
'Dim n, p As String
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data2.Recordset = db.OpenRecordset("select * from Users where Name='" & tv.SelectedItem.Parent & "'", dbOpenDynaset)
If Data2.Recordset.RecordCount = 0 Then Exit Sub
Data2.Recordset.Edit
Data2.Recordset("Password") = Text11.Text
Data2.Recordset.Update
Data2.Recordset.Close
db.Close
Else
Text11.SetFocus
End If
Text10.Text = ""
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Text11.Enabled = False
Text12.Enabled = False
Command13.Enabled = False
Command14.Enabled = False
End Sub

Private Sub Command14_Click()
Check1.Value = 0
Text10.Text = ""
Text9.Text = ""
Text11.Text = ""
Text12.Text = ""
Text11.Enabled = False
Text12.Enabled = False
Command13.Enabled = False
Command14.Enabled = False
End Sub

Private Sub Command2_Click()
Dim a, bb As String
Dim a1, bb1 As String
Dim fs As Object
bb = App.Path & "\Users\" & Label34.Caption & "\Active Jobs\" & Label31.Caption
a = App.Path & "\Users\" & Label35.Caption & "\Active Jobs\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.MoveFile bb, a
a1 = Mid(Label31.Caption, 1, 9)
'bb = App.Path & "\Users\" & Label34.Caption & "\Modified\" & a1 & "ML.txt"
'a = App.Path & "\Users\" & Label35.Caption & "\Modified\"
'MsgBox bb
'MsgBox a
If (Dir(App.Path & "\Users\" & Label34.Caption & "\Modified\" & a1 & "ML.txt")) = a1 & "ML.txt" Then
bb = App.Path & "\Users\" & Label34.Caption & "\Modified\" & a1 & "ML.txt"
a = App.Path & "\Users\" & Label35.Caption & "\Modified\"
'MsgBox bb
'MsgBox a
fs.MoveFile bb, a
End If
bb = App.Path & "\Transfered Jobs\" & Label31.Caption
a = App.Path & "\Users\" & Label35.Caption & "\Transfered Job\"
'MsgBox bb
'MsgBox a
fs.MoveFile bb, a
'Kill (App.Path & "\Transfered Jobs\" & Label31.Caption)
tv.Nodes.Remove (tv.SelectedItem.Index)
Label31.Caption = ""
Label34.Caption = ""
Label35.Caption = ""
Label37.Caption = ""
Command2.Enabled = False
Command3.Enabled = False
Call Form_Load
End Sub

Private Sub Command3_Click()
Kill (App.Path & "\Transfered Jobs\" & Label31.Caption)
tv.Nodes.Remove (tv.SelectedItem.Index)
Label31.Caption = ""
Label34.Caption = ""
Label35.Caption = ""
Label37.Caption = ""
Command2.Enabled = False
Command3.Enabled = False
End Sub

Private Sub Command4_Click()
Unload Me
Form1.Show
End Sub


Private Sub Command5_Click()

End Sub

Private Sub Form_Load()
Dim nodx As Node
Image1.Picture = LoadPicture(App.Path & "\logo.gif")
tv.Nodes.Clear
tv.LineStyle = tvwRootLines
Set nodx = tv.Nodes.Add(, , "r", "Users")
Set nodx = tv.Nodes.Add(, , "r1", "New Users")
Set nodx = tv.Nodes.Add(, , "r2", "Recent Requests for Transfer Jobs")
Call ShowsubFolderList(App.Path & "\Users")
Call job_fileNew(App.Path & "\New Users")
Call job_trans(App.Path & "\Transfered Jobs")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Splash.Show
End Sub


Private Sub ShowsubFolderList(folderspec)
    Dim nodx As Node
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 In sf
        s = f1.Name
        Set nodx = tv.Nodes.Add("r", tvwChild, s, s)
           Set nodx = tv.Nodes.Add(s, tvwChild, , "Password")
           Set nodx = tv.Nodes.Add(s, tvwChild, s & "J", "Jobs")
              Set nodx = tv.Nodes.Add(s & "J", tvwChild, s & "JA", "Active Jobs")
              Set nodx = tv.Nodes.Add(s & "J", tvwChild, s & "JI", "Inactive Jobs")
           Set nodx = tv.Nodes.Add(s, tvwChild, s & "M", "Active Dates and Times")
           Set nodx = tv.Nodes.Add(s, tvwChild, s & "R", "Requests")
           Call job_file(App.Path & "\users\" & s & "\Active Jobs", s & "JA")
           Call job_file(App.Path & "\users\" & s & "\Inactive Jobs", s & "JI")
    Next
End Sub


Private Sub job_file(folderspec, node_sub)
    Dim nodx As Node
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        Set nodx = tv.Nodes.Add(node_sub, tvwChild, , s)
  '      Call ShowFileInfo(folderspec & "\" & s)
    Next
End Sub

Private Sub ShowFileInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    's = f.DateCreated
    'MsgBox s
    s = s & "Created: " & f.DateCreated & vbCrLf
    s = s & "Last Modified: " & f.DateLastModified & vbCrLf
    MsgBox s, 0, "File Access Info"
End Sub




Private Sub tv_Click()
Dim i As Integer
Dim s, sum_, sd, s_key, ss_key As String
On Error GoTo hey
If tv.SelectedItem.Text = "Users" _
Or tv.SelectedItem.Text = "Recent Requests for Transfer Jobs" _
Or tv.SelectedItem.Text = "New Users" _
Then
'
Else
If tv.SelectedItem.Parent = "Active Jobs" Then
ss_key = tv.SelectedItem.Parent.Key
'MsgBox ss_key
i = Len(ss_key) - 2
'MsgBox i
s_key = Left(ss_key, i)
s = tv.SelectedItem.Text
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label1.Caption = s_key
Label25.Caption = sd
Call ShowFileInfo1(App.Path & "\Users\" & s_key & "\Active Jobs\" & tv.SelectedItem.Text)
Call ref(sd)
Label17.Caption = Label3.Caption & " " & Label4.Caption
If Label5.Caption = Empty Then
dt.Value = Date
Else
dt.Value = Label5.Caption
End If
Label22.Caption = Label6.Caption
End If
If tv.SelectedItem.Parent = "Inactive Jobs" Then
ss_key = tv.SelectedItem.Parent.Key
'MsgBox ss_key
i = Len(ss_key) - 2
'MsgBox i
s_key = Left(ss_key, i)
s = tv.SelectedItem.Text
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Label1.Caption = s_key
Call ShowFileInfo1(App.Path & "\Users\" & s_key & "\Inactive Jobs\" & tv.SelectedItem.Text)
Call ref(sd)
Label17.Caption = Label3.Caption & " " & Label4.Caption
Label22.Caption = Label6.Caption
If Label5.Caption = Empty Then
dt.Value = Date
Else
dt.Value = Label5.Caption
End If
End If
If tv.SelectedItem.Parent = "New Users" Then
SSTab1.Tab = 1
Dim intFileNum As Integer
Dim ss As String
Command10.Enabled = True
Command11.Enabled = True
intFileNum = FreeFile
Open App.Path & "\New Users\" & tv.SelectedItem.Text For Input As #intFileNum
Line Input #intFileNum, ss
Text6.Text = ss
Line Input #intFileNum, ss
Text7.Text = ss
Line Input #intFileNum, ss
Text4.Text = ss
Line Input #intFileNum, ss
Text3.Text = ss
Line Input #intFileNum, ss
Text2.Text = ss
Line Input #intFileNum, ss
Text8.Text = ss
Close #intFileNum
'*******************
Call ShowFileInfo1New(App.Path & "\New Users\" & tv.SelectedItem.Text)
End If
If tv.SelectedItem.Text = "Password" Then
Check1.Value = 0
SSTab1.Tab = 0
Dim db As Database
'Dim n, p As String
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data2.Recordset = db.OpenRecordset("select * from Users where Name='" & tv.SelectedItem.Parent & "'", dbOpenDynaset)
If Data2.Recordset.RecordCount = 0 Then Exit Sub
Text10.Text = Data2.Recordset("Name")
Text9.Text = Data2.Recordset("Password")
Data2.Recordset.Close
db.Close
End If
If tv.SelectedItem.Parent = "Recent Requests for Transfer Jobs" Then
SSTab1.Tab = 2
Dim intFileNum1 As Integer
Dim sss As String
Command2.Enabled = True
Command3.Enabled = True
intFileNum1 = FreeFile
Label31.Caption = tv.SelectedItem.Text
Open App.Path & "\Transfered Jobs\" & tv.SelectedItem.Text For Input As #intFileNum1
Line Input #intFileNum1, sss
Label34.Caption = sss
Line Input #intFileNum1, sss
Label35.Caption = sss
Line Input #intFileNum1, sss
Label37.Caption = sss
Close #intFileNum1
'*******************
End If
End If
Exit Sub
hey:
MsgBox "Press OK to continue"
End Sub

Private Sub ShowFileInfo_mo1(filespec)
    Dim fs, f
'    MsgBox filespec
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    Label21.Caption = f.DateLastModified
'    s = s & "Created: " & f.DateCreated & vbCrLf
'    s = s & "Last Modified: " & f.DateLastModified & vbCrLf
'    MsgBox s, 0, "File Access Info"
End Sub

Private Sub ShowFileInfo1(filespec)
    Dim fs, f
'    MsgBox filespec
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    Label20.Caption = f.DateCreated
    Label16.Caption = f.DateLastModified
Dim s, sum_, sd As String
Dim i As Integer
 Dim temp As String
 s = f.Name
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
   sd = sd & "ML.txt"
 temp = sd
   sd = App.Path & "\Users\" & Label1.Caption & "\Modified\" & sd
    If UCase(Dir(sd)) = UCase(temp) Then
      Call ShowFileInfo_mo1(sd)
    Else
        Label21.Caption = "No Last Modified!"
    End If
'    s = s & "Created: " & f.DateCreated & vbCrLf
'    s = s & "Last Modified: " & f.DateLastModified & vbCrLf
'    MsgBox s, 0, "File Access Info"
End Sub

Private Sub ref(Job_ID_Incoming)
Dim db As Database
Dim sqlq As String
Dim d_b As String
d_b = App.Path & "\Shabshab.mdb"
Set db = OpenDatabase(d_b)
sqlq = "Select * from File_Summaries Where Job_ID = '" _
& Job_ID_Incoming & "'"
'MsgBox sqlq
Set Data1.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Data1.Recordset.MoveLast
End Sub

Private Sub job_fileNew(folderspec)
    Dim nodx As Node
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        Set nodx = tv.Nodes.Add("r1", tvwChild, , s)
    Next
End Sub
Private Sub job_trans(folderspec)
    Dim nodx As Node
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        Set nodx = tv.Nodes.Add("r2", tvwChild, , s)
    Next
End Sub


Private Sub ShowFileInfo1New(filespec)
    Dim fs, f
'    MsgBox filespec
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    Text5.Text = f.DateCreated
 End Sub

