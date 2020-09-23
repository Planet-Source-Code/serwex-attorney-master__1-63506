VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master  [Users Area]"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   Icon            =   "User_pass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6285
   ScaleMode       =   0  'User
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   210
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   735
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   420
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   105
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame f 
      Enabled         =   0   'False
      Height          =   3480
      Index           =   0
      Left            =   105
      TabIndex        =   14
      Top             =   2415
      Width           =   2745
      Begin VB.OptionButton o 
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   6
         Top             =   1995
         Width           =   2220
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&You have transfered Job?"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2415
         Width           =   2535
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Notes &Here!"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2730
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Messages!"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3045
         Width           =   2535
      End
      Begin VB.OptionButton o 
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   2
         Top             =   735
         Width           =   2220
      End
      Begin VB.OptionButton o 
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   3
         Top             =   1050
         Width           =   2220
      End
      Begin VB.OptionButton o 
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   4
         Top             =   1365
         Width           =   2220
      End
      Begin VB.OptionButton o 
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   5
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Line Line4 
         X1              =   105
         X2              =   2100
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label Label9 
         Caption         =   "Your recent jobs:"
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
         Left            =   105
         TabIndex        =   13
         Top             =   315
         Width           =   1485
      End
   End
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   20
      Top             =   5955
      Width           =   6555
      _ExtentX        =   11562
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7514
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
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   105
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24576001
      CurrentDate     =   36823
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   2940
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   0
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   10398
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   -2147483638
      TabCaption(0)   =   "Sign in"
      TabPicture(0)   =   "User_pass.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Combo6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "f(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "f(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Sign up!"
      TabPicture(1)   =   "User_pass.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label11"
      Tab(1).Control(1)=   "Label10"
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(3)=   "Label7"
      Tab(1).Control(4)=   "Line2"
      Tab(1).Control(5)=   "Line1"
      Tab(1).Control(6)=   "Label6"
      Tab(1).Control(7)=   "Label5"
      Tab(1).Control(8)=   "Label2"
      Tab(1).Control(9)=   "Label1"
      Tab(1).Control(10)=   "Label32"
      Tab(1).Control(11)=   "Label33"
      Tab(1).Control(12)=   "Text7"
      Tab(1).Control(13)=   "Text6"
      Tab(1).Control(14)=   "Command4"
      Tab(1).Control(15)=   "Command2"
      Tab(1).Control(16)=   "Text5"
      Tab(1).Control(17)=   "Text3"
      Tab(1).Control(18)=   "Text1"
      Tab(1).Control(19)=   "Text2"
      Tab(1).Control(20)=   "Text8"
      Tab(1).ControlCount=   21
      Begin VB.TextBox Text8 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74055
         TabIndex        =   64
         Top             =   3885
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
         TabIndex        =   63
         Top             =   3255
         Width           =   2535
      End
      Begin VB.Frame f 
         Caption         =   "Job Style:"
         Enabled         =   0   'False
         Height          =   1380
         Index           =   2
         Left            =   105
         TabIndex        =   37
         Top             =   1260
         Width           =   3375
         Begin VB.ComboBox Combo2 
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   41
            TabStop         =   0   'False
            ToolTipText     =   "List of Inactive Jobs"
            Top             =   945
            Width           =   1800
         End
         Begin VB.ComboBox Combo3 
            BackColor       =   &H00FFC0C0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   1470
            Style           =   2  'Dropdown List
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   "List of Active Jobs."
            Top             =   630
            Width           =   1800
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Inactive File:"
            Height          =   330
            Left            =   105
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   945
            Width           =   1275
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Active File:"
            Height          =   330
            Left            =   105
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   630
            Value           =   -1  'True
            Width           =   1170
         End
         Begin VB.OptionButton Option6 
            Caption         =   "New File"
            Height          =   330
            Left            =   105
            TabIndex        =   38
            TabStop         =   0   'False
            ToolTipText     =   "Start a New Job."
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label26 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Last_Job_ID"
            DataSource      =   "Data1"
            Height          =   225
            Left            =   1680
            TabIndex        =   43
            Top             =   315
            Visible         =   0   'False
            Width           =   1485
         End
      End
      Begin VB.Frame f 
         Caption         =   "File Summary:"
         Enabled         =   0   'False
         Height          =   3060
         Index           =   1
         Left            =   105
         TabIndex        =   30
         Top             =   2730
         Width           =   3375
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Log out"
            Height          =   330
            Left            =   2310
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2625
            Width           =   960
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&Password"
            Height          =   330
            Left            =   1155
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Try change your Password here."
            Top             =   2625
            Width           =   1065
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "&New Job"
            Height          =   330
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2625
            Width           =   960
         End
         Begin VB.Label Label27 
            BorderStyle     =   1  'Fixed Single
            DataField       =   "Job_ID"
            DataSource      =   "Data2"
            Height          =   225
            Left            =   1050
            TabIndex        =   36
            Top             =   315
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label19 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1470
            TabIndex        =   29
            Top             =   1995
            Width           =   1800
         End
         Begin VB.Label Label18 
            Caption         =   "Return Date:"
            Height          =   225
            Left            =   105
            TabIndex        =   26
            Top             =   1995
            Width           =   1275
         End
         Begin VB.Label Label16 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1470
            TabIndex        =   31
            Top             =   1680
            Width           =   1800
         End
         Begin VB.Label Label13 
            Caption         =   "Last access:"
            Height          =   225
            Left            =   105
            TabIndex        =   25
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label Label25 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   2205
            TabIndex        =   35
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label24 
            Caption         =   "Job ID:"
            Height          =   225
            Left            =   105
            TabIndex        =   21
            Top             =   315
            Width           =   1065
         End
         Begin VB.Label Label23 
            Caption         =   "Inactivating Date:"
            Height          =   225
            Left            =   105
            TabIndex        =   27
            Top             =   2310
            Width           =   1275
         End
         Begin VB.Label Label22 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1470
            TabIndex        =   28
            Top             =   2310
            Width           =   1800
         End
         Begin VB.Label Label21 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1470
            TabIndex        =   32
            Top             =   1365
            Width           =   1800
         End
         Begin VB.Label Label20 
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   1470
            TabIndex        =   33
            Top             =   1050
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
            Left            =   1260
            TabIndex        =   34
            Top             =   630
            Width           =   2010
         End
         Begin VB.Label Label15 
            Caption         =   "Registered to:"
            Height          =   225
            Left            =   105
            TabIndex        =   22
            Top             =   630
            Width           =   1065
         End
         Begin VB.Label Label14 
            Caption         =   "Last Modified:"
            Height          =   225
            Left            =   105
            TabIndex        =   24
            Top             =   1365
            Width           =   1065
         End
         Begin VB.Label Label12 
            Caption         =   "Job Created in:"
            Height          =   225
            Left            =   105
            TabIndex        =   23
            Top             =   1050
            Width           =   1275
         End
      End
      Begin VB.ComboBox Combo6 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         ItemData        =   "User_pass.frx":0342
         Left            =   1890
         List            =   "User_pass.frx":0344
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   1590
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1890
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Enter Password for Login."
         Top             =   840
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -74055
         TabIndex        =   62
         Top             =   2835
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73110
         PasswordChar    =   "*"
         TabIndex        =   60
         Top             =   1995
         Width           =   1590
      End
      Begin VB.TextBox Text5 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -73110
         PasswordChar    =   "*"
         TabIndex        =   61
         Top             =   2415
         Width           =   1590
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Submit"
         Height          =   330
         Left            =   -73740
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   4515
         Width           =   1065
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Cancel"
         Height          =   330
         Left            =   -72585
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   4515
         Width           =   1065
      End
      Begin VB.TextBox Text6 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   -73110
         TabIndex        =   58
         Top             =   1155
         Width           =   1590
      End
      Begin VB.TextBox Text7 
         BackColor       =   &H00C0FFC0&
         Height          =   285
         Left            =   -73110
         TabIndex        =   59
         Top             =   1575
         Width           =   1590
      End
      Begin VB.Label Label33 
         Caption         =   "Tel #"
         Height          =   180
         Left            =   -74790
         TabIndex        =   57
         Top             =   3990
         Width           =   540
      End
      Begin VB.Label Label32 
         Caption         =   "Address:"
         Height          =   180
         Left            =   -74790
         TabIndex        =   56
         Top             =   3465
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Password:"
         Height          =   225
         Left            =   105
         TabIndex        =   55
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "User Name:"
         Height          =   225
         Left            =   105
         TabIndex        =   54
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   285
         Left            =   -74790
         TabIndex        =   53
         Top             =   1995
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "E-Mail:"
         Height          =   180
         Left            =   -74790
         TabIndex        =   52
         Top             =   2940
         Width           =   540
      End
      Begin VB.Label Label5 
         Caption         =   "Confirm Password:"
         Height          =   285
         Left            =   -74790
         TabIndex        =   51
         Top             =   2415
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Welcome!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -74790
         TabIndex        =   50
         Top             =   525
         Width           =   3165
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   -74790
         X2              =   -71535
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   -74790
         X2              =   -71535
         Y1              =   4935
         Y2              =   4935
      End
      Begin VB.Label Label7 
         Caption         =   "Note:"
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
         Left            =   -74790
         TabIndex        =   49
         Top             =   5040
         Width           =   3270
      End
      Begin VB.Label Label8 
         Caption         =   "Your superviser must accept your request to sign up successfully."
         Height          =   435
         Left            =   -74790
         TabIndex        =   48
         Top             =   5250
         Width           =   3165
      End
      Begin VB.Label Label10 
         Caption         =   "First Name:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   47
         Top             =   1155
         Width           =   1065
      End
      Begin VB.Label Label11 
         Caption         =   "Last Name:"
         Height          =   225
         Left            =   -74790
         TabIndex        =   46
         Top             =   1575
         Width           =   1380
      End
   End
   Begin VB.Label Label31 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label31"
      DataField       =   "Inactivation_Date"
      DataSource      =   "Data2"
      Height          =   120
      Left            =   0
      TabIndex        =   15
      Top             =   2415
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label30 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Return_Date"
      DataSource      =   "Data2"
      Height          =   330
      Left            =   0
      TabIndex        =   16
      Top             =   1890
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label29 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Last_Name"
      DataSource      =   "Data2"
      Height          =   330
      Left            =   0
      TabIndex        =   17
      Top             =   1470
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label28 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "First_Name"
      DataSource      =   "Data2"
      Height          =   330
      Left            =   0
      TabIndex        =   18
      Top             =   1050
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "Form1"
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
Option Explicit
Dim remember(14) As Variant

Private Sub Combo2_Click()
Dim s, sum_, sd As String
Dim i As Integer
If Combo2.Text <> "" Then
 s = Combo2.Text
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Inactive Jobs\" & Combo2.Text)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
End Sub

Private Sub Combo3_Click()
Dim s, sum_, sd As String
Dim i As Integer
If Combo3.Text <> "" Then
 s = Combo3.Text
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & Combo3.Text)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
End Sub

Private Sub Combo6_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     Text4.SetFocus
  End If
End Sub

Private Sub Command1_Click()
Dim i As Integer
Dim pass As String
Dim db As Database
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data3.Recordset = db.OpenRecordset("select * from Users where Name='" & Combo6.Text & "'", dbOpenDynaset)
If Data3.Recordset.RecordCount = 0 Then
i = MsgBox("You can't change the supervisor Password from here!", vbExclamation, "Attorney Master [Security Unit]")
Data3.Recordset.Close
db.Close
Else
Form8.Show
'Data3.Recordset.Close
'db.Close
End If
End Sub

Private Sub Command2_Click()
Dim intFileNum As Integer
Dim i As Integer
If Text3.Text <> Text5.Text Then
i = MsgBox("Try the Password again!", vbExclamation, "Attorney Master Alert")
Text3.Text = ""
Text5.Text = ""
Text3.SetFocus
Exit Sub
End If
'***************
intFileNum = FreeFile
Open App.Path & "\New Users\" & Text6.Text & " " & Text7.Text & ".txt" For Append As #intFileNum
Print #intFileNum, Text6.Text
Print #intFileNum, Text7.Text
Print #intFileNum, Text3.Text
Print #intFileNum, Text1.Text
Print #intFileNum, Text2.Text
Print #intFileNum, Text8.Text
Close #intFileNum
'***************
Splash.Label3.Caption = "You registered in superviser request list,New User."
Splash.Show
Unload Me
End Sub

Private Sub Command3_Click()
Dim temp1 As Double
Dim intFileNum As Integer
Dim Date_, Time_ As String
Dim i As Integer
If Option6.Value = True Then
i = MsgBox("Continue to Create a new Job?", vbOKCancel + vbQuestion + vbDefaultButton2, "Make me sure...")
  If i = 2 Then Exit Sub
End If
If Label25.Caption = "" Then
i = MsgBox("There is no selected Job!", vbExclamation, "Alert")
Else
'***************
intFileNum = FreeFile
If Option4.Value = True Then
Open App.Path & "\Users\" & Combo6.Text & "\Inactive Jobs\" & Label25.Caption & ".txt" For Append As #intFileNum
Else
Open App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & Label25.Caption & ".txt" For Append As #intFileNum
End If
Time_ = Time
Date_ = Date
Print #intFileNum, "Job opened at:"; Spc(3); Time_; Spc(3); Date_; Spc(3); "By_"; Combo6.Text
Close #intFileNum
'***************
'***************
If Option6.Value = True Then
'************************
  Data2.Recordset.AddNew
  Label27.Caption = Label26.Caption
'  Label19.Caption = Date
  Data2.Recordset.Update
  Label25.Caption = Label26.Caption
'************************
  Data1.Recordset.MoveFirst
  Data1.Recordset.Edit
  temp1 = Label26.Caption
  temp1 = temp1 + 1
  Label26.Caption = temp1
  Data1.Recordset.Update
MkDir (App.Path & "\Jobs\" & Label25.Caption)
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Documents")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Mails")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Mails\Txt")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Mails\Doc")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Notes")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Documents\" & "Movies")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Documents\" & "Soundes")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Documents\" & "Other")
MkDir (App.Path & "\Jobs\" & Label25.Caption & "\Documents\" & "Photoes")
End If
'***************
'Data1.Recordset.Close
'Data1.Database.Close
'Data2.Recordset.Close
'Data2.Database.Close
'***************
Form1.Hide
Form2.Show
End If
End Sub

Private Sub Command4_Click()
SSTab1.Tab = 0
End Sub

Private Sub Command5_Click()
Form5.Show
End Sub

Private Sub Command6_Click()
Form6.Show
End Sub

Private Sub Command7_Click()
Dim i As Integer
Dim c, cc As Integer
For i = 0 To 4
o(i).Caption = ""
Next i
For i = 0 To 14
remember(i) = Empty
Next i
Option6.Value = True
    c = Combo3.ListCount
    cc = Combo2.ListCount
    For i = 1 To c
      Combo3.RemoveItem (0)
    Next i
    For i = 1 To cc
      Combo2.RemoveItem (0)
    Next i
    f(0).Caption = ""
    For i = 0 To 2
        f(i).Enabled = False
    Next i
    Text4.Enabled = True
    Combo6.Enabled = True
    Me.Hide
    Splash.Label3 = "You logout successfully, " & Combo6.Text
    Splash.Show
End Sub

Private Sub Command9_Click()
Form7.Show
End Sub

Private Sub f_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 0 Then
sb.Panels.Item(1) = "See your spcial offers and reminders here."
End If
End Sub

Private Sub Form_Load()
Splash.Hide
Call ShowsubFolderList(App.Path & "\Users")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels.Item(1) = "No actions detected!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
Dim c, cc As Integer
For i = 0 To 4
o(i).Caption = ""
Next i
For i = 0 To 14
remember(i) = Empty
Next i
Option6.Value = True
    c = Combo3.ListCount
    cc = Combo2.ListCount
    For i = 1 To c
      Combo3.RemoveItem (0)
    Next i
    For i = 1 To cc
      Combo2.RemoveItem (0)
    Next i
    f(0).Caption = ""
      For i = 0 To 2
        f(i).Enabled = False
      Next i
    Text4.Enabled = True
    Combo6.Enabled = True
   If Splash.Label3.Caption <> "You registered in superviser request list,New User." Then
      Splash.Label3 = "You logout successfully, " & Combo6.Text
   End If
    Splash.Show
End Sub


Private Sub MonthView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels.Item(1) = "A simple calendar!"
End Sub



Private Sub o_Click(Index As Integer)
Dim s, sum_, sd As String
Dim i As Integer
Command3.Caption = "&Open Job"
Option6.Value = False
Option5.Value = False
Option4.Value = False
Select Case Index
Case 0
If o(Index).Caption <> "" Then
 s = o(Index).Caption
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & o(Index).Caption)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
Case 1
If o(Index).Caption <> "" Then
 s = o(Index).Caption
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & o(Index).Caption)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
Case 2
If o(Index).Caption <> "" Then
 s = o(Index).Caption
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & o(Index).Caption)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
Case 3
If o(Index).Caption <> "" Then
 s = o(Index).Caption
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & o(Index).Caption)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
Case 4
If o(Index).Caption <> "" Then
 s = o(Index).Caption
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & o(Index).Caption)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
End Select
If o(Index).Caption = "" Then
Label25.Caption = ""
Label17.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Label16.Caption = ""
Label19.Caption = ""
Label22.Caption = ""
End If
End Sub

Private Sub Option4_Click()
Dim s, sum_, sd As String
Dim i As Integer
Command3.Caption = "&Open Job"
  For i = 0 To 4
    o(i).Value = False
  Next i
If Combo2.Text <> "" Then
 s = Combo2.Text
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Inactive Jobs\" & Combo2.Text)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
If Combo2.Text = "" Then
Label25.Caption = ""
Label17.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Label16.Caption = ""
Label19.Caption = ""
Label22.Caption = ""
End If
If Option4.Value = True Then
Combo2.Enabled = True
Combo3.Enabled = False
End If
End Sub

Private Sub Option5_Click()
Dim s, sum_, sd As String
Dim i As Integer
Command3.Caption = "&Open Job"
  For i = 0 To 4
    o(i).Value = False
  Next i
If Combo3.Text <> "" Then
 s = Combo3.Text
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = "." Then
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
Label25.Caption = sd
Call ShowFileInfo(App.Path & "\Users\" & Combo6.Text & "\Active Jobs\" & Combo3.Text)
Call ref(sd)
Label17.Caption = Label28.Caption & " " & Label29.Caption
Label19.Caption = Label30.Caption
Label22.Caption = Label31.Caption
End If
If Combo3.Text = "" Then
Label25.Caption = ""
Label17.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Label16.Caption = ""
Label19.Caption = ""
Label22.Caption = ""
End If
If Option5.Value = True Then
Combo3.Enabled = True
Combo2.Enabled = False
End If
End Sub

Private Sub Option6_Click()
Dim i As Integer
Command3.Caption = "&New Job"
  For i = 0 To 4
    o(i).Value = False
  Next i
If Option6.Value = True Then
Label25.Caption = Label26.Caption
Label17.Caption = ""
Label20.Caption = ""
Label21.Caption = ""
Label16.Caption = ""
Label19.Caption = ""
Label22.Caption = ""
Combo2.Enabled = False
Combo3.Enabled = False
End If
End Sub

Private Sub sb_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sb.Panels.Item(1) = "Your status bar!"
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If SSTab1.Caption = "Sign in" Then
       sb.Panels.Item(1) = "Just for user's ,who already has active accounts."
    Else
       sb.Panels.Item(1) = "Sign up if,you don't have any accounts up to now."
    End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
'"""""""""""""""""""""""'
Dim pass As String
Dim db As Database
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data3.Recordset = db.OpenRecordset("select * from Users where Name='" & Combo6.Text & "'", dbOpenDynaset)
If Data3.Recordset.RecordCount = 0 Then
pass = ""
Else
pass = Data3.Recordset("Password").Value
End If
Data3.Recordset.Close
db.Close
'Dim pass As String
'Dim db As Database
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Data1.Recordset = db.OpenRecordset("select * from Fixed_IDS where Superviser_Name='" & Combo6.Text & "'", dbOpenDynaset)
If Data1.Recordset.RecordCount = 0 Then
'pass = ""
Else
pass = Data1.Recordset("Superviser_Password").Value
End If
Data1.Recordset.Close
db.Close
'??????????????????????????//
Dim i, c, cc As Integer
If KeyAscii = 13 Then
If Combo6.Text <> "" Then
 If Text4.Text = pass Then
  '******************
  Data1.DatabaseName = App.Path & "\Shabshab.mdb"
  Data1.RecordSource = "Fixed_IDS"
  Data1.Refresh
  Data1.Recordset.MoveFirst
  '******************
  '******************
  Data2.DatabaseName = App.Path & "\Shabshab.mdb"
  Data2.RecordSource = "File_Summaries"
   Data2.Refresh
'  Data2.Recordset.MoveFirst
  '******************
    f(0).Caption = "Hello, " & Combo6.Text
    For i = 0 To 2
        f(i).Enabled = True
    Next i
    Label25.Caption = Label26.Caption
    Text4.Text = ""
    Text4.Enabled = False
    Combo6.Enabled = False
    c = Combo3.ListCount
    cc = Combo2.ListCount
    For i = 1 To c
      Combo3.RemoveItem (0)
    Next i
    For i = 1 To cc
      Combo2.RemoveItem (0)
    Next i
    Call Active_job_file(App.Path & "\Users\" & Combo6.Text & "\Active Jobs")
    Call Inactive_job_file(App.Path & "\Users\" & Combo6.Text & "\Inactive Jobs")
    Call recent_jobs
 Else
 MsgBox "Invalid Password, try again!", , "Login Failed"
 SendKeys "{End}+{Home}"
 End If
Else
MsgBox "Select your name!"
Combo6.SetFocus
End If
End If
End Sub
Sub Active_job_file(folderspec)
    Dim i As Integer
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
       s = f1.Name
       Combo3.AddItem (s)
       Call ShowFileInfo_plus(folderspec & "\" & s)
    Next
End Sub
Sub Inactive_job_file(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        Combo2.AddItem (s)
    '  Call ShowFileInfo(folderspec & "\" & s)
    Next
End Sub
Sub ShowFileInfo_plus(filespec)
    Dim i, ii As Integer
    Dim st1, sd, st, sum_, temp As String
    Dim fs, f, s
    Dim cdatevar, cdatevar1 As Variant
 '   MsgBox filespec
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.DateLastModified
    st1 = f.Name
    For i = 1 To Len(s)
    sum_ = Mid(s, i, 1)
    If sum_ = " " Then
    ii = i
    Exit For
    Else
    sd = sd & sum_
    End If
    Next i
    st = Mid(s, ii, Len(s))
    For i = 0 To 14 Step 3
    If Not (IsEmpty(remember(i))) Then
'       MsgBox "sss"
       cdatevar = DateValue(remember(i + 1))
       cdatevar1 = DateValue(sd)
       If cdatevar < cdatevar1 Then
         temp = remember(i)
         remember(i) = st1
         st1 = temp
         temp = remember(i + 1)
         remember(i + 1) = sd
         sd = temp
         temp = remember(i + 2)
         remember(i + 2) = st
         st = temp
       End If
       If DateValue(CDate(remember(i + 1))) = DateValue(CDate(sd)) Then
        If TimeValue(remember(i + 2)) < TimeValue(st) Then
         temp = remember(i)
         remember(i) = st1
         st1 = temp
         temp = remember(i + 1)
         remember(i + 1) = sd
         sd = temp
         temp = remember(i + 2)
         remember(i + 2) = st
         st = temp
       End If
       End If
    Else
      remember(i) = st1
      remember(i + 1) = sd
      remember(i + 2) = st
    'MsgBox "Empty"
    'MsgBox i
    'MsgBox remember(i)
    'MsgBox remember(i + 1)
    'MsgBox remember(i + 2)
    Exit For
    End If
    Next i
End Sub
Sub ShowFileInfo_mo(filespec)
    Dim fs, f
'    MsgBox filespec
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    Label21.Caption = f.DateLastModified
'    s = s & "Created: " & f.DateCreated & vbCrLf
'    s = s & "Last Modified: " & f.DateLastModified & vbCrLf
'    MsgBox s, 0, "File Access Info"
End Sub


Sub ShowFileInfo(filespec)
    Dim fs, f
'    MsgBox filespec
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    Label20.Caption = f.DateCreated
    Label16.Caption = f.DateLastModified
    
Dim s, sum_, sd, sd1 As String
Dim i As Integer
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
   sd1 = sd
sd = App.Path & "\Users\" & Combo6.Text & "\Modified\" & sd
If UCase(Dir(sd)) = UCase(sd1) Then
    Call ShowFileInfo_mo(sd)
Else
Label21.Caption = "No Last Modify!"
End If
'    Label16.Caption = f.DateLastAccessed
'    s = s & "Created: " & f.DateCreated & vbCrLf
'    s = s & "Last Modified: " & f.DateLastModified & vbCrLf
'    MsgBox s, 0, "File Access Info"
End Sub

Sub ShowsubFolderList(folderspec)
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 In sf
        s = f1.Name
        Combo6.AddItem (s)
    Next
End Sub

Private Sub recent_jobs()
  Dim i As Integer
  For i = 0 To 4
    o(i).Caption = remember(i * 3)
'  MsgBox o(i).Caption
  Next i
End Sub
Public Sub total()
  '******************
Dim i, c, cc As Integer
  Data1.DatabaseName = App.Path & "\Shabshab.mdb"
  Data1.RecordSource = "Fixed_IDS"
  Data1.Refresh
  Data2.Refresh
  Data1.Recordset.MoveFirst
  '******************
'  Data2.DatabaseName = App.Path & "\Shabshab.mdb"
'  Data2.RecordSource = "File_Summaries"
'  Data2.Refresh
'  Data2.Recordset.MoveFirst
  '******************
'    f(0).Caption = "Hello, " & Combo6.Text
'    For i = 0 To 2
'        f(i).Enabled = True
'    Next i
'    Label25.Caption = Label26.Caption
'    Text4.Text = ""
'    Text4.Enabled = False
'    Combo6.Enabled = False
'If Option6.Value = True Then
'Label25.Caption = Label26.Caption
'End If
If Command3.Caption = "&New Job" Then
Label25.Caption = Label26.Caption
Else
Combo3.Enabled = False
Combo2.Enabled = False
Option4.Value = False
Option5.Value = False
For i = 0 To 4
o(i).Value = False
Next i
End If
    c = Combo3.ListCount
    cc = Combo2.ListCount
    For i = 1 To c
      Combo3.RemoveItem (0)
    Next i
    For i = 1 To cc
      Combo2.RemoveItem (0)
    Next i
       For i = 0 To 14
         remember(i) = Empty
       Next i
    Call Active_job_file(App.Path & "\Users\" & Combo6.Text & "\Active Jobs")
    Call Inactive_job_file(App.Path & "\Users\" & Combo6.Text & "\Inactive Jobs")
    Call recent_jobs
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
Set Data2.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Data2.Recordset.MoveLast
End Sub

