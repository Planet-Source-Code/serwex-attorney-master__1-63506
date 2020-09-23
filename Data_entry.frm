VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C1A8AF28-1257-101B-8FB0-0020AF039CA3}#1.1#0"; "MCI32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master  [Job Desktop]"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "Data_entry.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Data Data6 
      Caption         =   "Data6"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6300
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data5 
      Caption         =   "Data5"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4305
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.Frame f 
      Caption         =   "File Identification:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Index           =   4
      Left            =   9030
      TabIndex        =   277
      Top             =   5670
      Width           =   2850
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Close Job"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   298
         Top             =   1680
         Width           =   2640
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "&Transfer Job to..."
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   297
         Top             =   1365
         Width           =   2640
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "Return_Date"
         DataSource      =   "Data2"
         Height          =   330
         Index           =   18
         Left            =   1050
         TabIndex        =   296
         Top             =   945
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36823
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1050
         TabIndex        =   295
         Top             =   630
         Width           =   1695
      End
      Begin VB.Label l 
         Caption         =   "Last modify:"
         Height          =   225
         Index           =   57
         Left            =   105
         TabIndex        =   301
         Top             =   735
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Return Date:"
         Height          =   225
         Index           =   58
         Left            =   105
         TabIndex        =   300
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label l 
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   330
         Index           =   150
         Left            =   1050
         TabIndex        =   294
         Top             =   315
         Width           =   1695
      End
      Begin VB.Label l 
         Caption         =   "Job ID:"
         Height          =   225
         Index           =   56
         Left            =   105
         TabIndex        =   299
         Top             =   420
         Width           =   855
      End
   End
   Begin VB.Frame f 
      Caption         =   "Related Mails"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5475
      Index           =   5
      Left            =   9030
      TabIndex        =   276
      Top             =   105
      Width           =   2850
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Other..."
         Height          =   225
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   567
         Top             =   4725
         Width           =   2640
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Mail Manager"
         Height          =   330
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   293
         Top             =   5040
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Clear &All"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   292
         Top             =   5040
         Width           =   1275
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   13
         Left            =   105
         TabIndex        =   291
         Top             =   4410
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   12
         Left            =   105
         TabIndex        =   290
         Top             =   4095
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   11
         Left            =   105
         TabIndex        =   289
         Top             =   3780
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   10
         Left            =   105
         TabIndex        =   288
         Top             =   3465
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   9
         Left            =   105
         TabIndex        =   287
         Top             =   3150
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   8
         Left            =   105
         TabIndex        =   286
         Top             =   2835
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Caption         =   "Click on mail manager!"
         Height          =   225
         Index           =   7
         Left            =   105
         TabIndex        =   285
         Top             =   2520
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Caption         =   "Not active in this version."
         Height          =   225
         Index           =   6
         Left            =   105
         TabIndex        =   284
         Top             =   2205
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Caption         =   "short cuts to the Mail samples."
         Height          =   225
         Index           =   5
         Left            =   105
         TabIndex        =   283
         Top             =   1890
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Caption         =   "This section contains many"
         Height          =   225
         Index           =   4
         Left            =   105
         TabIndex        =   282
         Top             =   1575
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   3
         Left            =   105
         TabIndex        =   281
         Top             =   1260
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   2
         Left            =   105
         TabIndex        =   280
         Top             =   945
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   1
         Left            =   105
         TabIndex        =   279
         Top             =   630
         Width           =   2430
      End
      Begin VB.CheckBox c 
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   278
         Top             =   315
         Width           =   2430
      End
   End
   Begin TabDlg.SSTab st1 
      Height          =   7695
      Left            =   105
      TabIndex        =   302
      Top             =   210
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   13573
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   12
      Tab             =   11
      TabsPerRow      =   12
      TabHeight       =   520
      ShowFocusRect   =   0   'False
      BackColor       =   -2147483644
      ForeColor       =   -2147483640
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "P. 12"
      TabPicture(0)   =   "Data_entry.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "co(7)"
      Tab(0).Control(1)=   "tv(1)"
      Tab(0).Control(2)=   "Line57"
      Tab(0).Control(3)=   "l(147)"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "P. 11"
      TabPicture(1)   =   "Data_entry.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pi"
      Tab(1).Control(1)=   "co(6)"
      Tab(1).Control(2)=   "co(4)"
      Tab(1).Control(3)=   "co(5)"
      Tab(1).Control(4)=   "co(11)"
      Tab(1).Control(5)=   "mm"
      Tab(1).Control(6)=   "co(3)"
      Tab(1).Control(7)=   "co(2)"
      Tab(1).Control(8)=   "co(1)"
      Tab(1).Control(9)=   "co(0)"
      Tab(1).Control(10)=   "tv(0)"
      Tab(1).Control(11)=   "cd"
      Tab(1).Control(12)=   "p"
      Tab(1).Control(13)=   "l(149)"
      Tab(1).Control(14)=   "Line58"
      Tab(1).Control(15)=   "l(148)"
      Tab(1).Control(16)=   "Line56"
      Tab(1).Control(17)=   "l(146)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "P. 10"
      TabPicture(2)   =   "Data_entry.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "co(10)"
      Tab(2).Control(1)=   "co(9)"
      Tab(2).Control(2)=   "co(8)"
      Tab(2).Control(3)=   "Text4"
      Tab(2).Control(4)=   "l(145)"
      Tab(2).Control(5)=   "Line55"
      Tab(2).Control(6)=   "l(144)"
      Tab(2).ControlCount=   7
      TabCaption(3)   =   "P. 9"
      TabPicture(3)   =   "Data_entry.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "f(21)"
      Tab(3).Control(1)=   "t(183)"
      Tab(3).Control(2)=   "t(184)"
      Tab(3).Control(3)=   "t(185)"
      Tab(3).Control(4)=   "t(186)"
      Tab(3).Control(5)=   "t(187)"
      Tab(3).Control(6)=   "t(182)"
      Tab(3).Control(7)=   "co(15)"
      Tab(3).Control(8)=   "com(1)"
      Tab(3).Control(9)=   "com(0)"
      Tab(3).Control(10)=   "t(177)"
      Tab(3).Control(11)=   "t(178)"
      Tab(3).Control(12)=   "t(179)"
      Tab(3).Control(13)=   "t(180)"
      Tab(3).Control(14)=   "t(181)"
      Tab(3).Control(15)=   "t(176)"
      Tab(3).Control(16)=   "co(16)"
      Tab(3).Control(17)=   "co(14)"
      Tab(3).Control(18)=   "co(13)"
      Tab(3).Control(19)=   "Data1"
      Tab(3).Control(20)=   "DBGrid1"
      Tab(3).Control(21)=   "c(148)"
      Tab(3).Control(22)=   "c(149)"
      Tab(3).Control(23)=   "c(150)"
      Tab(3).Control(24)=   "c(151)"
      Tab(3).Control(25)=   "c(152)"
      Tab(3).Control(26)=   "c(147)"
      Tab(3).Control(27)=   "MSChart1"
      Tab(3).Control(28)=   "l(168)"
      Tab(3).Control(29)=   "Line65"
      Tab(3).Control(30)=   "l(167)"
      Tab(3).Control(31)=   "l(165)"
      Tab(3).Control(32)=   "Line64"
      Tab(3).Control(33)=   "li(26)"
      Tab(3).Control(34)=   "li(25)"
      Tab(3).Control(35)=   "l(163)"
      Tab(3).Control(36)=   "l(162)"
      Tab(3).Control(37)=   "l(161)"
      Tab(3).Control(38)=   "l(160)"
      Tab(3).Control(39)=   "li(23)"
      Tab(3).Control(40)=   "li(21)"
      Tab(3).Control(41)=   "Line60"
      Tab(3).Control(42)=   "l(159)"
      Tab(3).Control(43)=   "l(158)"
      Tab(3).Control(44)=   "l(157)"
      Tab(3).Control(45)=   "l(156)"
      Tab(3).Control(46)=   "l(155)"
      Tab(3).Control(47)=   "l(154)"
      Tab(3).Control(48)=   "l(153)"
      Tab(3).Control(49)=   "Line59"
      Tab(3).Control(50)=   "l(152)"
      Tab(3).Control(51)=   "l(151)"
      Tab(3).Control(52)=   "li(19)"
      Tab(3).Control(53)=   "li(24)"
      Tab(3).Control(54)=   "li(22)"
      Tab(3).Control(55)=   "li(20)"
      Tab(3).Control(56)=   "Line54"
      Tab(3).Control(57)=   "l(143)"
      Tab(3).ControlCount=   58
      TabCaption(4)   =   "P. 8"
      TabPicture(4)   =   "Data_entry.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "l(136)"
      Tab(4).Control(1)=   "li(0)"
      Tab(4).Control(2)=   "l(137)"
      Tab(4).Control(3)=   "l(138)"
      Tab(4).Control(4)=   "l(139)"
      Tab(4).Control(5)=   "li(1)"
      Tab(4).Control(6)=   "Line38"
      Tab(4).Control(7)=   "li(18)"
      Tab(4).Control(8)=   "l(140)"
      Tab(4).Control(9)=   "l(141)"
      Tab(4).Control(10)=   "l(142)"
      Tab(4).Control(11)=   "Line51"
      Tab(4).Control(12)=   "Line52"
      Tab(4).Control(13)=   "t(161)"
      Tab(4).Control(14)=   "t(162)"
      Tab(4).Control(15)=   "t(163)"
      Tab(4).Control(16)=   "f(19)"
      Tab(4).Control(17)=   "f(20)"
      Tab(4).Control(18)=   "t(173)"
      Tab(4).Control(19)=   "t(174)"
      Tab(4).Control(20)=   "t(175)"
      Tab(4).Control(21)=   "Data7"
      Tab(4).Control(22)=   "Data8"
      Tab(4).Control(23)=   "Data9"
      Tab(4).Control(24)=   "Data10"
      Tab(4).ControlCount=   25
      TabCaption(5)   =   "P. 7"
      TabPicture(5)   =   "Data_entry.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "l(125)"
      Tab(5).Control(1)=   "l(124)"
      Tab(5).Control(2)=   "l(123)"
      Tab(5).Control(3)=   "l(122)"
      Tab(5).Control(4)=   "l(121)"
      Tab(5).Control(5)=   "l(120)"
      Tab(5).Control(6)=   "l(119)"
      Tab(5).Control(7)=   "l(118)"
      Tab(5).Control(8)=   "l(117)"
      Tab(5).Control(9)=   "l(116)"
      Tab(5).Control(10)=   "l(115)"
      Tab(5).Control(11)=   "l(114)"
      Tab(5).Control(12)=   "l(113)"
      Tab(5).Control(13)=   "l(112)"
      Tab(5).Control(14)=   "l(111)"
      Tab(5).Control(15)=   "l(110)"
      Tab(5).Control(16)=   "l(109)"
      Tab(5).Control(17)=   "l(108)"
      Tab(5).Control(18)=   "l(107)"
      Tab(5).Control(19)=   "l(95)"
      Tab(5).Control(20)=   "l(94)"
      Tab(5).Control(21)=   "l(126)"
      Tab(5).Control(22)=   "l(127)"
      Tab(5).Control(23)=   "l(128)"
      Tab(5).Control(24)=   "l(129)"
      Tab(5).Control(25)=   "l(130)"
      Tab(5).Control(26)=   "l(131)"
      Tab(5).Control(27)=   "l(132)"
      Tab(5).Control(28)=   "l(133)"
      Tab(5).Control(29)=   "li(3)"
      Tab(5).Control(30)=   "li(4)"
      Tab(5).Control(31)=   "li(2)"
      Tab(5).Control(32)=   "li(5)"
      Tab(5).Control(33)=   "li(6)"
      Tab(5).Control(34)=   "dt(9)"
      Tab(5).Control(35)=   "dt(8)"
      Tab(5).Control(36)=   "dt(7)"
      Tab(5).Control(37)=   "dt(6)"
      Tab(5).Control(38)=   "dt(5)"
      Tab(5).Control(39)=   "dt(4)"
      Tab(5).Control(40)=   "t(151)"
      Tab(5).Control(41)=   "t(152)"
      Tab(5).Control(42)=   "t(153)"
      Tab(5).Control(43)=   "t(154)"
      Tab(5).Control(44)=   "t(155)"
      Tab(5).Control(45)=   "t(146)"
      Tab(5).Control(46)=   "t(147)"
      Tab(5).Control(47)=   "t(148)"
      Tab(5).Control(48)=   "t(149)"
      Tab(5).Control(49)=   "t(150)"
      Tab(5).Control(50)=   "t(145)"
      Tab(5).Control(51)=   "t(144)"
      Tab(5).Control(52)=   "c(131)"
      Tab(5).Control(53)=   "t(160)"
      Tab(5).Control(54)=   "t(159)"
      Tab(5).Control(55)=   "t(158)"
      Tab(5).Control(56)=   "t(157)"
      Tab(5).Control(57)=   "t(156)"
      Tab(5).ControlCount=   58
      TabCaption(6)   =   "P. 6"
      TabPicture(6)   =   "Data_entry.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "l(106)"
      Tab(6).Control(1)=   "li(9)"
      Tab(6).Control(2)=   "li(8)"
      Tab(6).Control(3)=   "Label7"
      Tab(6).Control(4)=   "l(99)"
      Tab(6).Control(5)=   "l(98)"
      Tab(6).Control(6)=   "l(97)"
      Tab(6).Control(7)=   "l(96)"
      Tab(6).Control(8)=   "l(92)"
      Tab(6).Control(9)=   "Line61"
      Tab(6).Control(10)=   "li(7)"
      Tab(6).Control(11)=   "t(143)"
      Tab(6).Control(12)=   "t(142)"
      Tab(6).Control(13)=   "c(130)"
      Tab(6).Control(14)=   "c(129)"
      Tab(6).Control(15)=   "c(128)"
      Tab(6).Control(16)=   "f(18)"
      Tab(6).Control(17)=   "f(17)"
      Tab(6).Control(18)=   "t(137)"
      Tab(6).Control(19)=   "t(136)"
      Tab(6).Control(20)=   "t(135)"
      Tab(6).Control(21)=   "t(134)"
      Tab(6).ControlCount=   22
      TabCaption(7)   =   "P. 5"
      TabPicture(7)   =   "Data_entry.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "li(12)"
      Tab(7).Control(1)=   "li(13)"
      Tab(7).Control(2)=   "l(88)"
      Tab(7).Control(3)=   "l(87)"
      Tab(7).Control(4)=   "li(14)"
      Tab(7).Control(5)=   "li(17)"
      Tab(7).Control(6)=   "l(86)"
      Tab(7).Control(7)=   "l(85)"
      Tab(7).Control(8)=   "l(84)"
      Tab(7).Control(9)=   "l(83)"
      Tab(7).Control(10)=   "li(10)"
      Tab(7).Control(11)=   "li(11)"
      Tab(7).Control(12)=   "l(82)"
      Tab(7).Control(13)=   "l(81)"
      Tab(7).Control(14)=   "l(80)"
      Tab(7).Control(15)=   "l(79)"
      Tab(7).Control(16)=   "l(78)"
      Tab(7).Control(17)=   "l(77)"
      Tab(7).Control(18)=   "l(76)"
      Tab(7).Control(19)=   "l(75)"
      Tab(7).Control(20)=   "Label3"
      Tab(7).Control(21)=   "Label1"
      Tab(7).Control(22)=   "l(74)"
      Tab(7).Control(23)=   "Line50"
      Tab(7).Control(24)=   "Line53"
      Tab(7).Control(25)=   "li(16)"
      Tab(7).Control(26)=   "li(15)"
      Tab(7).Control(27)=   "Label17"
      Tab(7).Control(28)=   "Line71"
      Tab(7).Control(29)=   "Line72"
      Tab(7).Control(30)=   "Label18"
      Tab(7).Control(31)=   "Line73"
      Tab(7).Control(32)=   "Label22"
      Tab(7).Control(33)=   "Label23"
      Tab(7).Control(34)=   "Label29"
      Tab(7).Control(35)=   "f(13)"
      Tab(7).Control(36)=   "c(121)"
      Tab(7).Control(37)=   "t(130)"
      Tab(7).Control(38)=   "t(128)"
      Tab(7).Control(39)=   "t(129)"
      Tab(7).Control(40)=   "t(126)"
      Tab(7).Control(41)=   "t(127)"
      Tab(7).Control(42)=   "t(125)"
      Tab(7).Control(43)=   "t(120)"
      Tab(7).Control(44)=   "t(118)"
      Tab(7).Control(45)=   "t(119)"
      Tab(7).Control(46)=   "t(121)"
      Tab(7).Control(47)=   "t(124)"
      Tab(7).Control(48)=   "t(123)"
      Tab(7).Control(49)=   "t(122)"
      Tab(7).Control(50)=   "c(34)"
      Tab(7).Control(51)=   "c(35)"
      Tab(7).Control(52)=   "c(36)"
      Tab(7).Control(53)=   "c(37)"
      Tab(7).Control(54)=   "c(38)"
      Tab(7).Control(55)=   "c(39)"
      Tab(7).Control(56)=   "c(40)"
      Tab(7).Control(57)=   "c(41)"
      Tab(7).Control(58)=   "c(42)"
      Tab(7).Control(59)=   "c(43)"
      Tab(7).Control(60)=   "c(44)"
      Tab(7).Control(61)=   "c(45)"
      Tab(7).Control(62)=   "c(46)"
      Tab(7).Control(63)=   "c(47)"
      Tab(7).Control(64)=   "c(49)"
      Tab(7).Control(65)=   "c(48)"
      Tab(7).ControlCount=   66
      TabCaption(8)   =   "P. 4"
      TabPicture(8)   =   "Data_entry.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "l(73)"
      Tab(8).Control(1)=   "l(72)"
      Tab(8).Control(2)=   "l(71)"
      Tab(8).Control(3)=   "l(69)"
      Tab(8).Control(4)=   "Line48"
      Tab(8).Control(5)=   "l(68)"
      Tab(8).Control(6)=   "l(66)"
      Tab(8).Control(7)=   "l(67)"
      Tab(8).Control(8)=   "Line47"
      Tab(8).Control(9)=   "Line46"
      Tab(8).Control(10)=   "l(65)"
      Tab(8).Control(11)=   "l(64)"
      Tab(8).Control(12)=   "l(63)"
      Tab(8).Control(13)=   "l(62)"
      Tab(8).Control(14)=   "l(61)"
      Tab(8).Control(15)=   "l(60)"
      Tab(8).Control(16)=   "l(59)"
      Tab(8).Control(17)=   "Label24"
      Tab(8).Control(18)=   "l(43)"
      Tab(8).Control(19)=   "Line43"
      Tab(8).Control(20)=   "l(42)"
      Tab(8).Control(21)=   "Line44"
      Tab(8).Control(22)=   "Line41"
      Tab(8).Control(23)=   "Line49"
      Tab(8).Control(24)=   "Line45"
      Tab(8).Control(25)=   "l(170)"
      Tab(8).Control(26)=   "Label9"
      Tab(8).Control(27)=   "Line67"
      Tab(8).Control(28)=   "Line68"
      Tab(8).Control(29)=   "Label15"
      Tab(8).Control(30)=   "Label16"
      Tab(8).Control(31)=   "Line69"
      Tab(8).Control(32)=   "Line70"
      Tab(8).Control(33)=   "c(120)"
      Tab(8).Control(34)=   "c(119)"
      Tab(8).Control(35)=   "c(118)"
      Tab(8).Control(36)=   "t(112)"
      Tab(8).Control(37)=   "t(113)"
      Tab(8).Control(38)=   "t(114)"
      Tab(8).Control(39)=   "t(117)"
      Tab(8).Control(40)=   "t(116)"
      Tab(8).Control(41)=   "t(115)"
      Tab(8).Control(42)=   "c(117)"
      Tab(8).Control(43)=   "f(8)"
      Tab(8).Control(44)=   "t(111)"
      Tab(8).Control(45)=   "t(110)"
      Tab(8).Control(46)=   "t(109)"
      Tab(8).Control(47)=   "t(108)"
      Tab(8).Control(48)=   "t(107)"
      Tab(8).Control(49)=   "t(106)"
      Tab(8).Control(50)=   "t(105)"
      Tab(8).Control(51)=   "c(111)"
      Tab(8).Control(52)=   "c(110)"
      Tab(8).Control(53)=   "c(25)"
      Tab(8).Control(54)=   "c(26)"
      Tab(8).Control(55)=   "c(27)"
      Tab(8).Control(56)=   "c(28)"
      Tab(8).Control(57)=   "c(29)"
      Tab(8).Control(58)=   "c(30)"
      Tab(8).Control(59)=   "c(31)"
      Tab(8).Control(60)=   "c(32)"
      Tab(8).Control(61)=   "c(33)"
      Tab(8).ControlCount=   62
      TabCaption(9)   =   "P. 3"
      TabPicture(9)   =   "Data_entry.frx":0406
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "l(93)"
      Tab(9).Control(1)=   "l(70)"
      Tab(9).Control(2)=   "Line40"
      Tab(9).Control(3)=   "l(48)"
      Tab(9).Control(4)=   "l(47)"
      Tab(9).Control(5)=   "l(46)"
      Tab(9).Control(6)=   "Line42"
      Tab(9).Control(7)=   "l(44)"
      Tab(9).Control(8)=   "l(45)"
      Tab(9).Control(9)=   "Line39"
      Tab(9).Control(10)=   "Label21"
      Tab(9).Control(11)=   "Label19"
      Tab(9).Control(12)=   "l(49)"
      Tab(9).Control(13)=   "l(52)"
      Tab(9).Control(14)=   "l(53)"
      Tab(9).Control(15)=   "Line37"
      Tab(9).Control(16)=   "l(51)"
      Tab(9).Control(17)=   "Line35"
      Tab(9).Control(18)=   "Line34"
      Tab(9).Control(19)=   "l(41)"
      Tab(9).Control(20)=   "l(40)"
      Tab(9).Control(21)=   "l(39)"
      Tab(9).Control(22)=   "l(38)"
      Tab(9).Control(23)=   "l(37)"
      Tab(9).Control(24)=   "l(36)"
      Tab(9).Control(25)=   "Label14"
      Tab(9).Control(26)=   "Label13"
      Tab(9).Control(27)=   "Label10"
      Tab(9).Control(28)=   "l(54)"
      Tab(9).Control(29)=   "l(55)"
      Tab(9).Control(30)=   "l(50)"
      Tab(9).Control(31)=   "Label6"
      Tab(9).Control(32)=   "Line32"
      Tab(9).Control(33)=   "Line33"
      Tab(9).Control(34)=   "Line36"
      Tab(9).Control(35)=   "Line21"
      Tab(9).Control(36)=   "Line66"
      Tab(9).Control(37)=   "l(169)"
      Tab(9).Control(38)=   "Label2"
      Tab(9).Control(39)=   "dt(3)"
      Tab(9).Control(40)=   "t(104)"
      Tab(9).Control(41)=   "t(103)"
      Tab(9).Control(42)=   "t(102)"
      Tab(9).Control(43)=   "t(100)"
      Tab(9).Control(44)=   "t(101)"
      Tab(9).Control(45)=   "c(109)"
      Tab(9).Control(46)=   "t(99)"
      Tab(9).Control(47)=   "t(98)"
      Tab(9).Control(48)=   "c(108)"
      Tab(9).Control(49)=   "t(97)"
      Tab(9).Control(50)=   "t(96)"
      Tab(9).Control(51)=   "t(94)"
      Tab(9).Control(52)=   "t(91)"
      Tab(9).Control(53)=   "t(92)"
      Tab(9).Control(54)=   "t(93)"
      Tab(9).Control(55)=   "t(95)"
      Tab(9).Control(56)=   "t(90)"
      Tab(9).Control(57)=   "t(89)"
      Tab(9).Control(58)=   "t(88)"
      Tab(9).Control(59)=   "c(15)"
      Tab(9).Control(60)=   "c(14)"
      Tab(9).Control(61)=   "check1(0)"
      Tab(9).Control(62)=   "c(16)"
      Tab(9).Control(63)=   "c(17)"
      Tab(9).Control(64)=   "c(18)"
      Tab(9).Control(65)=   "c(19)"
      Tab(9).Control(66)=   "c(24)"
      Tab(9).Control(67)=   "c(21)"
      Tab(9).Control(68)=   "c(20)"
      Tab(9).Control(69)=   "c(22)"
      Tab(9).Control(70)=   "c(23)"
      Tab(9).ControlCount=   71
      TabCaption(10)  =   "P. 2"
      TabPicture(10)  =   "Data_entry.frx":0422
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Line31"
      Tab(10).Control(1)=   "Line30"
      Tab(10).Control(2)=   "Line29"
      Tab(10).Control(3)=   "Line28"
      Tab(10).Control(4)=   "Line27"
      Tab(10).Control(5)=   "Line26"
      Tab(10).Control(6)=   "Line25"
      Tab(10).Control(7)=   "Label102"
      Tab(10).Control(8)=   "Label101"
      Tab(10).Control(9)=   "Label100"
      Tab(10).Control(10)=   "Label99"
      Tab(10).Control(11)=   "Label98"
      Tab(10).Control(12)=   "Label97"
      Tab(10).Control(13)=   "Line23"
      Tab(10).Control(14)=   "Label96"
      Tab(10).Control(15)=   "Label95"
      Tab(10).Control(16)=   "Label94"
      Tab(10).Control(17)=   "Label93"
      Tab(10).Control(18)=   "Label92"
      Tab(10).Control(19)=   "Label91"
      Tab(10).Control(20)=   "Label90"
      Tab(10).Control(21)=   "Label89"
      Tab(10).Control(22)=   "Label88"
      Tab(10).Control(23)=   "Label87"
      Tab(10).Control(24)=   "Label86"
      Tab(10).Control(25)=   "Label85"
      Tab(10).Control(26)=   "Label84(0)"
      Tab(10).Control(27)=   "Label83(0)"
      Tab(10).Control(28)=   "Label82(0)"
      Tab(10).Control(29)=   "Label81(0)"
      Tab(10).Control(30)=   "Label80(0)"
      Tab(10).Control(31)=   "Label79(0)"
      Tab(10).Control(32)=   "Label78"
      Tab(10).Control(33)=   "Label77"
      Tab(10).Control(34)=   "Label75"
      Tab(10).Control(35)=   "Label74"
      Tab(10).Control(36)=   "Label73"
      Tab(10).Control(37)=   "Label72"
      Tab(10).Control(38)=   "Label71"
      Tab(10).Control(39)=   "Label70"
      Tab(10).Control(40)=   "Label69"
      Tab(10).Control(41)=   "Label68"
      Tab(10).Control(42)=   "Label67"
      Tab(10).Control(43)=   "Label66"
      Tab(10).Control(44)=   "Label65"
      Tab(10).Control(45)=   "Label64"
      Tab(10).Control(46)=   "Label63"
      Tab(10).Control(47)=   "Label62"
      Tab(10).Control(48)=   "Label61"
      Tab(10).Control(49)=   "Label60"
      Tab(10).Control(50)=   "Label59"
      Tab(10).Control(51)=   "Label58"
      Tab(10).Control(52)=   "Label57"
      Tab(10).Control(53)=   "Label56"
      Tab(10).Control(54)=   "Label55"
      Tab(10).Control(55)=   "Line16"
      Tab(10).Control(56)=   "Line22"
      Tab(10).Control(57)=   "Line24"
      Tab(10).Control(58)=   "Line20"
      Tab(10).Control(59)=   "Line19"
      Tab(10).Control(60)=   "Line17"
      Tab(10).Control(61)=   "Line18"
      Tab(10).Control(62)=   "l(164)"
      Tab(10).Control(63)=   "dt(2)"
      Tab(10).Control(64)=   "t(44)"
      Tab(10).Control(65)=   "t(45)"
      Tab(10).Control(66)=   "t(46)"
      Tab(10).Control(67)=   "t(47)"
      Tab(10).Control(68)=   "t(48)"
      Tab(10).Control(69)=   "t(49)"
      Tab(10).Control(70)=   "t(51)"
      Tab(10).Control(71)=   "t(52)"
      Tab(10).Control(72)=   "t(53)"
      Tab(10).Control(73)=   "t(54)"
      Tab(10).Control(74)=   "t(55)"
      Tab(10).Control(75)=   "t(58)"
      Tab(10).Control(76)=   "t(57)"
      Tab(10).Control(77)=   "t(56)"
      Tab(10).Control(78)=   "t(61)"
      Tab(10).Control(79)=   "t(60)"
      Tab(10).Control(80)=   "t(59)"
      Tab(10).Control(81)=   "t(63)"
      Tab(10).Control(82)=   "t(62)"
      Tab(10).Control(83)=   "t(64)"
      Tab(10).Control(84)=   "t(67)"
      Tab(10).Control(85)=   "t(66)"
      Tab(10).Control(86)=   "t(65)"
      Tab(10).Control(87)=   "c(105)"
      Tab(10).Control(88)=   "c(106)"
      Tab(10).Control(89)=   "c(107)"
      Tab(10).Control(90)=   "t(50)"
      Tab(10).Control(91)=   "t(70)"
      Tab(10).Control(92)=   "t(69)"
      Tab(10).Control(93)=   "t(68)"
      Tab(10).Control(94)=   "t(75)"
      Tab(10).Control(95)=   "t(74)"
      Tab(10).Control(96)=   "t(73)"
      Tab(10).Control(97)=   "t(72)"
      Tab(10).Control(98)=   "t(76)"
      Tab(10).Control(99)=   "t(77)"
      Tab(10).Control(100)=   "t(78)"
      Tab(10).Control(101)=   "t(79)"
      Tab(10).Control(102)=   "t(80)"
      Tab(10).Control(103)=   "t(81)"
      Tab(10).Control(104)=   "t(82)"
      Tab(10).Control(105)=   "t(83)"
      Tab(10).Control(106)=   "t(84)"
      Tab(10).Control(107)=   "t(85)"
      Tab(10).Control(108)=   "t(86)"
      Tab(10).Control(109)=   "t(87)"
      Tab(10).Control(110)=   "t(71)"
      Tab(10).ControlCount=   111
      TabCaption(11)  =   "P. 1"
      TabPicture(11)  =   "Data_entry.frx":043E
      Tab(11).ControlEnabled=   -1  'True
      Tab(11).Control(0)=   "Label41"
      Tab(11).Control(0).Enabled=   0   'False
      Tab(11).Control(1)=   "Label40"
      Tab(11).Control(1).Enabled=   0   'False
      Tab(11).Control(2)=   "Label39"
      Tab(11).Control(2).Enabled=   0   'False
      Tab(11).Control(3)=   "Label38"
      Tab(11).Control(3).Enabled=   0   'False
      Tab(11).Control(4)=   "Label37"
      Tab(11).Control(4).Enabled=   0   'False
      Tab(11).Control(5)=   "Label36"
      Tab(11).Control(5).Enabled=   0   'False
      Tab(11).Control(6)=   "l(3)"
      Tab(11).Control(6).Enabled=   0   'False
      Tab(11).Control(7)=   "l(2)"
      Tab(11).Control(7).Enabled=   0   'False
      Tab(11).Control(8)=   "l(1)"
      Tab(11).Control(8).Enabled=   0   'False
      Tab(11).Control(9)=   "l(0)"
      Tab(11).Control(9).Enabled=   0   'False
      Tab(11).Control(10)=   "l(34)"
      Tab(11).Control(10).Enabled=   0   'False
      Tab(11).Control(11)=   "Label53"
      Tab(11).Control(11).Enabled=   0   'False
      Tab(11).Control(12)=   "Label52"
      Tab(11).Control(12).Enabled=   0   'False
      Tab(11).Control(13)=   "l(33)"
      Tab(11).Control(13).Enabled=   0   'False
      Tab(11).Control(14)=   "l(32)"
      Tab(11).Control(14).Enabled=   0   'False
      Tab(11).Control(15)=   "l(31)"
      Tab(11).Control(15).Enabled=   0   'False
      Tab(11).Control(16)=   "l(30)"
      Tab(11).Control(16).Enabled=   0   'False
      Tab(11).Control(17)=   "Label47"
      Tab(11).Control(17).Enabled=   0   'False
      Tab(11).Control(18)=   "Label46"
      Tab(11).Control(18).Enabled=   0   'False
      Tab(11).Control(19)=   "l(29)"
      Tab(11).Control(19).Enabled=   0   'False
      Tab(11).Control(20)=   "l(28)"
      Tab(11).Control(20).Enabled=   0   'False
      Tab(11).Control(21)=   "l(27)"
      Tab(11).Control(21).Enabled=   0   'False
      Tab(11).Control(22)=   "Line13"
      Tab(11).Control(22).Enabled=   0   'False
      Tab(11).Control(23)=   "l(16)"
      Tab(11).Control(23).Enabled=   0   'False
      Tab(11).Control(24)=   "l(15)"
      Tab(11).Control(24).Enabled=   0   'False
      Tab(11).Control(25)=   "l(14)"
      Tab(11).Control(25).Enabled=   0   'False
      Tab(11).Control(26)=   "l(13)"
      Tab(11).Control(26).Enabled=   0   'False
      Tab(11).Control(27)=   "l(17)"
      Tab(11).Control(27).Enabled=   0   'False
      Tab(11).Control(28)=   "Line11"
      Tab(11).Control(28).Enabled=   0   'False
      Tab(11).Control(29)=   "l(21)"
      Tab(11).Control(29).Enabled=   0   'False
      Tab(11).Control(30)=   "l(20)"
      Tab(11).Control(30).Enabled=   0   'False
      Tab(11).Control(31)=   "l(19)"
      Tab(11).Control(31).Enabled=   0   'False
      Tab(11).Control(32)=   "l(18)"
      Tab(11).Control(32).Enabled=   0   'False
      Tab(11).Control(33)=   "l(22)"
      Tab(11).Control(33).Enabled=   0   'False
      Tab(11).Control(34)=   "l(5)"
      Tab(11).Control(34).Enabled=   0   'False
      Tab(11).Control(35)=   "Line9"
      Tab(11).Control(35).Enabled=   0   'False
      Tab(11).Control(36)=   "Line10"
      Tab(11).Control(36).Enabled=   0   'False
      Tab(11).Control(37)=   "Label35"
      Tab(11).Control(37).Enabled=   0   'False
      Tab(11).Control(38)=   "Label34"
      Tab(11).Control(38).Enabled=   0   'False
      Tab(11).Control(39)=   "l(23)"
      Tab(11).Control(39).Enabled=   0   'False
      Tab(11).Control(40)=   "l(24)"
      Tab(11).Control(40).Enabled=   0   'False
      Tab(11).Control(41)=   "l(26)"
      Tab(11).Control(41).Enabled=   0   'False
      Tab(11).Control(42)=   "l(25)"
      Tab(11).Control(42).Enabled=   0   'False
      Tab(11).Control(43)=   "Label28"
      Tab(11).Control(43).Enabled=   0   'False
      Tab(11).Control(44)=   "Label27"
      Tab(11).Control(44).Enabled=   0   'False
      Tab(11).Control(45)=   "Label26"
      Tab(11).Control(45).Enabled=   0   'False
      Tab(11).Control(46)=   "Label25"
      Tab(11).Control(46).Enabled=   0   'False
      Tab(11).Control(47)=   "Line8"
      Tab(11).Control(47).Enabled=   0   'False
      Tab(11).Control(48)=   "l(12)"
      Tab(11).Control(48).Enabled=   0   'False
      Tab(11).Control(49)=   "Line6"
      Tab(11).Control(49).Enabled=   0   'False
      Tab(11).Control(50)=   "Label20"
      Tab(11).Control(50).Enabled=   0   'False
      Tab(11).Control(51)=   "l(11)"
      Tab(11).Control(51).Enabled=   0   'False
      Tab(11).Control(52)=   "Line4"
      Tab(11).Control(52).Enabled=   0   'False
      Tab(11).Control(53)=   "l(10)"
      Tab(11).Control(53).Enabled=   0   'False
      Tab(11).Control(54)=   "l(9)"
      Tab(11).Control(54).Enabled=   0   'False
      Tab(11).Control(55)=   "l(8)"
      Tab(11).Control(55).Enabled=   0   'False
      Tab(11).Control(56)=   "l(7)"
      Tab(11).Control(56).Enabled=   0   'False
      Tab(11).Control(57)=   "l(6)"
      Tab(11).Control(57).Enabled=   0   'False
      Tab(11).Control(58)=   "l(4)"
      Tab(11).Control(58).Enabled=   0   'False
      Tab(11).Control(59)=   "Line2"
      Tab(11).Control(59).Enabled=   0   'False
      Tab(11).Control(60)=   "Label12"
      Tab(11).Control(60).Enabled=   0   'False
      Tab(11).Control(61)=   "Label11"
      Tab(11).Control(61).Enabled=   0   'False
      Tab(11).Control(62)=   "l(35)"
      Tab(11).Control(62).Enabled=   0   'False
      Tab(11).Control(63)=   "Line1"
      Tab(11).Control(63).Enabled=   0   'False
      Tab(11).Control(64)=   "Line3"
      Tab(11).Control(64).Enabled=   0   'False
      Tab(11).Control(65)=   "Line5"
      Tab(11).Control(65).Enabled=   0   'False
      Tab(11).Control(66)=   "Line12"
      Tab(11).Control(66).Enabled=   0   'False
      Tab(11).Control(67)=   "Line7"
      Tab(11).Control(67).Enabled=   0   'False
      Tab(11).Control(68)=   "Line14"
      Tab(11).Control(68).Enabled=   0   'False
      Tab(11).Control(69)=   "Line15"
      Tab(11).Control(69).Enabled=   0   'False
      Tab(11).Control(70)=   "dt(1)"
      Tab(11).Control(70).Enabled=   0   'False
      Tab(11).Control(71)=   "Text15"
      Tab(11).Control(71).Enabled=   0   'False
      Tab(11).Control(72)=   "Text14"
      Tab(11).Control(72).Enabled=   0   'False
      Tab(11).Control(73)=   "Text13"
      Tab(11).Control(73).Enabled=   0   'False
      Tab(11).Control(74)=   "t(43)"
      Tab(11).Control(74).Enabled=   0   'False
      Tab(11).Control(75)=   "t(40)"
      Tab(11).Control(75).Enabled=   0   'False
      Tab(11).Control(76)=   "t(41)"
      Tab(11).Control(76).Enabled=   0   'False
      Tab(11).Control(77)=   "t(42)"
      Tab(11).Control(77).Enabled=   0   'False
      Tab(11).Control(78)=   "t(39)"
      Tab(11).Control(78).Enabled=   0   'False
      Tab(11).Control(79)=   "t(38)"
      Tab(11).Control(79).Enabled=   0   'False
      Tab(11).Control(80)=   "t(37)"
      Tab(11).Control(80).Enabled=   0   'False
      Tab(11).Control(81)=   "t(34)"
      Tab(11).Control(81).Enabled=   0   'False
      Tab(11).Control(82)=   "t(35)"
      Tab(11).Control(82).Enabled=   0   'False
      Tab(11).Control(83)=   "t(36)"
      Tab(11).Control(83).Enabled=   0   'False
      Tab(11).Control(84)=   "t(33)"
      Tab(11).Control(84).Enabled=   0   'False
      Tab(11).Control(85)=   "t(32)"
      Tab(11).Control(85).Enabled=   0   'False
      Tab(11).Control(86)=   "t(28)"
      Tab(11).Control(86).Enabled=   0   'False
      Tab(11).Control(87)=   "t(29)"
      Tab(11).Control(87).Enabled=   0   'False
      Tab(11).Control(88)=   "t(30)"
      Tab(11).Control(88).Enabled=   0   'False
      Tab(11).Control(89)=   "t(31)"
      Tab(11).Control(89).Enabled=   0   'False
      Tab(11).Control(90)=   "t(24)"
      Tab(11).Control(90).Enabled=   0   'False
      Tab(11).Control(91)=   "t(25)"
      Tab(11).Control(91).Enabled=   0   'False
      Tab(11).Control(92)=   "t(26)"
      Tab(11).Control(92).Enabled=   0   'False
      Tab(11).Control(93)=   "t(6)"
      Tab(11).Control(93).Enabled=   0   'False
      Tab(11).Control(94)=   "c(104)"
      Tab(11).Control(94).Enabled=   0   'False
      Tab(11).Control(95)=   "c(103)"
      Tab(11).Control(95).Enabled=   0   'False
      Tab(11).Control(96)=   "c(102)"
      Tab(11).Control(96).Enabled=   0   'False
      Tab(11).Control(97)=   "t(21)"
      Tab(11).Control(97).Enabled=   0   'False
      Tab(11).Control(98)=   "t(22)"
      Tab(11).Control(98).Enabled=   0   'False
      Tab(11).Control(99)=   "t(23)"
      Tab(11).Control(99).Enabled=   0   'False
      Tab(11).Control(100)=   "t(20)"
      Tab(11).Control(100).Enabled=   0   'False
      Tab(11).Control(101)=   "t(18)"
      Tab(11).Control(101).Enabled=   0   'False
      Tab(11).Control(102)=   "t(19)"
      Tab(11).Control(102).Enabled=   0   'False
      Tab(11).Control(103)=   "t(17)"
      Tab(11).Control(103).Enabled=   0   'False
      Tab(11).Control(104)=   "t(16)"
      Tab(11).Control(104).Enabled=   0   'False
      Tab(11).Control(105)=   "t(15)"
      Tab(11).Control(105).Enabled=   0   'False
      Tab(11).Control(106)=   "t(14)"
      Tab(11).Control(106).Enabled=   0   'False
      Tab(11).Control(107)=   "t(13)"
      Tab(11).Control(107).Enabled=   0   'False
      Tab(11).Control(108)=   "t(12)"
      Tab(11).Control(108).Enabled=   0   'False
      Tab(11).Control(109)=   "t(11)"
      Tab(11).Control(109).Enabled=   0   'False
      Tab(11).Control(110)=   "t(10)"
      Tab(11).Control(110).Enabled=   0   'False
      Tab(11).Control(111)=   "t(9)"
      Tab(11).Control(111).Enabled=   0   'False
      Tab(11).Control(112)=   "t(8)"
      Tab(11).Control(112).Enabled=   0   'False
      Tab(11).Control(113)=   "t(7)"
      Tab(11).Control(113).Enabled=   0   'False
      Tab(11).Control(114)=   "t(5)"
      Tab(11).Control(114).Enabled=   0   'False
      Tab(11).Control(115)=   "t(4)"
      Tab(11).Control(115).Enabled=   0   'False
      Tab(11).Control(116)=   "t(3)"
      Tab(11).Control(116).Enabled=   0   'False
      Tab(11).Control(117)=   "t(2)"
      Tab(11).Control(117).Enabled=   0   'False
      Tab(11).Control(118)=   "t(1)"
      Tab(11).Control(118).Enabled=   0   'False
      Tab(11).Control(119)=   "t(0)"
      Tab(11).Control(119).Enabled=   0   'False
      Tab(11).Control(120)=   "c(101)"
      Tab(11).Control(120).Enabled=   0   'False
      Tab(11).Control(121)=   "t(27)"
      Tab(11).Control(121).Enabled=   0   'False
      Tab(11).ControlCount=   122
      Begin VB.CheckBox c 
         Caption         =   "Each witness"
         DataField       =   "DESM"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   48
         Left            =   -68070
         TabIndex        =   616
         Top             =   7035
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant"
         DataField       =   "DDSM"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   49
         Left            =   -69225
         TabIndex        =   615
         Top             =   7035
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Investigated officer"
         DataField       =   "DISM"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   47
         Left            =   -68280
         TabIndex        =   614
         Top             =   6615
         Width           =   1905
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff"
         DataField       =   "DPSM"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   46
         Left            =   -69225
         TabIndex        =   613
         Top             =   6615
         Width           =   1275
      End
      Begin VB.CheckBox c 
         Caption         =   "Other"
         DataField       =   "RPSO"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   45
         Left            =   -67545
         TabIndex        =   611
         Top             =   5145
         Width           =   960
      End
      Begin VB.CheckBox c 
         Caption         =   "By ambulance"
         DataField       =   "RPSAM"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   44
         Left            =   -69015
         TabIndex        =   610
         Top             =   5145
         Width           =   1380
      End
      Begin VB.CheckBox c 
         Caption         =   "By aid car"
         DataField       =   "RPSA"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   43
         Left            =   -70170
         TabIndex        =   609
         Top             =   5145
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "By self"
         DataField       =   "RPSS"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   42
         Left            =   -71220
         TabIndex        =   608
         Top             =   5145
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Other"
         DataField       =   "TSO"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   41
         Left            =   -67860
         TabIndex        =   606
         Top             =   4200
         Width           =   1275
      End
      Begin VB.CheckBox c 
         Caption         =   "Ambulance"
         DataField       =   "TSAMB"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   40
         Left            =   -69645
         TabIndex        =   605
         Top             =   4200
         Width           =   1170
      End
      Begin VB.CheckBox c 
         Caption         =   "Aid Car"
         DataField       =   "TSA"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   39
         Left            =   -71115
         TabIndex        =   604
         Top             =   4200
         Width           =   960
      End
      Begin VB.CheckBox c 
         Caption         =   "Directional Signal Lights"
         DataField       =   "Directional_S"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   38
         Left            =   -68385
         TabIndex        =   602
         Top             =   1155
         Width           =   2010
      End
      Begin VB.CheckBox c 
         Caption         =   "Tail Lights"
         DataField       =   "Tail_Lights"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   37
         Left            =   -69645
         TabIndex        =   601
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Brake Lights"
         DataField       =   "Brakelights"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   36
         Left            =   -71220
         TabIndex        =   600
         Top             =   1155
         Width           =   1275
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant"
         DataField       =   "APO"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   35
         Left            =   -73215
         TabIndex        =   598
         Top             =   1155
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff"
         DataField       =   "APP"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   34
         Left            =   -74370
         TabIndex        =   597
         Top             =   1155
         Width           =   960
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant"
         DataField       =   "ED"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   33
         Left            =   -69540
         TabIndex        =   595
         Top             =   7140
         Width           =   1380
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff"
         DataField       =   "EP"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   32
         Left            =   -69540
         TabIndex        =   594
         Top             =   6825
         Width           =   1380
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant"
         DataField       =   "TTD"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   31
         Left            =   -68280
         TabIndex        =   592
         Top             =   2100
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff"
         DataField       =   "TTP"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   30
         Left            =   -69330
         TabIndex        =   591
         Top             =   2100
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant"
         DataField       =   "FD"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   29
         Left            =   -68070
         TabIndex        =   589
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff"
         DataField       =   "FP"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   28
         Left            =   -69120
         TabIndex        =   588
         Top             =   1155
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Other"
         DataField       =   "DO"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   27
         Left            =   -70065
         TabIndex        =   586
         Top             =   1155
         Width           =   750
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant"
         DataField       =   "DD"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   26
         Left            =   -71220
         TabIndex        =   585
         Top             =   1155
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff"
         DataField       =   "DP"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   25
         Left            =   -72165
         TabIndex        =   584
         Top             =   1155
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Other"
         DataField       =   "Photoes_O2"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   23
         Left            =   -67230
         TabIndex        =   582
         Top             =   3675
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Scene"
         DataField       =   "Photoes_S"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   22
         Left            =   -68070
         TabIndex        =   581
         Top             =   3675
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Defendant's Veh."
         DataField       =   "Photoes_D"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   20
         Left            =   -69645
         TabIndex        =   580
         Top             =   3675
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Plaintiff's Veh."
         DataField       =   "Photoes_P"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   21
         Left            =   -71010
         TabIndex        =   579
         Top             =   3675
         Width           =   1380
      End
      Begin VB.CheckBox c 
         Caption         =   "Other"
         DataField       =   "Photoes_O"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   24
         Left            =   -67335
         TabIndex        =   578
         Top             =   2835
         Width           =   750
      End
      Begin VB.CheckBox c 
         Caption         =   "Of damage"
         DataField       =   "Photoes_Od"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   19
         Left            =   -68490
         TabIndex        =   577
         Top             =   2835
         Width           =   1170
      End
      Begin VB.CheckBox c 
         Caption         =   "Of accident or..."
         DataField       =   "Photoes_Oa"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   18
         Left            =   -69960
         TabIndex        =   576
         Top             =   2835
         Width           =   1485
      End
      Begin VB.CheckBox c 
         Caption         =   "By whom"
         DataField       =   "Photoes_B"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   17
         Left            =   -71010
         TabIndex        =   575
         Top             =   2835
         Width           =   960
      End
      Begin VB.CheckBox c 
         Caption         =   "Oral"
         DataField       =   "D_Oral"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   16
         Left            =   -72375
         TabIndex        =   572
         Top             =   6825
         Width           =   855
      End
      Begin VB.CheckBox check1 
         Caption         =   "Written"
         DataField       =   "D_Written"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   0
         Left            =   -72375
         TabIndex        =   571
         Top             =   7140
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Made by the Plaintiff"
         DataField       =   "P_Written"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   14
         Left            =   -74475
         TabIndex        =   570
         Top             =   6825
         Width           =   2010
      End
      Begin VB.CheckBox c 
         Caption         =   "Made by the Defendant"
         DataField       =   "P_Written"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   15
         Left            =   -74475
         TabIndex        =   569
         Top             =   7140
         Width           =   2010
      End
      Begin VB.Data Data10 
         Caption         =   "Data10"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Left            =   -71535
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   630
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Data Data9 
         Caption         =   "Data9"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -68175
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   315
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Data Data8 
         Caption         =   "Data8"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   330
         Left            =   -70275
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   210
         Visible         =   0   'False
         Width           =   1905
      End
      Begin VB.Data Data7 
         Caption         =   "Data7"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   -72690
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   210
         Visible         =   0   'False
         Width           =   2220
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Prior_Employment_History"
         DataSource      =   "Data4"
         Height          =   540
         Index           =   71
         Left            =   -70065
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   78
         Top             =   4620
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Prior_Employment_History"
         DataSource      =   "Data3"
         Height          =   540
         Index           =   27
         Left            =   4935
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   31
         Top             =   4620
         Width           =   1485
      End
      Begin VB.Frame f 
         Caption         =   "Select Filter..."
         Height          =   1905
         Index           =   21
         Left            =   -69015
         TabIndex        =   561
         Top             =   5670
         Width           =   2640
         Begin VB.CheckBox c 
            Caption         =   "Today:"
            Height          =   225
            Index           =   154
            Left            =   105
            TabIndex        =   256
            Top             =   735
            Width           =   1170
         End
         Begin VB.CheckBox c 
            Caption         =   "Date Start:"
            Height          =   225
            Index           =   153
            Left            =   105
            TabIndex        =   254
            Top             =   420
            Width           =   1170
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   945
            TabIndex        =   258
            Text            =   "(None)"
            Top             =   945
            Width           =   1590
         End
         Begin VB.CommandButton co 
            BackColor       =   &H00C0FFFF&
            Caption         =   "View"
            Height          =   330
            Index           =   18
            Left            =   1365
            Style           =   1  'Graphical
            TabIndex        =   260
            Top             =   1470
            Width           =   1170
         End
         Begin VB.CommandButton co 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Print"
            Height          =   330
            Index           =   17
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   259
            Top             =   1470
            Width           =   1170
         End
         Begin MSComCtl2.DTPicker dt 
            Height          =   330
            Index           =   16
            Left            =   1365
            TabIndex        =   255
            Top             =   315
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin MSComCtl2.DTPicker dt 
            Height          =   330
            Index           =   17
            Left            =   1365
            TabIndex        =   257
            Top             =   630
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin VB.Label l 
            Caption         =   "Filter:"
            Height          =   225
            Index           =   166
            Left            =   105
            TabIndex        =   562
            Top             =   1050
            Width           =   435
         End
      End
      Begin VB.TextBox t 
         DataField       =   "M"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   183
         Left            =   -67860
         TabIndex        =   249
         Text            =   "Text5"
         Top             =   1680
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "I"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   184
         Left            =   -67860
         TabIndex        =   250
         Text            =   "Text5"
         Top             =   1995
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "C"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   185
         Left            =   -67860
         TabIndex        =   251
         Text            =   "Text5"
         Top             =   2310
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "CO"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   186
         Left            =   -67860
         TabIndex        =   252
         Text            =   "Text5"
         Top             =   2625
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "CC"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   187
         Left            =   -67860
         TabIndex        =   253
         Text            =   "Text5"
         Top             =   2940
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "L"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   182
         Left            =   -67860
         TabIndex        =   248
         Text            =   "Text5"
         Top             =   1365
         Width           =   1380
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Edit"
         Height          =   330
         Index           =   15
         Left            =   -69435
         Style           =   1  'Graphical
         TabIndex        =   556
         Top             =   2205
         Width           =   1380
      End
      Begin VB.ComboBox com 
         Height          =   315
         Index           =   1
         Left            =   -69435
         TabIndex        =   555
         Text            =   "Account #"
         Top             =   1470
         Width           =   1380
      End
      Begin VB.ComboBox com 
         Height          =   315
         Index           =   0
         Left            =   -69435
         TabIndex        =   554
         Text            =   "Select Bank"
         Top             =   1050
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "TM"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   177
         Left            =   -72585
         TabIndex        =   243
         Text            =   "Text5"
         Top             =   1680
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "TI"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   178
         Left            =   -72585
         TabIndex        =   244
         Text            =   "Text5"
         Top             =   1995
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "TC"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   179
         Left            =   -72585
         TabIndex        =   245
         Text            =   "Text5"
         Top             =   2310
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "T_Copy"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   180
         Left            =   -72585
         TabIndex        =   246
         Text            =   "Text5"
         Top             =   2625
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "T_CC"
         DataSource      =   "Data1"
         Height          =   285
         Index           =   181
         Left            =   -72585
         TabIndex        =   247
         Text            =   "Text5"
         Top             =   2940
         Width           =   1380
      End
      Begin VB.TextBox t 
         DataField       =   "TL"
         DataSource      =   "Data1"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   176
         Left            =   -72585
         TabIndex        =   242
         Text            =   "Text5"
         Top             =   1365
         Width           =   1380
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add"
         Height          =   330
         Index           =   16
         Left            =   -69435
         Style           =   1  'Graphical
         TabIndex        =   544
         Top             =   1890
         Width           =   1380
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Accept"
         Height          =   330
         Index           =   14
         Left            =   -69435
         Style           =   1  'Graphical
         TabIndex        =   543
         Top             =   2520
         Width           =   1380
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancel"
         Height          =   330
         Index           =   13
         Left            =   -69435
         Style           =   1  'Graphical
         TabIndex        =   542
         Top             =   2835
         Width           =   1380
      End
      Begin VB.Data Data1 
         Caption         =   "View Payments"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -69015
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5145
         Width           =   2640
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "Data_entry.frx":045A
         Height          =   1170
         Left            =   -74475
         OleObjectBlob   =   "Data_entry.frx":046E
         TabIndex        =   541
         Top             =   3885
         Width           =   7995
      End
      Begin VB.CheckBox c 
         Caption         =   "Medical Records"
         DataField       =   "Medical_Records"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   148
         Left            =   -74370
         TabIndex        =   237
         Top             =   1680
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Investigative"
         DataField       =   "Investigative"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   149
         Left            =   -74370
         TabIndex        =   238
         Top             =   1995
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Cost Advanced"
         DataField       =   "Cost_Advanced"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   150
         Left            =   -74370
         TabIndex        =   239
         Top             =   2310
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Copy"
         DataField       =   "Copy"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   151
         Left            =   -74370
         TabIndex        =   240
         Top             =   2625
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Court Costs"
         DataField       =   "Court_Costs"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   152
         Left            =   -74370
         TabIndex        =   241
         Top             =   2940
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Liens"
         DataField       =   "Liens"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   147
         Left            =   -74370
         TabIndex        =   236
         Top             =   1365
         Width           =   1590
      End
      Begin VB.PictureBox pi 
         BackColor       =   &H80000007&
         Height          =   4110
         Left            =   -71745
         ScaleHeight     =   4050
         ScaleWidth      =   5205
         TabIndex        =   539
         Top             =   945
         Visible         =   0   'False
         Width           =   5265
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Go to !"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   7
         Left            =   -68280
         MaskColor       =   &H80000003&
         Style           =   1  'Graphical
         TabIndex        =   275
         Top             =   420
         Width           =   1485
      End
      Begin VB.CheckBox c 
         Caption         =   "Job Already Active!"
         DataField       =   "Job_Activation"
         DataSource      =   "Data3"
         Height          =   225
         Index           =   101
         Left            =   2835
         TabIndex        =   538
         Top             =   315
         Width           =   2220
      End
      Begin VB.CommandButton co 
         Caption         =   "Play"
         Enabled         =   0   'False
         Height          =   435
         Index           =   6
         Left            =   -69330
         TabIndex        =   271
         Top             =   7035
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton co 
         Caption         =   "Close"
         Enabled         =   0   'False
         Height          =   435
         Index           =   4
         Left            =   -67440
         TabIndex        =   273
         Top             =   7035
         Width           =   960
      End
      Begin VB.CommandButton co 
         Caption         =   "Pause"
         Enabled         =   0   'False
         Height          =   435
         Index           =   5
         Left            =   -68385
         TabIndex        =   272
         Top             =   7035
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.CommandButton co 
         Caption         =   "Remove Document"
         Height          =   540
         Index           =   11
         Left            =   -74475
         TabIndex        =   270
         Top             =   6930
         Width           =   1695
      End
      Begin MCI.MMControl mm 
         Height          =   435
         Left            =   -69120
         TabIndex        =   535
         Top             =   7035
         Width           =   1650
         _ExtentX        =   2910
         _ExtentY        =   767
         _Version        =   393216
         PlayEnabled     =   -1  'True
         PauseEnabled    =   -1  'True
         PrevVisible     =   0   'False
         NextVisible     =   0   'False
         BackVisible     =   0   'False
         StepVisible     =   0   'False
         StopVisible     =   0   'False
         RecordVisible   =   0   'False
         EjectVisible    =   0   'False
         DeviceType      =   ""
         FileName        =   ""
      End
      Begin VB.CommandButton co 
         Caption         =   "Add Other..."
         Height          =   540
         Index           =   3
         Left            =   -69435
         TabIndex        =   269
         Top             =   5355
         Width           =   1695
      End
      Begin VB.CommandButton co 
         Caption         =   "Add Movies"
         Height          =   540
         Index           =   2
         Left            =   -71115
         TabIndex        =   268
         Top             =   5355
         Width           =   1695
      End
      Begin VB.CommandButton co 
         Caption         =   "Add Soundes"
         Height          =   540
         Index           =   1
         Left            =   -72795
         TabIndex        =   267
         Top             =   5355
         Width           =   1695
      End
      Begin VB.CommandButton co 
         Caption         =   "Add Photoes"
         Height          =   540
         Index           =   0
         Left            =   -74475
         TabIndex        =   266
         Top             =   5355
         Width           =   1695
      End
      Begin MSComctlLib.TreeView tv 
         Height          =   4110
         Index           =   0
         Left            =   -74475
         TabIndex        =   265
         Top             =   945
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   7250
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         SingleSel       =   -1  'True
         Appearance      =   1
      End
      Begin MSComDlg.CommonDialog cd 
         Left            =   -67440
         Top             =   5985
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.TreeView tv 
         Height          =   6525
         Index           =   1
         Left            =   -74265
         TabIndex        =   274
         Top             =   945
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   11509
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFC0&
         Caption         =   "A&ccept"
         Height          =   435
         Index           =   10
         Left            =   -71955
         Style           =   1  'Graphical
         TabIndex        =   264
         Top             =   6405
         Width           =   1170
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Append"
         Height          =   435
         Index           =   9
         Left            =   -73110
         Style           =   1  'Graphical
         TabIndex        =   263
         Top             =   6405
         Width           =   1170
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&New"
         Height          =   435
         Index           =   8
         Left            =   -74265
         Style           =   1  'Graphical
         TabIndex        =   262
         Top             =   6405
         Width           =   1170
      End
      Begin VB.TextBox Text4 
         ForeColor       =   &H00FF0000&
         Height          =   5160
         Left            =   -74265
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   261
         Top             =   945
         Width           =   7470
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFC0C0&
         DataField       =   "TMS"
         DataSource      =   "Data10"
         Height          =   330
         Index           =   175
         Left            =   -68175
         TabIndex        =   235
         Text            =   "Text1"
         Top             =   6825
         Width           =   1800
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFC0C0&
         DataField       =   "SO"
         DataSource      =   "Data10"
         Height          =   330
         Index           =   174
         Left            =   -68175
         TabIndex        =   234
         Text            =   "Text1"
         Top             =   5775
         Width           =   1800
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFC0C0&
         DataField       =   "IC"
         DataSource      =   "Data10"
         Height          =   330
         Index           =   173
         Left            =   -68175
         TabIndex        =   233
         Text            =   "Text1"
         Top             =   5145
         Width           =   1800
      End
      Begin VB.Frame f 
         Caption         =   "Plaintiff's prior Medical History:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Index           =   20
         Left            =   -70065
         TabIndex        =   526
         Top             =   1155
         Width           =   3690
         Begin VB.CheckBox c 
            Caption         =   "Past injuries or symptoms to same area of body injured in injury or incident"
            DataField       =   "CPI"
            DataSource      =   "Data10"
            Height          =   435
            Index           =   146
            Left            =   105
            TabIndex        =   231
            Top             =   2835
            Width           =   3480
         End
         Begin VB.CheckBox c 
            Caption         =   "Past Claims"
            DataField       =   "CPC"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   145
            Left            =   105
            TabIndex        =   229
            Top             =   2415
            Width           =   1800
         End
         Begin VB.CheckBox c 
            Caption         =   "Past Auto Accidents"
            DataField       =   "CPAA"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   144
            Left            =   105
            TabIndex        =   227
            Top             =   1995
            Width           =   1800
         End
         Begin VB.CheckBox c 
            Caption         =   "Past accident of broken bones or other serious injuries"
            DataField       =   "CPA"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   143
            Left            =   105
            TabIndex        =   225
            Top             =   1260
            Width           =   3480
         End
         Begin VB.CheckBox c 
            Caption         =   "Past Serious Injuries"
            DataField       =   "CPSI"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   142
            Left            =   105
            TabIndex        =   223
            Top             =   840
            Width           =   1800
         End
         Begin VB.CheckBox c 
            Caption         =   "Past Hospitalizations"
            DataField       =   "CPH"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   141
            Left            =   105
            TabIndex        =   221
            Top             =   420
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dt 
            DataField       =   "PH"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   10
            Left            =   2100
            TabIndex        =   222
            Top             =   420
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin MSComCtl2.DTPicker dt 
            DataField       =   "PSI"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   11
            Left            =   2100
            TabIndex        =   224
            Top             =   840
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin MSComCtl2.DTPicker dt 
            DataField       =   "PA"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   12
            Left            =   2100
            TabIndex        =   226
            Top             =   1575
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin MSComCtl2.DTPicker dt 
            DataField       =   "PAA"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   13
            Left            =   2100
            TabIndex        =   228
            Top             =   1995
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin MSComCtl2.DTPicker dt 
            DataField       =   "PC"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   14
            Left            =   2100
            TabIndex        =   230
            Top             =   2415
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24969217
            CurrentDate     =   36825
         End
         Begin MSComCtl2.DTPicker dt 
            DataField       =   "PI"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   15
            Left            =   2100
            TabIndex        =   232
            Top             =   3255
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   24969217
            CurrentDate     =   36825
         End
      End
      Begin VB.Frame f 
         Caption         =   "Out-of-Pocket Expenses:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Index           =   19
         Left            =   -74265
         TabIndex        =   525
         Top             =   2625
         Width           =   4005
         Begin VB.CheckBox c 
            Caption         =   "Other"
            DataField       =   "CO"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   140
            Left            =   210
            TabIndex        =   219
            Top             =   4095
            Width           =   1065
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "O"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   172
            Left            =   2310
            TabIndex        =   220
            Text            =   "Text1"
            Top             =   4095
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Lost wages"
            DataField       =   "CL"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   139
            Left            =   210
            TabIndex        =   217
            Top             =   3675
            Width           =   1485
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "L"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   171
            Left            =   2310
            TabIndex        =   218
            Text            =   "Text1"
            Top             =   3675
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Auto repair"
            DataField       =   "CAU"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   138
            Left            =   210
            TabIndex        =   215
            Top             =   3255
            Width           =   1065
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "AU"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   170
            Left            =   2310
            TabIndex        =   216
            Text            =   "Text1"
            Top             =   3255
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Domestic Help"
            DataField       =   "CDO"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   137
            Left            =   210
            TabIndex        =   213
            Top             =   2835
            Width           =   1695
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "DO"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   169
            Left            =   2310
            TabIndex        =   214
            Text            =   "Text1"
            Top             =   2835
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Crutches,Appliances"
            DataField       =   "CC"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   136
            Left            =   210
            TabIndex        =   211
            Top             =   2415
            Width           =   1905
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "C"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   168
            Left            =   2310
            TabIndex        =   212
            Text            =   "Text1"
            Top             =   2415
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Drugs"
            DataField       =   "CD"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   135
            Left            =   210
            TabIndex        =   209
            Top             =   1995
            Width           =   1065
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "D"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   167
            Left            =   2310
            TabIndex        =   210
            Text            =   "Text1"
            Top             =   1995
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Ambulance"
            DataField       =   "CA"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   134
            Left            =   210
            TabIndex        =   207
            Top             =   1575
            Width           =   1275
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "A"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   166
            Left            =   2310
            TabIndex        =   208
            Text            =   "Text1"
            Top             =   1575
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Hospitals"
            DataField       =   "CH"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   133
            Left            =   210
            TabIndex        =   205
            Top             =   1155
            Width           =   1065
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "H"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   165
            Left            =   2310
            TabIndex        =   206
            Text            =   "Text1"
            Top             =   1155
            Width           =   1485
         End
         Begin VB.CheckBox c 
            Caption         =   "Physicians"
            DataField       =   "CP"
            DataSource      =   "Data10"
            Height          =   225
            Index           =   132
            Left            =   210
            TabIndex        =   203
            Top             =   735
            Width           =   1065
         End
         Begin VB.TextBox t 
            BackColor       =   &H0080C0FF&
            DataField       =   "P"
            DataSource      =   "Data10"
            Height          =   330
            Index           =   164
            Left            =   2310
            TabIndex        =   204
            Text            =   "Text1"
            Top             =   735
            Width           =   1485
         End
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Reason_For"
         DataSource      =   "Data10"
         Height          =   330
         Index           =   163
         Left            =   -72480
         TabIndex        =   202
         Text            =   "Text1"
         Top             =   2100
         Width           =   2220
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "R_of_Com"
         DataSource      =   "Data10"
         Height          =   330
         Index           =   162
         Left            =   -71745
         TabIndex        =   201
         Text            =   "Text1"
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Dates_Missed"
         DataSource      =   "Data10"
         Height          =   330
         Index           =   161
         Left            =   -71745
         TabIndex        =   200
         Text            =   "Text1"
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H0080C0FF&
         DataField       =   "TN"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   156
         Left            =   -67860
         TabIndex        =   193
         Text            =   "Text1"
         Top             =   3675
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H0080C0FF&
         DataField       =   "TS"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   157
         Left            =   -67860
         TabIndex        =   194
         Text            =   "Text1"
         Top             =   4095
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H0080C0FF&
         DataField       =   "TC"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   158
         Left            =   -67860
         TabIndex        =   195
         Text            =   "Text1"
         Top             =   4515
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H0080C0FF&
         DataField       =   "TCO"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   159
         Left            =   -67860
         TabIndex        =   196
         Text            =   "Text1"
         Top             =   4935
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H0080C0FF&
         DataField       =   "TST"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   160
         Left            =   -67860
         TabIndex        =   197
         Text            =   "Text1"
         Top             =   5355
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Last_Name"
         DataSource      =   "Data2"
         Height          =   330
         Index           =   0
         Left            =   1785
         TabIndex        =   0
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "First_Name"
         DataSource      =   "Data2"
         Height          =   330
         Index           =   1
         Left            =   1785
         TabIndex        =   1
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "MI"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   2
         Left            =   2940
         TabIndex        =   2
         Top             =   1680
         Width           =   330
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "SS1"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   3
         Left            =   1785
         TabIndex        =   3
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "SS2"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   4
         Left            =   2310
         TabIndex        =   4
         Top             =   2100
         Width           =   330
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "SS3"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   5
         Left            =   2730
         TabIndex        =   5
         Top             =   2100
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_StreetName"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   7
         Left            =   1785
         TabIndex        =   7
         Top             =   3360
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_City"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   8
         Left            =   1785
         TabIndex        =   8
         Top             =   3780
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_County"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   9
         Left            =   1785
         TabIndex        =   9
         Top             =   4200
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_State"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   10
         Left            =   1785
         TabIndex        =   10
         Top             =   4620
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_Zip_Code"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   11
         Left            =   2730
         TabIndex        =   11
         Top             =   5040
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "I_Phone_Home3"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   12
         Left            =   2730
         TabIndex        =   15
         Top             =   5880
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "I_Phone_Home2"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   13
         Left            =   2205
         TabIndex        =   14
         Top             =   5880
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "I_Phone_Home1"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   14
         Left            =   1680
         TabIndex        =   13
         Top             =   5880
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "I_Phone_Business3"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   15
         Left            =   2730
         TabIndex        =   19
         Top             =   6300
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "I_Phone_Business2"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   16
         Left            =   2205
         TabIndex        =   18
         Top             =   6300
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "I_Phone_Business1"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   17
         Left            =   1680
         TabIndex        =   17
         Top             =   6300
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_First_Name"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   19
         Left            =   4935
         TabIndex        =   23
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Last_name"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   18
         Left            =   4935
         TabIndex        =   22
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Occupation"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   20
         Left            =   4935
         TabIndex        =   24
         Top             =   2100
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Phone3"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   23
         Left            =   5880
         TabIndex        =   27
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Phone2"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   22
         Left            =   5355
         TabIndex        =   26
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Phone1"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   21
         Left            =   4830
         TabIndex        =   25
         Top             =   2520
         Width           =   435
      End
      Begin VB.CheckBox c 
         Caption         =   "Spouse:"
         DataField       =   "Spouse"
         DataSource      =   "Data3"
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
         Index           =   102
         Left            =   3570
         TabIndex        =   21
         Top             =   840
         Width           =   1275
      End
      Begin VB.CheckBox c 
         Caption         =   "Home:"
         DataField       =   "C_I_Phone_Home"
         DataSource      =   "Data3"
         Height          =   225
         Index           =   103
         Left            =   630
         TabIndex        =   12
         Top             =   5985
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Business:"
         DataField       =   "C_I_Phone_Business"
         DataSource      =   "Data3"
         Height          =   225
         Index           =   104
         Left            =   630
         TabIndex        =   16
         Top             =   6405
         Width           =   1065
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_Number"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   6
         Left            =   2205
         TabIndex        =   6
         Top             =   2940
         Width           =   1065
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "How_Lomg_Employed"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   26
         Left            =   5565
         TabIndex        =   30
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "Nature_Of_Work"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   25
         Left            =   5250
         TabIndex        =   29
         Top             =   3780
         Width           =   1170
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "Occupation_By_Whom"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   24
         Left            =   5250
         TabIndex        =   28
         Top             =   3360
         Width           =   1170
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Limits"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   31
         Left            =   4935
         TabIndex        =   35
         Top             =   6930
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Medical"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   30
         Left            =   4935
         TabIndex        =   34
         Top             =   6510
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "PIP"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   29
         Left            =   4935
         TabIndex        =   33
         Top             =   6090
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Liability"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   28
         Left            =   4935
         TabIndex        =   32
         Top             =   5670
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   32
         Left            =   7140
         TabIndex        =   36
         Top             =   1155
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Name"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   33
         Left            =   7140
         TabIndex        =   37
         Top             =   1785
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Phone1"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   36
         Left            =   7035
         TabIndex        =   38
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Phone2"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   35
         Left            =   7560
         TabIndex        =   39
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Phone3"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   34
         Left            =   8085
         TabIndex        =   40
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Claim"
         DataSource      =   "Data3"
         Height          =   345
         Index           =   37
         Left            =   7140
         TabIndex        =   41
         Top             =   3150
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   38
         Left            =   7140
         TabIndex        =   42
         Top             =   3885
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Name"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   39
         Left            =   7140
         TabIndex        =   43
         Top             =   4515
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Phone1"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   42
         Left            =   7035
         TabIndex        =   44
         Top             =   5250
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Phone2"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   41
         Left            =   7560
         TabIndex        =   45
         Top             =   5250
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Phone3"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   40
         Left            =   8085
         TabIndex        =   46
         Top             =   5250
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Claim"
         DataSource      =   "Data3"
         Height          =   345
         Index           =   43
         Left            =   7140
         TabIndex        =   47
         Top             =   5880
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Claim"
         DataSource      =   "Data4"
         Height          =   345
         Index           =   87
         Left            =   -67860
         TabIndex        =   94
         Top             =   5880
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Phone3"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   86
         Left            =   -66915
         TabIndex        =   93
         Top             =   5250
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Phone2"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   85
         Left            =   -67440
         TabIndex        =   92
         Top             =   5250
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Phone1"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   84
         Left            =   -67965
         TabIndex        =   91
         Top             =   5250
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD_Name"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   83
         Left            =   -67860
         TabIndex        =   90
         Top             =   4515
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "PD"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   82
         Left            =   -67860
         TabIndex        =   89
         Top             =   3885
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Claim"
         DataSource      =   "Data4"
         Height          =   345
         Index           =   81
         Left            =   -67860
         TabIndex        =   88
         Top             =   3150
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Phone3"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   80
         Left            =   -66915
         TabIndex        =   87
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Phone2"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   79
         Left            =   -67440
         TabIndex        =   86
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Phone1"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   78
         Left            =   -67965
         TabIndex        =   85
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI_Name"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   77
         Left            =   -67860
         TabIndex        =   84
         Top             =   1785
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "BI"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   76
         Left            =   -67860
         TabIndex        =   83
         Top             =   1155
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Liability"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   72
         Left            =   -70065
         TabIndex        =   79
         Top             =   5670
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "PIP"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   73
         Left            =   -70065
         TabIndex        =   80
         Top             =   6090
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Medical"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   74
         Left            =   -70065
         TabIndex        =   81
         Top             =   6510
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Limits"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   75
         Left            =   -70065
         TabIndex        =   82
         Top             =   6930
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "Occupation_By_Whom"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   68
         Left            =   -69750
         TabIndex        =   75
         Top             =   3360
         Width           =   1170
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "Nature_Of_Work"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   69
         Left            =   -69750
         TabIndex        =   76
         Top             =   3780
         Width           =   1170
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "How_Long_Employed"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   70
         Left            =   -69435
         TabIndex        =   77
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_Number"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   50
         Left            =   -72795
         TabIndex        =   54
         Top             =   2940
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Business:"
         DataField       =   "C_D_Phone_Business"
         DataSource      =   "Data4"
         Height          =   225
         Index           =   107
         Left            =   -74370
         TabIndex        =   64
         Top             =   6405
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Home:"
         DataField       =   "C_D_Phone_Home"
         DataSource      =   "Data4"
         Height          =   225
         Index           =   106
         Left            =   -74370
         TabIndex        =   60
         Top             =   5985
         Width           =   855
      End
      Begin VB.CheckBox c 
         Caption         =   "Spouse:"
         DataField       =   "Spouse"
         DataSource      =   "Data4"
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
         Index           =   105
         Left            =   -71430
         TabIndex        =   69
         Top             =   840
         Width           =   1275
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Phone1"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   65
         Left            =   -70170
         TabIndex        =   157
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Phone2"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   66
         Left            =   -69645
         TabIndex        =   73
         Top             =   2520
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Phone3"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   67
         Left            =   -69120
         TabIndex        =   74
         Top             =   2520
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Occupation"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   64
         Left            =   -70065
         TabIndex        =   72
         Top             =   2100
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Lastname"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   62
         Left            =   -70065
         TabIndex        =   70
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Spouse_Firstname"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   63
         Left            =   -70065
         TabIndex        =   71
         Top             =   1680
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "D_Phone_Business1"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   59
         Left            =   -73320
         TabIndex        =   65
         Top             =   6300
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "D_Phone_Business2"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   60
         Left            =   -72795
         TabIndex        =   66
         Top             =   6300
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "D_Phone_Business3"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   61
         Left            =   -72270
         TabIndex        =   67
         Top             =   6300
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "D_Phone_Home1"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   56
         Left            =   -73320
         TabIndex        =   61
         Top             =   5880
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "D_Phone_Home2"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   57
         Left            =   -72795
         TabIndex        =   62
         Top             =   5880
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "D_Phone_Home3"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   58
         Left            =   -72270
         TabIndex        =   63
         Top             =   5880
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_Zip_code"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   55
         Left            =   -72270
         TabIndex        =   59
         Top             =   5040
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_State"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   54
         Left            =   -73215
         TabIndex        =   58
         Top             =   4620
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_County"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   53
         Left            =   -73215
         TabIndex        =   57
         Top             =   4200
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_City"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   52
         Left            =   -73215
         TabIndex        =   56
         Top             =   3780
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "ADD_Streetname"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   51
         Left            =   -73215
         TabIndex        =   55
         Top             =   3360
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "SS3"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   49
         Left            =   -72270
         TabIndex        =   53
         Top             =   2100
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "SS2"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   48
         Left            =   -72690
         TabIndex        =   52
         Top             =   2100
         Width           =   330
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "SS1"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   47
         Left            =   -73215
         TabIndex        =   51
         Top             =   2100
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "MI"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   46
         Left            =   -72060
         TabIndex        =   50
         Top             =   1680
         Width           =   330
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "First_Name"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   45
         Left            =   -73215
         TabIndex        =   49
         Top             =   1260
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0C0FF&
         DataField       =   "Last_Name"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   44
         Left            =   -73215
         TabIndex        =   48
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox t 
         DataField       =   "T1"
         DataSource      =   "Data5"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   88
         Left            =   -73320
         TabIndex        =   96
         Text            =   "12"
         Top             =   1575
         Width           =   330
      End
      Begin VB.TextBox t 
         DataField       =   "T2"
         DataSource      =   "Data5"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   89
         Left            =   -72900
         TabIndex        =   97
         Text            =   "59"
         Top             =   1575
         Width           =   330
      End
      Begin VB.TextBox t 
         DataField       =   "T3"
         DataSource      =   "Data5"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   90
         Left            =   -72480
         TabIndex        =   98
         Text            =   "59"
         Top             =   1575
         Width           =   330
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Intersecting_Street"
         DataSource      =   "Data5"
         Height          =   330
         Index           =   95
         Left            =   -72900
         TabIndex        =   104
         Top             =   4095
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "State"
         DataSource      =   "Data5"
         Height          =   330
         Index           =   93
         Left            =   -72900
         TabIndex        =   102
         Top             =   3255
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "County"
         DataSource      =   "Data5"
         Height          =   330
         Index           =   92
         Left            =   -72900
         TabIndex        =   101
         Top             =   2835
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "City"
         DataSource      =   "Data5"
         Height          =   330
         Index           =   91
         Left            =   -72900
         TabIndex        =   100
         Top             =   2415
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Street"
         DataSource      =   "Data5"
         Height          =   330
         Index           =   94
         Left            =   -72900
         TabIndex        =   103
         Top             =   3675
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Narrative_Description"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   96
         Left            =   -72900
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   105
         Top             =   4935
         Width           =   1590
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "Diagram_of_Scene"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   97
         Left            =   -72900
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   106
         Top             =   5460
         Width           =   1590
      End
      Begin VB.CheckBox c 
         Caption         =   "Investigated."
         DataField       =   "Investigated"
         DataSource      =   "Data5"
         Height          =   225
         Index           =   108
         Left            =   -71010
         TabIndex        =   108
         Top             =   1155
         Width           =   1800
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "Inv_Agency"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   98
         Left            =   -71115
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   109
         Top             =   1680
         Width           =   2325
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "Inv_Officer"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   99
         Left            =   -68700
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   110
         Top             =   1680
         Width           =   2325
      End
      Begin VB.CheckBox c 
         Caption         =   "Statements"
         DataField       =   "Statements"
         DataSource      =   "Data5"
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
         Index           =   109
         Left            =   -74475
         TabIndex        =   107
         Top             =   6195
         Width           =   1380
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Addresses"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   101
         Left            =   -69750
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   112
         Top             =   4935
         Width           =   3165
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Names"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   100
         Left            =   -69750
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   111
         Top             =   4410
         Width           =   3165
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Occupations"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   102
         Left            =   -69750
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   113
         Top             =   5460
         Width           =   3165
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Phone_Numbers"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   103
         Left            =   -69750
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   114
         Top             =   5985
         Width           =   3165
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Subject_Matter"
         DataSource      =   "Data5"
         Height          =   540
         Index           =   104
         Left            =   -71115
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   115
         Top             =   6825
         Width           =   4530
      End
      Begin VB.TextBox Text13 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7665
         Locked          =   -1  'True
         TabIndex        =   321
         Text            =   "59"
         Top             =   375
         Width           =   330
      End
      Begin VB.TextBox Text14 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   7245
         Locked          =   -1  'True
         TabIndex        =   320
         Text            =   "59"
         Top             =   375
         Width           =   330
      End
      Begin VB.TextBox Text15 
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   6825
         Locked          =   -1  'True
         TabIndex        =   319
         Text            =   "12"
         Top             =   375
         Width           =   330
      End
      Begin VB.CheckBox c 
         BackColor       =   &H8000000A&
         Caption         =   "Nature of Weather"
         DataField       =   "Nature_of_Weather"
         DataSource      =   "Data6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   110
         Left            =   -74370
         TabIndex        =   116
         Top             =   840
         Width           =   1905
      End
      Begin VB.CheckBox c 
         BackColor       =   &H8000000A&
         Caption         =   "Visibility"
         DataField       =   "Visibility"
         DataSource      =   "Data6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   111
         Left            =   -74370
         TabIndex        =   117
         Top             =   1155
         Width           =   1695
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Names"
         DataSource      =   "Data6"
         Height          =   540
         Index           =   105
         Left            =   -72900
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   118
         Text            =   "Data_entry.frx":2A5D
         Top             =   1995
         Width           =   3165
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Roadway"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   106
         Left            =   -71220
         TabIndex        =   119
         Text            =   "Text1"
         Top             =   2625
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "How_Many_Lanes"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   107
         Left            =   -71220
         TabIndex        =   120
         Text            =   "Text1"
         Top             =   3045
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "One_or_Two"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   108
         Left            =   -71220
         TabIndex        =   121
         Text            =   "Text1"
         Top             =   3465
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Interchange"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   109
         Left            =   -71220
         TabIndex        =   122
         Text            =   "Text1"
         Top             =   3885
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Description_of_Road"
         DataSource      =   "Data6"
         Height          =   540
         Index           =   110
         Left            =   -72900
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   123
         Text            =   "Data_entry.frx":2A63
         Top             =   4200
         Width           =   3165
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Locality"
         DataSource      =   "Data6"
         Height          =   540
         Index           =   111
         Left            =   -72900
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   124
         Text            =   "Data_entry.frx":2A69
         Top             =   4725
         Width           =   3165
      End
      Begin VB.Frame f 
         Caption         =   "Applicable traffic control devices:"
         Height          =   960
         Index           =   8
         Left            =   -74370
         TabIndex        =   318
         Top             =   5355
         Width           =   4635
         Begin VB.CheckBox c 
            Caption         =   "Stop signs"
            DataField       =   "Stop_S"
            DataSource      =   "Data6"
            Height          =   225
            Index           =   112
            Left            =   210
            TabIndex        =   125
            Top             =   315
            Width           =   1065
         End
         Begin VB.CheckBox c 
            Caption         =   "Stop lights"
            DataField       =   "Stop_L"
            DataSource      =   "Data6"
            Height          =   225
            Index           =   113
            Left            =   210
            TabIndex        =   126
            Top             =   630
            Width           =   1065
         End
         Begin VB.CheckBox c 
            Caption         =   "Warning light"
            DataField       =   "Warnning_L"
            DataSource      =   "Data6"
            Height          =   225
            Index           =   114
            Left            =   1470
            TabIndex        =   127
            Top             =   315
            Width           =   1380
         End
         Begin VB.CheckBox c 
            Caption         =   "Walk lights"
            DataField       =   "Walk_L"
            DataSource      =   "Data6"
            Height          =   225
            Index           =   115
            Left            =   1470
            TabIndex        =   128
            Top             =   630
            Width           =   1170
         End
         Begin VB.CheckBox c 
            Caption         =   "Traffic signing"
            DataField       =   "Traffic"
            DataSource      =   "Data6"
            Height          =   225
            Index           =   116
            Left            =   2940
            TabIndex        =   129
            Top             =   315
            Width           =   1380
         End
      End
      Begin VB.CheckBox c 
         Caption         =   "Artificial lighting"
         DataField       =   "Artificial"
         DataSource      =   "Data6"
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
         Index           =   117
         Left            =   -68805
         TabIndex        =   136
         Top             =   5985
         Width           =   2115
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "W_Did"
         DataSource      =   "Data6"
         Height          =   540
         Index           =   115
         Left            =   -69435
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   133
         Text            =   "Data_entry.frx":2A6F
         Top             =   2835
         Width           =   2955
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "W_D"
         DataSource      =   "Data6"
         Height          =   540
         Index           =   116
         Left            =   -69435
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   134
         Text            =   "Data_entry.frx":2A75
         Top             =   3780
         Width           =   2955
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Purpose"
         DataSource      =   "Data6"
         Height          =   540
         Index           =   117
         Left            =   -69435
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   135
         Text            =   "Data_entry.frx":2A7B
         Top             =   4830
         Width           =   2955
      End
      Begin VB.TextBox t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         DataField       =   "Speed_Impact"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   114
         Left            =   -71220
         TabIndex        =   132
         Text            =   "Text1"
         Top             =   7140
         Width           =   1485
      End
      Begin VB.TextBox t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         DataField       =   "Speed_of_All"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   113
         Left            =   -71220
         TabIndex        =   131
         Text            =   "Text1"
         Top             =   6825
         Width           =   1485
      End
      Begin VB.TextBox t 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         DataField       =   "Speed_Limit"
         DataSource      =   "Data6"
         Height          =   330
         Index           =   112
         Left            =   -71220
         TabIndex        =   130
         Text            =   "Text1"
         Top             =   6510
         Width           =   1485
      End
      Begin VB.CheckBox c 
         Caption         =   "Brakes applied"
         DataField       =   "PB"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   118
         Left            =   -67860
         TabIndex        =   137
         Top             =   6615
         Width           =   1380
      End
      Begin VB.CheckBox c 
         Caption         =   "Skidding"
         DataField       =   "PS"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   119
         Left            =   -67860
         TabIndex        =   138
         Top             =   6930
         Width           =   1380
      End
      Begin VB.CheckBox c 
         Caption         =   "Horn"
         DataField       =   "PH"
         DataSource      =   "Data6"
         Height          =   225
         Index           =   120
         Left            =   -67860
         TabIndex        =   139
         Top             =   7245
         Width           =   1380
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Phone1"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   122
         Left            =   -73110
         TabIndex        =   144
         Text            =   "123"
         Top             =   3255
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Phone2"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   123
         Left            =   -72585
         TabIndex        =   145
         Text            =   "456"
         Top             =   3255
         Width           =   435
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Phone3"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   124
         Left            =   -72060
         TabIndex        =   146
         Text            =   "7890"
         Top             =   3255
         Width           =   540
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Occupation"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   121
         Left            =   -73005
         TabIndex        =   143
         Text            =   "Text20"
         Top             =   2835
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Name"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   119
         Left            =   -73005
         TabIndex        =   141
         Text            =   "Text1"
         Top             =   1995
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Number"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   118
         Left            =   -73005
         TabIndex        =   140
         Text            =   "Text1"
         Top             =   1575
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "Address"
         DataSource      =   "Data7"
         Height          =   540
         Index           =   120
         Left            =   -73530
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   142
         Text            =   "Data_entry.frx":2A81
         Top             =   2310
         Width           =   2010
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "W_Located"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   125
         Left            =   -73005
         TabIndex        =   147
         Text            =   "Text20"
         Top             =   3990
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Injuries"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   127
         Left            =   -73005
         TabIndex        =   149
         Text            =   "Text20"
         Top             =   5145
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Relationship"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   126
         Left            =   -73005
         TabIndex        =   148
         Text            =   "Text20"
         Top             =   4725
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "R_P_All"
         DataSource      =   "Data7"
         Height          =   540
         Index           =   129
         Left            =   -71220
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   151
         Text            =   "Data_entry.frx":2A87
         Top             =   6510
         Width           =   1695
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "WGR"
         DataSource      =   "Data7"
         Height          =   540
         Index           =   128
         Left            =   -71220
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   150
         Text            =   "Data_entry.frx":2A8D
         Top             =   5985
         Width           =   1695
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Approximate_Time"
         DataSource      =   "Data7"
         Height          =   330
         Index           =   130
         Left            =   -71010
         TabIndex        =   153
         Text            =   "Text20"
         Top             =   7140
         Width           =   1485
      End
      Begin VB.CheckBox c 
         Caption         =   "Impact"
         DataField       =   "Impact"
         DataSource      =   "Data7"
         Height          =   225
         Index           =   121
         Left            =   -74370
         TabIndex        =   152
         Top             =   7140
         Width           =   1275
      End
      Begin VB.Frame f 
         Caption         =   "How was each vehicle removed from the scene:"
         Height          =   1590
         Index           =   13
         Left            =   -71325
         TabIndex        =   314
         Top             =   1995
         Width           =   4950
         Begin VB.TextBox t 
            BackColor       =   &H00FFC0C0&
            DataField       =   "RVP"
            DataSource      =   "Data7"
            Height          =   330
            Index           =   131
            Left            =   1155
            TabIndex        =   154
            Text            =   "Text1"
            Top             =   315
            Width           =   3690
         End
         Begin VB.TextBox t 
            BackColor       =   &H00FFC0C0&
            DataField       =   "RVD"
            DataSource      =   "Data7"
            Height          =   330
            Index           =   132
            Left            =   1155
            TabIndex        =   155
            Text            =   "Text1"
            Top             =   735
            Width           =   3690
         End
         Begin VB.TextBox t 
            BackColor       =   &H00FFC0C0&
            DataField       =   "RVO"
            DataSource      =   "Data7"
            Height          =   330
            Index           =   133
            Left            =   1155
            TabIndex        =   156
            Text            =   "Text1"
            Top             =   1155
            Width           =   3690
         End
         Begin VB.Label l 
            Caption         =   "Plaintiff:"
            Height          =   225
            Index           =   89
            Left            =   210
            TabIndex        =   317
            Top             =   420
            Width           =   645
         End
         Begin VB.Label l 
            Caption         =   "Defendant:"
            Height          =   225
            Index           =   90
            Left            =   210
            TabIndex        =   316
            Top             =   840
            Width           =   855
         End
         Begin VB.Label l 
            Caption         =   "Other:"
            Height          =   225
            Index           =   91
            Left            =   210
            TabIndex        =   315
            Top             =   1260
            Width           =   645
         End
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "WPVC"
         DataSource      =   "Data8"
         Height          =   330
         Index           =   134
         Left            =   -71115
         TabIndex        =   158
         Text            =   "Text1"
         Top             =   1155
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "WDEVE"
         DataSource      =   "Data8"
         Height          =   330
         Index           =   135
         Left            =   -71115
         TabIndex        =   159
         Text            =   "Text1"
         Top             =   1575
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "HFAWV"
         DataSource      =   "Data8"
         Height          =   330
         Index           =   136
         Left            =   -71115
         TabIndex        =   160
         Text            =   "Text1"
         Top             =   1995
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "DS"
         DataSource      =   "Data8"
         Height          =   540
         Index           =   137
         Left            =   -74265
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   161
         Text            =   "Data_entry.frx":2A93
         Top             =   2835
         Width           =   4635
      End
      Begin VB.Frame f 
         Caption         =   "Exterior damage to vehicles:"
         Height          =   2850
         Index           =   17
         Left            =   -74370
         TabIndex        =   307
         Top             =   3465
         Width           =   4635
         Begin VB.TextBox t 
            BackColor       =   &H00C0FFC0&
            DataField       =   "CPDD"
            DataSource      =   "Data8"
            Height          =   330
            Index           =   139
            Left            =   2835
            TabIndex        =   163
            Text            =   "Text1"
            Top             =   1155
            Width           =   1380
         End
         Begin VB.TextBox t 
            BackColor       =   &H00C0FFC0&
            DataField       =   "CDDD"
            DataSource      =   "Data8"
            Height          =   330
            Index           =   141
            Left            =   2835
            TabIndex        =   165
            Text            =   "Text1"
            Top             =   2415
            Width           =   1380
         End
         Begin VB.TextBox t 
            BackColor       =   &H00C0FFFF&
            DataField       =   "PDD"
            DataSource      =   "Data8"
            Height          =   540
            Index           =   138
            Left            =   2100
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   162
            Text            =   "Data_entry.frx":2A99
            Top             =   630
            Width           =   2430
         End
         Begin VB.TextBox t 
            BackColor       =   &H00C0FFFF&
            DataField       =   "DDD"
            DataSource      =   "Data8"
            Height          =   540
            Index           =   140
            Left            =   2100
            MultiLine       =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   164
            Text            =   "Data_entry.frx":2A9F
            Top             =   1890
            Width           =   2430
         End
         Begin VB.Label l 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   135
            Left            =   4305
            TabIndex        =   520
            Top             =   2415
            Width           =   225
         End
         Begin VB.Label l 
            Caption         =   "$"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   134
            Left            =   4305
            TabIndex        =   519
            Top             =   1155
            Width           =   225
         End
         Begin VB.Label l 
            Caption         =   "Cost of repair:"
            Height          =   225
            Index           =   101
            Left            =   315
            TabIndex        =   313
            Top             =   1260
            Width           =   1065
         End
         Begin VB.Label l 
            Caption         =   "Cost of repair:"
            Height          =   225
            Index           =   100
            Left            =   315
            TabIndex        =   312
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label l 
            Caption         =   "Plaintiff's vehicle:"
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
            Index           =   102
            Left            =   105
            TabIndex        =   311
            Top             =   420
            Width           =   1590
         End
         Begin VB.Label l 
            Caption         =   "Defendant's vehicle:"
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
            Index           =   103
            Left            =   105
            TabIndex        =   310
            Top             =   1680
            Width           =   1800
         End
         Begin VB.Label Label4 
            Caption         =   "Description of Damage:"
            Height          =   225
            Left            =   315
            TabIndex        =   309
            Top             =   735
            Width           =   1695
         End
         Begin VB.Label Label8 
            Caption         =   "Description of Damage:"
            Height          =   225
            Left            =   315
            TabIndex        =   308
            Top             =   1995
            Width           =   1695
         End
      End
      Begin VB.Frame f 
         Caption         =   "Interior damage to vehicles"
         Height          =   3375
         Index           =   18
         Left            =   -69540
         TabIndex        =   304
         Top             =   735
         Width           =   3165
         Begin VB.CheckBox c 
            Caption         =   "Bent steering wheel?"
            DataField       =   "BSW"
            DataSource      =   "Data8"
            Height          =   225
            Index           =   122
            Left            =   210
            TabIndex        =   166
            Top             =   735
            Width           =   1800
         End
         Begin VB.CheckBox c 
            Caption         =   "Broken seat?"
            DataField       =   "BS"
            DataSource      =   "Data8"
            Height          =   225
            Index           =   123
            Left            =   210
            TabIndex        =   167
            Top             =   1050
            Width           =   1695
         End
         Begin VB.CheckBox c 
            Caption         =   "Dents in dash or interior of vehicle?"
            DataField       =   "DDIV"
            DataSource      =   "Data8"
            Height          =   330
            Index           =   124
            Left            =   210
            TabIndex        =   168
            Top             =   1365
            Width           =   2850
         End
         Begin VB.CheckBox c 
            Caption         =   "Glasses come off?"
            DataField       =   "GC"
            DataSource      =   "Data8"
            Height          =   225
            Index           =   125
            Left            =   210
            TabIndex        =   169
            Top             =   2310
            Width           =   1800
         End
         Begin VB.CheckBox c 
            Caption         =   "Things fly around inside vehicle?"
            DataField       =   "TFAV"
            DataSource      =   "Data8"
            Height          =   225
            Index           =   126
            Left            =   210
            TabIndex        =   170
            Top             =   2625
            Width           =   2850
         End
         Begin VB.CheckBox c 
            Caption         =   "Other?"
            DataField       =   "O"
            DataSource      =   "Data8"
            Height          =   330
            Index           =   127
            Left            =   210
            TabIndex        =   171
            Top             =   2940
            Width           =   855
         End
         Begin VB.Line Line62 
            X1              =   105
            X2              =   2835
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Line Line63 
            X1              =   105
            X2              =   2205
            Y1              =   2100
            Y2              =   2100
         End
         Begin VB.Label l 
            Caption         =   "Describe damage to interior of vehicle:"
            Height          =   225
            Index           =   104
            Left            =   105
            TabIndex        =   306
            Top             =   315
            Width           =   2745
         End
         Begin VB.Label l 
            Caption         =   "Other events inside vehicle:"
            Height          =   225
            Index           =   105
            Left            =   105
            TabIndex        =   305
            Top             =   1890
            Width           =   2745
         End
      End
      Begin VB.CheckBox c 
         Caption         =   "Seatbelt?"
         DataField       =   "SEATBELT"
         DataSource      =   "Data8"
         Height          =   225
         Index           =   128
         Left            =   -69330
         TabIndex        =   172
         Top             =   4620
         Width           =   1065
      End
      Begin VB.CheckBox c 
         Caption         =   "Did body strike interior of vehicle?"
         DataField       =   "DBS"
         DataSource      =   "Data8"
         Height          =   330
         Index           =   129
         Left            =   -69330
         TabIndex        =   173
         Top             =   4935
         Width           =   2955
      End
      Begin VB.CheckBox c 
         Caption         =   "Loss of consciousness?"
         DataField       =   "LC"
         DataSource      =   "Data8"
         Height          =   225
         Index           =   130
         Left            =   -69330
         TabIndex        =   174
         Top             =   5355
         Width           =   2640
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFFF&
         DataField       =   "D_Blue"
         DataSource      =   "Data8"
         Height          =   540
         Index           =   142
         Left            =   -69435
         MultiLine       =   -1  'True
         ScrollBars      =   1  'Horizontal
         TabIndex        =   175
         Text            =   "Data_entry.frx":2AA5
         Top             =   6405
         Width           =   3060
      End
      Begin VB.TextBox t 
         Appearance      =   0  'Flat
         ForeColor       =   &H00FF0000&
         Height          =   540
         Index           =   143
         Left            =   -69435
         MultiLine       =   -1  'True
         TabIndex        =   303
         Text            =   "Data_entry.frx":2AAB
         Top             =   5775
         Width           =   2745
      End
      Begin VB.CheckBox c 
         BackColor       =   &H8000000A&
         Caption         =   "Immediately after injury?"
         DataField       =   "IAI"
         DataSource      =   "Data9"
         Height          =   225
         Index           =   131
         Left            =   -74370
         TabIndex        =   176
         Top             =   1050
         Width           =   2115
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "When_First_Seen"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   144
         Left            =   -70065
         TabIndex        =   177
         Text            =   "Text1"
         Top             =   1470
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0FFC0&
         DataField       =   "Several"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   145
         Left            =   -70065
         TabIndex        =   178
         Text            =   "Text1"
         Top             =   1890
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "HST"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   150
         Left            =   -73425
         TabIndex        =   183
         Text            =   "Text1"
         Top             =   5355
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "HCO"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   149
         Left            =   -73425
         TabIndex        =   182
         Text            =   "Text1"
         Top             =   4935
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "HC"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   148
         Left            =   -73425
         TabIndex        =   181
         Text            =   "Text1"
         Top             =   4515
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "HS"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   147
         Left            =   -73425
         TabIndex        =   180
         Text            =   "Text1"
         Top             =   4095
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00C0E0FF&
         DataField       =   "HN"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   146
         Left            =   -73425
         TabIndex        =   179
         Text            =   "Text1"
         Top             =   3675
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DST"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   155
         Left            =   -70695
         TabIndex        =   190
         Text            =   "Text1"
         Top             =   5355
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DCO"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   154
         Left            =   -70695
         TabIndex        =   189
         Text            =   "Text1"
         Top             =   4935
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DC"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   153
         Left            =   -70695
         TabIndex        =   188
         Text            =   "Text1"
         Top             =   4515
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DS"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   152
         Left            =   -70695
         TabIndex        =   187
         Text            =   "Text1"
         Top             =   4095
         Width           =   1485
      End
      Begin VB.TextBox t 
         BackColor       =   &H00FFFFC0&
         DataField       =   "DN"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   151
         Left            =   -70695
         TabIndex        =   186
         Text            =   "Text1"
         Top             =   3675
         Width           =   1485
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "HDB"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   4
         Left            =   -73425
         TabIndex        =   184
         Top             =   5775
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36825
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "Date"
         DataSource      =   "Data5"
         Height          =   330
         Index           =   3
         Left            =   -73320
         TabIndex        =   95
         Top             =   1155
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36824
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "Age"
         DataSource      =   "Data3"
         Height          =   330
         Index           =   1
         Left            =   1785
         TabIndex        =   20
         Top             =   6930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36824
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "Age"
         DataSource      =   "Data4"
         Height          =   330
         Index           =   2
         Left            =   -73215
         TabIndex        =   68
         Top             =   6930
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36824
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "HDE"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   5
         Left            =   -73425
         TabIndex        =   185
         Top             =   6195
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36825
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "DDB"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   6
         Left            =   -70695
         TabIndex        =   191
         Top             =   5775
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36825
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "DDE"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   7
         Left            =   -70695
         TabIndex        =   192
         Top             =   6195
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36825
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "TDB"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   8
         Left            =   -67860
         TabIndex        =   198
         Top             =   5775
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36825
      End
      Begin MSComCtl2.DTPicker dt 
         DataField       =   "TDE"
         DataSource      =   "Data9"
         Height          =   330
         Index           =   9
         Left            =   -67860
         TabIndex        =   199
         Top             =   6195
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   582
         _Version        =   393216
         Format          =   24969217
         CurrentDate     =   36825
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2430
         Left            =   -74475
         OleObjectBlob   =   "Data_entry.frx":2AEE
         TabIndex        =   540
         Top             =   5145
         Width           =   5370
      End
      Begin VB.Label Label29 
         Caption         =   "Detail each statement made by:"
         Height          =   225
         Left            =   -69225
         TabIndex        =   612
         Top             =   6090
         Width           =   2430
      End
      Begin VB.Label Label23 
         Caption         =   "Removal of plaintiff from scene:"
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
         Left            =   -71325
         TabIndex        =   607
         Top             =   4725
         Width           =   2850
      End
      Begin VB.Label Label22 
         Caption         =   "Treatment at Scene:"
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
         Left            =   -71325
         TabIndex        =   603
         Top             =   3780
         Width           =   2325
      End
      Begin VB.Line Line73 
         X1              =   -71220
         X2              =   -69540
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label Label18 
         Caption         =   "Signals and Lights:"
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
         Left            =   -71220
         TabIndex        =   599
         Top             =   735
         Width           =   2640
      End
      Begin VB.Line Line72 
         X1              =   -74475
         X2              =   -71535
         Y1              =   1470
         Y2              =   1470
      End
      Begin VB.Line Line71 
         X1              =   -74475
         X2              =   -72900
         Y1              =   945
         Y2              =   945
      End
      Begin VB.Label Label17 
         Caption         =   "About Passengers:"
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
         Left            =   -74475
         TabIndex        =   596
         Top             =   735
         Width           =   2640
      End
      Begin VB.Line Line70 
         X1              =   -69540
         X2              =   -66705
         Y1              =   2415
         Y2              =   2415
      End
      Begin VB.Line Line69 
         X1              =   -67965
         X2              =   -67965
         Y1              =   6510
         Y2              =   7455
      End
      Begin VB.Label Label16 
         Caption         =   "Evasive action:"
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
         Left            =   -69540
         TabIndex        =   593
         Top             =   6510
         Width           =   1485
      End
      Begin VB.Label Label15 
         Caption         =   "The Trip:"
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
         Left            =   -69540
         TabIndex        =   590
         Top             =   1680
         Width           =   1800
      End
      Begin VB.Line Line68 
         X1              =   -69225
         X2              =   -69225
         Y1              =   840
         Y2              =   1470
      End
      Begin VB.Line Line67 
         X1              =   -72375
         X2              =   -72375
         Y1              =   840
         Y2              =   1470
      End
      Begin VB.Label Label9 
         Caption         =   "Familiarity with route:"
         Height          =   225
         Left            =   -69120
         TabIndex        =   587
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label l 
         Caption         =   "Direction of Travel:"
         Height          =   225
         Index           =   170
         Left            =   -72165
         TabIndex        =   583
         Top             =   840
         Width           =   1380
      End
      Begin VB.Label Label2 
         Caption         =   "Photographs that should Be Taken:"
         Height          =   225
         Left            =   -71115
         TabIndex        =   574
         Top             =   3360
         Width           =   3375
      End
      Begin VB.Label l 
         Caption         =   "Have been taken:"
         Height          =   225
         Index           =   169
         Left            =   -71115
         TabIndex        =   573
         Top             =   2520
         Width           =   1485
      End
      Begin VB.Line Line66 
         X1              =   -74475
         X2              =   -73215
         Y1              =   6510
         Y2              =   6510
      End
      Begin VB.Label l 
         Caption         =   "Note: Not active in this version."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   168
         Left            =   -70275
         TabIndex        =   568
         Top             =   315
         Width           =   3375
      End
      Begin VB.Line Line65 
         BorderWidth     =   2
         X1              =   -74475
         X2              =   -66390
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Pay_Date"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   167
         Left            =   -69435
         TabIndex        =   566
         Top             =   3570
         Width           =   1380
      End
      Begin VB.Label l 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Paid in:"
         Height          =   225
         Index           =   165
         Left            =   -69435
         TabIndex        =   565
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label l 
         Caption         =   "Return Date: --->"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   145
         Left            =   -67965
         TabIndex        =   564
         Top             =   6510
         Width           =   1590
      End
      Begin VB.Label l 
         Caption         =   "Prior Employment History:"
         Height          =   435
         Index           =   164
         Left            =   -71430
         TabIndex        =   563
         Top             =   4725
         Width           =   1380
      End
      Begin VB.Line Line64 
         BorderWidth     =   2
         X1              =   -66390
         X2              =   -66390
         Y1              =   840
         Y2              =   3780
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   26
         X1              =   -67965
         X2              =   -66390
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   25
         X1              =   -72690
         X2              =   -69540
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   163
         Left            =   -72585
         TabIndex        =   560
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   162
         Left            =   -71010
         TabIndex        =   559
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   161
         Left            =   -67860
         TabIndex        =   558
         Top             =   3360
         Width           =   1380
      End
      Begin VB.Label l 
         Caption         =   "Input Area:"
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
         Index           =   160
         Left            =   -67860
         TabIndex        =   557
         Top             =   945
         Width           =   1380
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   23
         X1              =   -67965
         X2              =   -66390
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   21
         X1              =   -67965
         X2              =   -67965
         Y1              =   840
         Y2              =   3780
      End
      Begin VB.Line Line60 
         BorderWidth     =   2
         X1              =   -69540
         X2              =   -69540
         Y1              =   840
         Y2              =   3780
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   159
         Left            =   -71010
         TabIndex        =   553
         Top             =   1680
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   158
         Left            =   -71010
         TabIndex        =   552
         Top             =   1995
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   157
         Left            =   -71010
         TabIndex        =   551
         Top             =   2310
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   156
         Left            =   -71010
         TabIndex        =   550
         Top             =   2625
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   155
         Left            =   -71010
         TabIndex        =   549
         Top             =   2940
         Width           =   1380
      End
      Begin VB.Label l 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label2"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   154
         Left            =   -71010
         TabIndex        =   548
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label l 
         Caption         =   "Total Payment:"
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
         Index           =   153
         Left            =   -71010
         TabIndex        =   547
         Top             =   945
         Width           =   1380
      End
      Begin VB.Line Line59 
         BorderWidth     =   2
         X1              =   -71115
         X2              =   -71115
         Y1              =   840
         Y2              =   3780
      End
      Begin VB.Label l 
         Caption         =   "Total:"
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
         Index           =   152
         Left            =   -72585
         TabIndex        =   546
         Top             =   945
         Width           =   750
      End
      Begin VB.Label l 
         Caption         =   "Type:"
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
         Index           =   151
         Left            =   -74370
         TabIndex        =   545
         Top             =   945
         Width           =   1590
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   19
         X1              =   -74475
         X2              =   -66390
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   24
         X1              =   -74475
         X2              =   -74475
         Y1              =   3255
         Y2              =   840
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   22
         X1              =   -74475
         X2              =   -69540
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   20
         X1              =   -72690
         X2              =   -72690
         Y1              =   840
         Y2              =   3780
      End
      Begin VB.Image p 
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         Height          =   4110
         Left            =   -71745
         Stretch         =   -1  'True
         Top             =   945
         Width           =   5265
      End
      Begin VB.Label l 
         Caption         =   "Soundes and  Movies:"
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
         Index           =   149
         Left            =   -70275
         TabIndex        =   536
         Top             =   6615
         Width           =   2010
      End
      Begin VB.Line Line58 
         BorderWidth     =   3
         X1              =   -70275
         X2              =   -66495
         Y1              =   6510
         Y2              =   6510
      End
      Begin VB.Label l 
         Caption         =   "Multimedia Files such as: *.wav,*.gif,*.jpg,*.bmp,*.avi,..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   148
         Left            =   -71550
         TabIndex        =   534
         Top             =   285
         Width           =   4935
      End
      Begin VB.Line Line57 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -71115
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label l 
         Caption         =   "File Documents Diagrams"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   147
         Left            =   -74265
         TabIndex        =   533
         Top             =   210
         Width           =   3585
      End
      Begin VB.Line Line56 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -66810
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label l 
         Caption         =   "Relative Documments,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   146
         Left            =   -74265
         TabIndex        =   532
         Top             =   210
         Width           =   2745
      End
      Begin VB.Line Line21 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -69225
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line55 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -71745
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label l 
         Caption         =   "Relative Comments:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   144
         Left            =   -74265
         TabIndex        =   531
         Top             =   210
         Width           =   2745
      End
      Begin VB.Line Line54 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -71955
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label l 
         Caption         =   "Expense Account:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   143
         Left            =   -74265
         TabIndex        =   530
         Top             =   210
         Width           =   2745
      End
      Begin VB.Line Line52 
         BorderWidth     =   2
         X1              =   -70170
         X2              =   -66390
         Y1              =   6300
         Y2              =   6300
      End
      Begin VB.Line Line51 
         BorderWidth     =   2
         X1              =   -70170
         X2              =   -66390
         Y1              =   5670
         Y2              =   5670
      End
      Begin VB.Label l 
         Caption         =   "Total Medical Specials:"
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
         Index           =   142
         Left            =   -70065
         TabIndex        =   529
         Top             =   6510
         Width           =   2955
      End
      Begin VB.Label l 
         Caption         =   "Settlement Offers:"
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
         Index           =   141
         Left            =   -70065
         TabIndex        =   528
         Top             =   5880
         Width           =   1800
      End
      Begin VB.Label l 
         Caption         =   "Impression of Client:"
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
         Index           =   140
         Left            =   -70065
         TabIndex        =   527
         Top             =   5250
         Width           =   1905
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   18
         X1              =   -70170
         X2              =   -66390
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line38 
         BorderWidth     =   2
         X1              =   -70170
         X2              =   -70170
         Y1              =   1260
         Y2              =   7140
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   1
         X1              =   -74265
         X2              =   -70170
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label l 
         Caption         =   "Reason for wage loss:"
         Height          =   225
         Index           =   139
         Left            =   -74265
         TabIndex        =   524
         Top             =   2205
         Width           =   1695
      End
      Begin VB.Label l 
         Caption         =   "Rate of compensation:"
         Height          =   225
         Index           =   138
         Left            =   -74265
         TabIndex        =   523
         Top             =   1785
         Width           =   1905
      End
      Begin VB.Label l 
         Caption         =   "Dates Missed:"
         Height          =   225
         Index           =   137
         Left            =   -74265
         TabIndex        =   522
         Top             =   1365
         Width           =   1170
      End
      Begin VB.Line li 
         BorderColor     =   &H0000FFFF&
         Index           =   0
         X1              =   -74265
         X2              =   -72795
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Label l 
         Caption         =   "Wage Loss"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   136
         Left            =   -74265
         TabIndex        =   521
         Top             =   210
         Width           =   2745
      End
      Begin VB.Line Line15 
         X1              =   6615
         X2              =   7875
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line14 
         X1              =   6615
         X2              =   7875
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line7 
         X1              =   3570
         X2              =   4620
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Line Line12 
         X1              =   3570
         X2              =   5565
         Y1              =   5565
         Y2              =   5565
      End
      Begin VB.Line Line5 
         X1              =   630
         X2              =   1995
         Y1              =   5775
         Y2              =   5775
      End
      Begin VB.Line Line3 
         X1              =   630
         X2              =   1575
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000FFFF&
         X1              =   735
         X2              =   1995
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line18 
         X1              =   -68385
         X2              =   -67230
         Y1              =   3780
         Y2              =   3780
      End
      Begin VB.Line Line17 
         X1              =   -68385
         X2              =   -67125
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line19 
         X1              =   -71430
         X2              =   -69540
         Y1              =   5565
         Y2              =   5565
      End
      Begin VB.Line Line20 
         X1              =   -71430
         X2              =   -70275
         Y1              =   3255
         Y2              =   3255
      End
      Begin VB.Line Line24 
         X1              =   -74370
         X2              =   -73005
         Y1              =   5775
         Y2              =   5775
      End
      Begin VB.Line Line22 
         X1              =   -74370
         X2              =   -73530
         Y1              =   2835
         Y2              =   2835
      End
      Begin VB.Line Line16 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -72795
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line36 
         X1              =   -71115
         X2              =   -69015
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line33 
         X1              =   -74370
         X2              =   -73320
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Line Line32 
         X1              =   -74370
         X2              =   -73635
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line45 
         X1              =   -69540
         X2              =   -67650
         Y1              =   5775
         Y2              =   5775
      End
      Begin VB.Line Line49 
         X1              =   -74370
         X2              =   -73740
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line41 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -71220
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line44 
         X1              =   -74370
         X2              =   -72060
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line li 
         Index           =   15
         X1              =   -71220
         X2              =   -69435
         Y1              =   1890
         Y2              =   1890
      End
      Begin VB.Line li 
         Index           =   16
         X1              =   -69225
         X2              =   -67125
         Y1              =   5985
         Y2              =   5985
      End
      Begin VB.Line Line53 
         X1              =   -74370
         X2              =   -71115
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line50 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -71325
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line li 
         Index           =   7
         X1              =   -74370
         X2              =   -72690
         Y1              =   1050
         Y2              =   1050
      End
      Begin VB.Line Line61 
         BorderColor     =   &H0000FFFF&
         X1              =   -74265
         X2              =   -71325
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line li 
         Index           =   6
         X1              =   -69015
         X2              =   -67965
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line li 
         Index           =   5
         X1              =   -71745
         X2              =   -70905
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line li 
         Index           =   2
         X1              =   -74475
         X2              =   -73530
         Y1              =   3570
         Y2              =   3570
      End
      Begin VB.Line li 
         BorderColor     =   &H0000FFFF&
         Index           =   4
         X1              =   -74265
         X2              =   -67125
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line li 
         BorderColor     =   &H0000FFFF&
         Index           =   3
         X1              =   -74265
         X2              =   -71535
         Y1              =   3150
         Y2              =   3150
      End
      Begin VB.Label l 
         Caption         =   "Date End:"
         Height          =   225
         Index           =   133
         Left            =   -69015
         TabIndex        =   518
         Top             =   6300
         Width           =   1170
      End
      Begin VB.Label l 
         Caption         =   "Date Begin:"
         Height          =   225
         Index           =   132
         Left            =   -69015
         TabIndex        =   517
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Label l 
         Caption         =   "Street Name:"
         Height          =   225
         Index           =   131
         Left            =   -69015
         TabIndex        =   516
         Top             =   4200
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "City:"
         Height          =   225
         Index           =   130
         Left            =   -69015
         TabIndex        =   515
         Top             =   4620
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "County:"
         Height          =   225
         Index           =   129
         Left            =   -69015
         TabIndex        =   514
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "State:"
         Height          =   225
         Index           =   128
         Left            =   -69015
         TabIndex        =   513
         Top             =   5460
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Name:"
         Height          =   225
         Index           =   127
         Left            =   -69015
         TabIndex        =   512
         Top             =   3780
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Therapists:"
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
         Index           =   126
         Left            =   -69015
         TabIndex        =   511
         Top             =   3360
         Width           =   2430
      End
      Begin VB.Label l 
         Caption         =   "Interview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   35
         Left            =   735
         TabIndex        =   510
         Top             =   210
         Width           =   2115
      End
      Begin VB.Label Label11 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2205
         TabIndex        =   509
         Top             =   1995
         Width           =   120
      End
      Begin VB.Label Label12 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2625
         TabIndex        =   508
         Top             =   1995
         Width           =   120
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   630
         X2              =   3360
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label l 
         Caption         =   "Address:"
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
         Index           =   4
         Left            =   630
         TabIndex        =   507
         Top             =   2625
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Street Name:"
         Height          =   225
         Index           =   6
         Left            =   630
         TabIndex        =   506
         Top             =   3465
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "City:"
         Height          =   225
         Index           =   7
         Left            =   630
         TabIndex        =   505
         Top             =   3885
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "County:"
         Height          =   225
         Index           =   8
         Left            =   630
         TabIndex        =   504
         Top             =   4305
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "State:"
         Height          =   225
         Index           =   9
         Left            =   630
         TabIndex        =   503
         Top             =   4725
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Zip Code:"
         Height          =   225
         Index           =   10
         Left            =   630
         TabIndex        =   502
         Top             =   5145
         Width           =   855
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   630
         X2              =   3360
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label l 
         Caption         =   "Phone Number:"
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
         Index           =   11
         Left            =   630
         TabIndex        =   501
         Top             =   5565
         Width           =   1485
      End
      Begin VB.Label Label20 
         Caption         =   "Label20"
         Height          =   120
         Left            =   2205
         TabIndex        =   500
         Top             =   3885
         Width           =   120
      End
      Begin VB.Line Line6 
         BorderWidth     =   2
         X1              =   630
         X2              =   3360
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Label l 
         Caption         =   "Age:"
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
         Index           =   12
         Left            =   735
         TabIndex        =   499
         Top             =   7035
         Width           =   855
      End
      Begin VB.Line Line8 
         BorderWidth     =   2
         X1              =   3360
         X2              =   3360
         Y1              =   735
         Y2              =   7245
      End
      Begin VB.Label Label25 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2625
         TabIndex        =   498
         Top             =   5775
         Width           =   120
      End
      Begin VB.Label Label26 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2100
         TabIndex        =   497
         Top             =   5775
         Width           =   120
      End
      Begin VB.Label Label27 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2625
         TabIndex        =   496
         Top             =   6195
         Width           =   120
      End
      Begin VB.Label Label28 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2100
         TabIndex        =   495
         Top             =   6195
         Width           =   120
      End
      Begin VB.Label l 
         Caption         =   "First Name:"
         Height          =   225
         Index           =   25
         Left            =   3570
         TabIndex        =   494
         Top             =   1785
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Last Name:"
         Height          =   225
         Index           =   26
         Left            =   3570
         TabIndex        =   493
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Occupation:"
         Height          =   225
         Index           =   24
         Left            =   3570
         TabIndex        =   492
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Phone:"
         Height          =   225
         Index           =   23
         Left            =   3570
         TabIndex        =   491
         Top             =   2625
         Width           =   750
      End
      Begin VB.Label Label34 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5775
         TabIndex        =   490
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label Label35 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5250
         TabIndex        =   489
         Top             =   2415
         Width           =   120
      End
      Begin VB.Line Line10 
         BorderWidth     =   2
         X1              =   3360
         X2              =   6510
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line9 
         X1              =   3570
         X2              =   4935
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Label l 
         Caption         =   "Number:"
         Height          =   225
         Index           =   5
         Left            =   630
         TabIndex        =   488
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label l 
         Caption         =   "Occupation:"
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
         Index           =   22
         Left            =   3570
         TabIndex        =   487
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label l 
         Caption         =   "Prior Employment History:"
         Height          =   435
         Index           =   18
         Left            =   3570
         TabIndex        =   486
         Top             =   4725
         Width           =   1380
      End
      Begin VB.Label l 
         Caption         =   "How Long Employed:"
         Height          =   225
         Index           =   19
         Left            =   3570
         TabIndex        =   485
         Top             =   4305
         Width           =   1590
      End
      Begin VB.Label l 
         Caption         =   "Nature of Work:"
         Height          =   225
         Index           =   20
         Left            =   3570
         TabIndex        =   484
         Top             =   3885
         Width           =   1485
      End
      Begin VB.Label l 
         Caption         =   "By Whom Employed:"
         Height          =   225
         Index           =   21
         Left            =   3570
         TabIndex        =   483
         Top             =   3465
         Width           =   1485
      End
      Begin VB.Line Line11 
         BorderWidth     =   2
         X1              =   3360
         X2              =   6510
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Label l 
         Caption         =   "Insurance Companies Name:"
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
         Index           =   17
         Left            =   3570
         TabIndex        =   482
         Top             =   5355
         Width           =   2220
      End
      Begin VB.Label l 
         Caption         =   "Limits:"
         Height          =   225
         Index           =   13
         Left            =   3570
         TabIndex        =   481
         Top             =   7035
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Medical:"
         Height          =   225
         Index           =   14
         Left            =   3570
         TabIndex        =   480
         Top             =   6615
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "PIP, If applicable:"
         Height          =   225
         Index           =   15
         Left            =   3570
         TabIndex        =   479
         Top             =   6195
         Width           =   1275
      End
      Begin VB.Label l 
         Caption         =   "Liability:"
         Height          =   225
         Index           =   16
         Left            =   3570
         TabIndex        =   478
         Top             =   5775
         Width           =   1065
      End
      Begin VB.Line Line13 
         BorderWidth     =   2
         X1              =   6510
         X2              =   6510
         Y1              =   840
         Y2              =   7245
      End
      Begin VB.Label l 
         Caption         =   "BI Adjusters:"
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
         Index           =   27
         Left            =   6615
         TabIndex        =   477
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label l 
         Caption         =   "Name:"
         Height          =   225
         Index           =   28
         Left            =   6615
         TabIndex        =   476
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Phone Number:"
         Height          =   225
         Index           =   29
         Left            =   6615
         TabIndex        =   475
         Top             =   2205
         Width           =   1380
      End
      Begin VB.Label Label46 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7455
         TabIndex        =   474
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label Label47 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7980
         TabIndex        =   473
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label l 
         Caption         =   "Claim#"
         Height          =   225
         Index           =   30
         Left            =   6615
         TabIndex        =   472
         Top             =   2940
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "PD Adjusters:"
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
         Index           =   31
         Left            =   6615
         TabIndex        =   471
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label l 
         Caption         =   "Name:"
         Height          =   225
         Index           =   32
         Left            =   6615
         TabIndex        =   470
         Top             =   4305
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Phone Number:"
         Height          =   225
         Index           =   33
         Left            =   6615
         TabIndex        =   469
         Top             =   4935
         Width           =   1380
      End
      Begin VB.Label Label52 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7455
         TabIndex        =   468
         Top             =   5145
         Width           =   120
      End
      Begin VB.Label Label53 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7980
         TabIndex        =   467
         Top             =   5145
         Width           =   120
      End
      Begin VB.Label l 
         Caption         =   "Claim#"
         Height          =   225
         Index           =   34
         Left            =   6615
         TabIndex        =   466
         Top             =   5670
         Width           =   960
      End
      Begin VB.Label Label55 
         Caption         =   "Defendant"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   435
         Left            =   -74265
         TabIndex        =   465
         Top             =   210
         Width           =   1380
      End
      Begin VB.Label Label56 
         Caption         =   "Claim#"
         Height          =   225
         Left            =   -68385
         TabIndex        =   464
         Top             =   5670
         Width           =   960
      End
      Begin VB.Label Label57 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67020
         TabIndex        =   463
         Top             =   5145
         Width           =   120
      End
      Begin VB.Label Label58 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67545
         TabIndex        =   462
         Top             =   5145
         Width           =   120
      End
      Begin VB.Label Label59 
         Caption         =   "Phone Number:"
         Height          =   225
         Left            =   -68385
         TabIndex        =   461
         Top             =   4935
         Width           =   1380
      End
      Begin VB.Label Label60 
         Caption         =   "Name:"
         Height          =   225
         Left            =   -68385
         TabIndex        =   460
         Top             =   4305
         Width           =   960
      End
      Begin VB.Label Label61 
         Caption         =   "PD Adjusters:"
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
         Left            =   -68385
         TabIndex        =   459
         Top             =   3570
         Width           =   1590
      End
      Begin VB.Label Label62 
         Caption         =   "Claim#"
         Height          =   225
         Left            =   -68385
         TabIndex        =   458
         Top             =   2940
         Width           =   960
      End
      Begin VB.Label Label63 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67020
         TabIndex        =   457
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label Label64 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -67545
         TabIndex        =   456
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label Label65 
         Caption         =   "Phone Number:"
         Height          =   225
         Left            =   -68385
         TabIndex        =   455
         Top             =   2205
         Width           =   1380
      End
      Begin VB.Label Label66 
         Caption         =   "Name:"
         Height          =   225
         Left            =   -68385
         TabIndex        =   454
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label67 
         Caption         =   "BI Adjusters:"
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
         Left            =   -68385
         TabIndex        =   453
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label68 
         Caption         =   "Liability:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   452
         Top             =   5775
         Width           =   1065
      End
      Begin VB.Label Label69 
         Caption         =   "PIP, If applicable:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   451
         Top             =   6195
         Width           =   1275
      End
      Begin VB.Label Label70 
         Caption         =   "Medical:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   450
         Top             =   6615
         Width           =   855
      End
      Begin VB.Label Label71 
         Caption         =   "Limits:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   449
         Top             =   7035
         Width           =   855
      End
      Begin VB.Label Label72 
         Caption         =   "Insurance Companies Name:"
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
         Left            =   -71430
         TabIndex        =   448
         Top             =   5355
         Width           =   2220
      End
      Begin VB.Label Label73 
         Caption         =   "By Whom Employed:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   447
         Top             =   3465
         Width           =   1485
      End
      Begin VB.Label Label74 
         Caption         =   "Nature of Work:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   446
         Top             =   3885
         Width           =   1485
      End
      Begin VB.Label Label75 
         Caption         =   "How Long Employed:"
         Height          =   225
         Left            =   -71430
         TabIndex        =   445
         Top             =   4305
         Width           =   1590
      End
      Begin VB.Label Label77 
         Caption         =   "Occupation:"
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
         Left            =   -71430
         TabIndex        =   444
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label78 
         Caption         =   "Number:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   443
         Top             =   3045
         Width           =   1065
      End
      Begin VB.Label Label79 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   -69750
         TabIndex        =   442
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label Label80 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   -69225
         TabIndex        =   441
         Top             =   2415
         Width           =   120
      End
      Begin VB.Label Label81 
         Caption         =   "Phone:"
         Height          =   225
         Index           =   0
         Left            =   -71430
         TabIndex        =   440
         Top             =   2625
         Width           =   750
      End
      Begin VB.Label Label82 
         Caption         =   "Occupation:"
         Height          =   225
         Index           =   0
         Left            =   -71430
         TabIndex        =   439
         Top             =   2205
         Width           =   855
      End
      Begin VB.Label Label83 
         Caption         =   "Last Name:"
         Height          =   225
         Index           =   0
         Left            =   -71430
         TabIndex        =   438
         Top             =   1365
         Width           =   855
      End
      Begin VB.Label Label84 
         Caption         =   "First Name:"
         Height          =   225
         Index           =   0
         Left            =   -71430
         TabIndex        =   437
         Top             =   1785
         Width           =   855
      End
      Begin VB.Label Label85 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72900
         TabIndex        =   436
         Top             =   6195
         Width           =   120
      End
      Begin VB.Label Label86 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72375
         TabIndex        =   435
         Top             =   6195
         Width           =   120
      End
      Begin VB.Label Label87 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72900
         TabIndex        =   434
         Top             =   5775
         Width           =   120
      End
      Begin VB.Label Label88 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72375
         TabIndex        =   433
         Top             =   5775
         Width           =   120
      End
      Begin VB.Label Label89 
         Caption         =   "Age:"
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
         Left            =   -74265
         TabIndex        =   432
         Top             =   7035
         Width           =   855
      End
      Begin VB.Label Label90 
         Caption         =   "Phone Number:"
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
         TabIndex        =   431
         Top             =   5565
         Width           =   1485
      End
      Begin VB.Label Label91 
         Caption         =   "Zip Code:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   430
         Top             =   5145
         Width           =   855
      End
      Begin VB.Label Label92 
         Caption         =   "State:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   429
         Top             =   4725
         Width           =   855
      End
      Begin VB.Label Label93 
         Caption         =   "County:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   428
         Top             =   4305
         Width           =   855
      End
      Begin VB.Label Label94 
         Caption         =   "City:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   427
         Top             =   3885
         Width           =   855
      End
      Begin VB.Label Label95 
         Caption         =   "Street Name:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   426
         Top             =   3465
         Width           =   960
      End
      Begin VB.Label Label96 
         Caption         =   "Address:"
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
         TabIndex        =   425
         Top             =   2625
         Width           =   960
      End
      Begin VB.Line Line23 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -71640
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label97 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72375
         TabIndex        =   424
         Top             =   1995
         Width           =   120
      End
      Begin VB.Label Label98 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72795
         TabIndex        =   423
         Top             =   1995
         Width           =   120
      End
      Begin VB.Label Label99 
         Caption         =   "SS#:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   422
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label100 
         BackColor       =   &H80000004&
         Caption         =   "MI:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   421
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label101 
         Caption         =   "First Name:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   420
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label Label102 
         Caption         =   "Last Name:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   419
         Top             =   840
         Width           =   855
      End
      Begin VB.Line Line25 
         X1              =   -71430
         X2              =   -70275
         Y1              =   1155
         Y2              =   1155
      End
      Begin VB.Line Line26 
         BorderWidth     =   2
         X1              =   -71640
         X2              =   -71640
         Y1              =   840
         Y2              =   7245
      End
      Begin VB.Line Line27 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -71640
         Y1              =   6720
         Y2              =   6720
      End
      Begin VB.Line Line28 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -71640
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Line Line29 
         BorderWidth     =   2
         X1              =   -71640
         X2              =   -68490
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line30 
         BorderWidth     =   2
         X1              =   -71640
         X2              =   -68490
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line Line31 
         BorderWidth     =   2
         X1              =   -68490
         X2              =   -68490
         Y1              =   7245
         Y2              =   840
      End
      Begin VB.Label l 
         Caption         =   "Last Name:"
         Height          =   225
         Index           =   0
         Left            =   630
         TabIndex        =   418
         Top             =   840
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "First Name:"
         Height          =   225
         Index           =   1
         Left            =   630
         TabIndex        =   417
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label l 
         BackColor       =   &H80000004&
         Caption         =   "MI:"
         Height          =   225
         Index           =   2
         Left            =   630
         TabIndex        =   416
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "SS#:"
         Height          =   225
         Index           =   3
         Left            =   630
         TabIndex        =   415
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Facts Relating to the Injury or Occurrence"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   -74265
         TabIndex        =   414
         Top             =   210
         Width           =   5265
      End
      Begin VB.Label l 
         Caption         =   "When:"
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
         Index           =   50
         Left            =   -74370
         TabIndex        =   413
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label l 
         Caption         =   "Date: "
         Height          =   225
         Index           =   55
         Left            =   -74370
         TabIndex        =   412
         Top             =   1260
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Time:"
         Height          =   225
         Index           =   54
         Left            =   -74370
         TabIndex        =   411
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   ":"
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
         Left            =   -72990
         TabIndex        =   410
         Top             =   1515
         Width           =   120
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " AM"
         DataField       =   "TW"
         DataSource      =   "Data5"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   -71925
         TabIndex        =   99
         Top             =   1575
         Width           =   390
      End
      Begin VB.Label Label14 
         Caption         =   ":"
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
         Left            =   -72585
         TabIndex        =   409
         Top             =   1515
         Width           =   120
      End
      Begin VB.Label l 
         Caption         =   "Intersecting Street:"
         Height          =   225
         Index           =   36
         Left            =   -74370
         TabIndex        =   408
         Top             =   4200
         Width           =   1380
      End
      Begin VB.Label l 
         Caption         =   "State:"
         Height          =   225
         Index           =   37
         Left            =   -74370
         TabIndex        =   407
         Top             =   3360
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "County:"
         Height          =   225
         Index           =   38
         Left            =   -74370
         TabIndex        =   406
         Top             =   2940
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "City:"
         Height          =   225
         Index           =   39
         Left            =   -74370
         TabIndex        =   405
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Street Name:"
         Height          =   225
         Index           =   40
         Left            =   -74370
         TabIndex        =   404
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Location:"
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
         Index           =   41
         Left            =   -74370
         TabIndex        =   403
         Top             =   2100
         Width           =   960
      End
      Begin VB.Line Line34 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -71220
         Y1              =   1995
         Y2              =   1995
      End
      Begin VB.Line Line35 
         BorderWidth     =   2
         X1              =   -71220
         X2              =   -71220
         Y1              =   840
         Y2              =   7350
      End
      Begin VB.Label l 
         Caption         =   "Description of Injury or Occurence:"
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
         Index           =   51
         Left            =   -74370
         TabIndex        =   402
         Top             =   4620
         Width           =   3060
      End
      Begin VB.Line Line37 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -71220
         Y1              =   4515
         Y2              =   4515
      End
      Begin VB.Label l 
         Caption         =   "Narrative Description:"
         Height          =   225
         Index           =   53
         Left            =   -74475
         TabIndex        =   401
         Top             =   5040
         Width           =   1590
      End
      Begin VB.Label l 
         Caption         =   "Diagram of Scene:"
         Height          =   225
         Index           =   52
         Left            =   -74475
         TabIndex        =   400
         Top             =   5670
         Width           =   1590
      End
      Begin VB.Label l 
         Caption         =   "Investigation or Incident Reports:"
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
         Index           =   49
         Left            =   -71115
         TabIndex        =   399
         Top             =   840
         Width           =   2535
      End
      Begin VB.Label Label19 
         Caption         =   "Names of Investigating Agency:"
         Height          =   225
         Left            =   -71115
         TabIndex        =   398
         Top             =   1470
         Width           =   2325
      End
      Begin VB.Label Label21 
         Caption         =   "Name of Investigating Officer:"
         Height          =   225
         Left            =   -68700
         TabIndex        =   397
         Top             =   1470
         Width           =   2220
      End
      Begin VB.Line Line39 
         BorderWidth     =   2
         X1              =   -71220
         X2              =   -66390
         Y1              =   2310
         Y2              =   2310
      End
      Begin VB.Label l 
         Caption         =   "Addresses:"
         Height          =   225
         Index           =   45
         Left            =   -71115
         TabIndex        =   396
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Names:"
         Height          =   225
         Index           =   44
         Left            =   -71115
         TabIndex        =   395
         Top             =   4515
         Width           =   645
      End
      Begin VB.Line Line42 
         BorderWidth     =   2
         X1              =   -74475
         X2              =   -71220
         Y1              =   6090
         Y2              =   6090
      End
      Begin VB.Label l 
         Caption         =   "Occupations:"
         Height          =   225
         Index           =   46
         Left            =   -71115
         TabIndex        =   394
         Top             =   5565
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Phone numbers:"
         Height          =   225
         Index           =   47
         Left            =   -71115
         TabIndex        =   393
         Top             =   6090
         Width           =   1170
      End
      Begin VB.Label l 
         Caption         =   "Subject Matter of Information:"
         Height          =   225
         Index           =   48
         Left            =   -71115
         TabIndex        =   392
         Top             =   6615
         Width           =   2850
      End
      Begin VB.Line Line40 
         BorderWidth     =   2
         X1              =   -71220
         X2              =   -66390
         Y1              =   4095
         Y2              =   4095
      End
      Begin VB.Label l 
         Caption         =   "Automobile Collision (1)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   42
         Left            =   -74265
         TabIndex        =   391
         Top             =   210
         Width           =   5265
      End
      Begin VB.Label Label36 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7140
         TabIndex        =   390
         Top             =   105
         Width           =   855
      End
      Begin VB.Label Label37 
         Caption         =   "Time Created:"
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
         Left            =   5250
         TabIndex        =   389
         Top             =   420
         Width           =   1275
      End
      Begin VB.Label Label38 
         Caption         =   "Date Created:"
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
         Left            =   5250
         TabIndex        =   388
         Top             =   105
         Width           =   1275
      End
      Begin VB.Label Label39 
         Caption         =   ":"
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
         Left            =   7560
         TabIndex        =   387
         Top             =   315
         Width           =   120
      End
      Begin VB.Label Label40 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   8220
         TabIndex        =   322
         Top             =   375
         Width           =   390
      End
      Begin VB.Label Label41 
         Caption         =   ":"
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
         Left            =   7155
         TabIndex        =   386
         Top             =   315
         Width           =   120
      End
      Begin VB.Line Line43 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -66495
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Label l 
         Caption         =   "Streets or roads involved:"
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
         Index           =   43
         Left            =   -74370
         TabIndex        =   385
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Label Label24 
         Caption         =   "Names:"
         Height          =   225
         Left            =   -74370
         TabIndex        =   384
         Top             =   2100
         Width           =   645
      End
      Begin VB.Label l 
         Caption         =   "Roadway Surface:"
         Height          =   225
         Index           =   59
         Left            =   -74370
         TabIndex        =   383
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label l 
         Caption         =   "How many lanes did each roadway have?"
         Height          =   225
         Index           =   60
         Left            =   -74370
         TabIndex        =   382
         Top             =   3150
         Width           =   3060
      End
      Begin VB.Label l 
         Caption         =   "One-way or two-way?"
         Height          =   225
         Index           =   61
         Left            =   -74370
         TabIndex        =   381
         Top             =   3570
         Width           =   1800
      End
      Begin VB.Label l 
         Caption         =   "Interchange:"
         Height          =   225
         Index           =   62
         Left            =   -74370
         TabIndex        =   380
         Top             =   3990
         Width           =   1065
      End
      Begin VB.Label l 
         Caption         =   "Description of road:"
         Height          =   225
         Index           =   63
         Left            =   -74370
         TabIndex        =   379
         Top             =   4410
         Width           =   2115
      End
      Begin VB.Label l 
         Caption         =   "Locality:"
         Height          =   225
         Index           =   64
         Left            =   -74370
         TabIndex        =   378
         Top             =   4935
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Obstructions to view:"
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
         Index           =   65
         Left            =   -69540
         TabIndex        =   377
         Top             =   5565
         Width           =   1905
      End
      Begin VB.Line Line46 
         BorderWidth     =   2
         X1              =   -69645
         X2              =   -69645
         Y1              =   1575
         Y2              =   7455
      End
      Begin VB.Line Line47 
         BorderWidth     =   2
         X1              =   -74370
         X2              =   -66495
         Y1              =   6405
         Y2              =   6405
      End
      Begin VB.Label l 
         Caption         =   "Where did the trip begin?"
         Height          =   225
         Index           =   67
         Left            =   -69540
         TabIndex        =   376
         Top             =   2520
         Width           =   2850
      End
      Begin VB.Label l 
         Caption         =   "Witnesses:"
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
         Index           =   70
         Left            =   -71115
         TabIndex        =   375
         Top             =   4200
         Width           =   1065
      End
      Begin VB.Label l 
         Caption         =   "Where was destination?"
         Height          =   225
         Index           =   66
         Left            =   -69540
         TabIndex        =   374
         Top             =   3465
         Width           =   2850
      End
      Begin VB.Label l 
         Caption         =   "Purpose?"
         Height          =   225
         Index           =   68
         Left            =   -69540
         TabIndex        =   373
         Top             =   4515
         Width           =   2850
      End
      Begin VB.Line Line48 
         BorderWidth     =   2
         X1              =   -69645
         X2              =   -66495
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Label l 
         Alignment       =   1  'Right Justify
         Caption         =   "Speeds of vehicles at impact:"
         Height          =   225
         Index           =   69
         Left            =   -73635
         TabIndex        =   372
         Top             =   7245
         Width           =   2220
      End
      Begin VB.Label l 
         Alignment       =   1  'Right Justify
         Caption         =   "Speeds of all vehicles prior to braking:"
         Height          =   225
         Index           =   71
         Left            =   -74475
         TabIndex        =   371
         Top             =   6930
         Width           =   3060
      End
      Begin VB.Label l 
         Alignment       =   1  'Right Justify
         Caption         =   "Speed limit of roadways:"
         Height          =   225
         Index           =   72
         Left            =   -73215
         TabIndex        =   370
         Top             =   6615
         Width           =   1800
      End
      Begin VB.Label l 
         Caption         =   "Speed:"
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
         Index           =   73
         Left            =   -74370
         TabIndex        =   369
         Top             =   6510
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Automobile Collision (2)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   74
         Left            =   -74265
         TabIndex        =   368
         Top             =   210
         Width           =   5265
      End
      Begin VB.Label Label1 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72690
         TabIndex        =   367
         Top             =   3150
         Width           =   120
      End
      Begin VB.Label Label3 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   -72165
         TabIndex        =   366
         Top             =   3150
         Width           =   120
      End
      Begin VB.Label l 
         Caption         =   "Phone:"
         Height          =   225
         Index           =   75
         Left            =   -74370
         TabIndex        =   365
         Top             =   3255
         Width           =   750
      End
      Begin VB.Label l 
         Caption         =   "Occupation:"
         Height          =   225
         Index           =   76
         Left            =   -74370
         TabIndex        =   364
         Top             =   2835
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Name:"
         Height          =   225
         Index           =   77
         Left            =   -74370
         TabIndex        =   363
         Top             =   1995
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Address:"
         Height          =   225
         Index           =   78
         Left            =   -74370
         TabIndex        =   362
         Top             =   2415
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Number:"
         Height          =   225
         Index           =   79
         Left            =   -74370
         TabIndex        =   361
         Top             =   1575
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Where located in automobile?"
         Height          =   225
         Index           =   80
         Left            =   -74475
         TabIndex        =   360
         Top             =   3675
         Width           =   2325
      End
      Begin VB.Label l 
         Caption         =   "Injuries:"
         Height          =   225
         Index           =   81
         Left            =   -74370
         TabIndex        =   359
         Top             =   5145
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Relationship to driver or other occupants:"
         Height          =   225
         Index           =   82
         Left            =   -74475
         TabIndex        =   358
         Top             =   4410
         Width           =   2955
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   11
         X1              =   -71430
         X2              =   -71430
         Y1              =   945
         Y2              =   5565
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   10
         X1              =   -74370
         X2              =   -66390
         Y1              =   5565
         Y2              =   5565
      End
      Begin VB.Label l 
         Caption         =   "First knowledge of danger of collision:"
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
         Index           =   83
         Left            =   -74370
         TabIndex        =   357
         Top             =   5670
         Width           =   3270
      End
      Begin VB.Label l 
         Caption         =   "Relative position of all vehicles at that time:"
         Height          =   225
         Index           =   84
         Left            =   -74370
         TabIndex        =   356
         Top             =   6720
         Width           =   3060
      End
      Begin VB.Label l 
         Caption         =   "What gave rise to notice of danger?"
         Height          =   225
         Index           =   85
         Left            =   -74370
         TabIndex        =   355
         Top             =   6195
         Width           =   2640
      End
      Begin VB.Label l 
         Caption         =   "Approximate time:"
         Height          =   225
         Index           =   86
         Left            =   -72585
         TabIndex        =   354
         Top             =   7140
         Width           =   1275
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   17
         X1              =   -69435
         X2              =   -69435
         Y1              =   5565
         Y2              =   7455
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   14
         X1              =   -71430
         X2              =   -66390
         Y1              =   1575
         Y2              =   1575
      End
      Begin VB.Label l 
         Caption         =   "Removal of Vehicles"
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
         Index           =   87
         Left            =   -71220
         TabIndex        =   353
         Top             =   1680
         Width           =   2220
      End
      Begin VB.Label l 
         Caption         =   "Conversations at Scene"
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
         Index           =   88
         Left            =   -69225
         TabIndex        =   352
         Top             =   5775
         Width           =   2220
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   13
         X1              =   -71430
         X2              =   -66390
         Y1              =   3675
         Y2              =   3675
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   12
         X1              =   -71430
         X2              =   -66390
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Label l 
         Caption         =   "Automobile Collision (3)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   92
         Left            =   -74265
         TabIndex        =   351
         Top             =   210
         Width           =   5265
      End
      Begin VB.Label l 
         Caption         =   "Photographs"
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
         Index           =   93
         Left            =   -67545
         TabIndex        =   350
         Top             =   2415
         Width           =   1065
      End
      Begin VB.Label l 
         Caption         =   "What parts of vehicle collided?"
         Height          =   225
         Index           =   96
         Left            =   -74265
         TabIndex        =   349
         Top             =   1260
         Width           =   2430
      End
      Begin VB.Label l 
         Caption         =   "Impact Description:"
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
         Index           =   97
         Left            =   -74370
         TabIndex        =   348
         Top             =   840
         Width           =   2220
      End
      Begin VB.Label l 
         Caption         =   "Where did each vehicle end up?"
         Height          =   225
         Index           =   98
         Left            =   -74265
         TabIndex        =   347
         Top             =   1680
         Width           =   2430
      End
      Begin VB.Label l 
         Caption         =   "How far apart were the vehicles?"
         Height          =   225
         Index           =   99
         Left            =   -74265
         TabIndex        =   346
         Top             =   2100
         Width           =   2430
      End
      Begin VB.Label Label7 
         Caption         =   "Describe any spinning or other action after impact:"
         Height          =   225
         Left            =   -74265
         TabIndex        =   345
         Top             =   2520
         Width           =   3795
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   8
         X1              =   -69540
         X2              =   -66390
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line li 
         BorderWidth     =   2
         Index           =   9
         X1              =   -69540
         X2              =   -69540
         Y1              =   4200
         Y2              =   6930
      End
      Begin VB.Label l 
         Caption         =   "Forces on body at impact:"
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
         Index           =   106
         Left            =   -69435
         TabIndex        =   344
         Top             =   4305
         Width           =   2325
      End
      Begin VB.Label l 
         Caption         =   "Description of Injuries,"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   94
         Left            =   -74265
         TabIndex        =   343
         Top             =   210
         Width           =   2745
      End
      Begin VB.Label l 
         Caption         =   "starting at the top of the head to the tip of the toes."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   95
         Left            =   -71535
         TabIndex        =   342
         Top             =   285
         Width           =   4425
      End
      Begin VB.Label l 
         Caption         =   "When first seen by health care provider:"
         Height          =   225
         Index           =   107
         Left            =   -74265
         TabIndex        =   341
         Top             =   1575
         Width           =   2850
      End
      Begin VB.Label l 
         Caption         =   "Several days later when all injuries could be appreciated:"
         Height          =   225
         Index           =   108
         Left            =   -74265
         TabIndex        =   340
         Top             =   1995
         Width           =   4215
      End
      Begin VB.Label l 
         Caption         =   "Medical Treatment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Index           =   109
         Left            =   -74265
         TabIndex        =   339
         Top             =   2835
         Width           =   2745
      End
      Begin VB.Label l 
         Caption         =   "Hospitals:"
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
         Index           =   110
         Left            =   -74475
         TabIndex        =   338
         Top             =   3360
         Width           =   2430
      End
      Begin VB.Label l 
         Caption         =   "Name:"
         Height          =   225
         Index           =   111
         Left            =   -74475
         TabIndex        =   337
         Top             =   3780
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "State:"
         Height          =   225
         Index           =   112
         Left            =   -74475
         TabIndex        =   336
         Top             =   5460
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "County:"
         Height          =   225
         Index           =   113
         Left            =   -74475
         TabIndex        =   335
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "City:"
         Height          =   225
         Index           =   114
         Left            =   -74475
         TabIndex        =   334
         Top             =   4620
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Street Name:"
         Height          =   225
         Index           =   115
         Left            =   -74475
         TabIndex        =   333
         Top             =   4200
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Date Begin:"
         Height          =   225
         Index           =   116
         Left            =   -74475
         TabIndex        =   332
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Label l 
         Caption         =   "Date End:"
         Height          =   225
         Index           =   117
         Left            =   -74475
         TabIndex        =   331
         Top             =   6300
         Width           =   1170
      End
      Begin VB.Label l 
         Caption         =   "Doctors:"
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
         Index           =   118
         Left            =   -71745
         TabIndex        =   330
         Top             =   3360
         Width           =   2430
      End
      Begin VB.Label l 
         Caption         =   "Name:"
         Height          =   225
         Index           =   119
         Left            =   -71745
         TabIndex        =   329
         Top             =   3780
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "State:"
         Height          =   225
         Index           =   120
         Left            =   -71745
         TabIndex        =   328
         Top             =   5460
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "County:"
         Height          =   225
         Index           =   121
         Left            =   -71745
         TabIndex        =   327
         Top             =   5040
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "City:"
         Height          =   225
         Index           =   122
         Left            =   -71745
         TabIndex        =   326
         Top             =   4620
         Width           =   855
      End
      Begin VB.Label l 
         Caption         =   "Street Name:"
         Height          =   225
         Index           =   123
         Left            =   -71745
         TabIndex        =   325
         Top             =   4200
         Width           =   960
      End
      Begin VB.Label l 
         Caption         =   "Date Begin:"
         Height          =   225
         Index           =   124
         Left            =   -71745
         TabIndex        =   324
         Top             =   5880
         Width           =   1170
      End
      Begin VB.Label l 
         Caption         =   "Date End:"
         Height          =   225
         Index           =   125
         Left            =   -71745
         TabIndex        =   323
         Top             =   6300
         Width           =   1170
      End
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   537
      ToolTipText     =   "All fields are Lock for edit operations, while you manage them from here."
      Top             =   8310
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10292
            Text            =   "No actions detected!"
            TextSave        =   "No actions detected!"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Edit"
            TextSave        =   "Edit"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Accept"
            TextSave        =   "Accept"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Cancel"
            TextSave        =   "Cancel"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2937
            Text            =   "www.MixofTix.net"
            TextSave        =   "www.MixofTix.net"
            Object.ToolTipText     =   "Visit our site for extra supports ..."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   13
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   5670
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   14
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   6090
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   15
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   6510
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   16
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   6930
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   17
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   7350
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   12
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   5250
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   7
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   3150
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   8
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   3570
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   9
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   3990
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   10
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   4410
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   11
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   4830
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   6
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   2730
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   1
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   630
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   2
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   1050
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   3
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   1470
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   4
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   1890
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   5
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   2310
      Width           =   750
   End
   Begin VB.Shape s 
      Height          =   435
      Index           =   0
      Left            =   8190
      Shape           =   2  'Oval
      Top             =   210
      Width           =   750
   End
End
Attribute VB_Name = "Form2"
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
Dim ADD_NEW_OK As String


Private Sub c_Click(Index As Integer)
Dim i As Integer
Dim fs As Object
Dim bb, a As String
If ADD_NEW_OK = "ERROR" Then
Exit Sub
End If
Select Case Index
  Case 152
    If c(152).Value = 1 Then
    t(187).Enabled = True
    Else
    t(187).Text = ""
    t(187).Enabled = False
    End If
  Case 151
    If c(151).Value = 1 Then
    t(186).Enabled = True
    Else
    t(186).Text = ""
    t(186).Enabled = False
    End If
  Case 150
    If c(150).Value = 1 Then
    t(185).Enabled = True
    Else
    t(185).Text = ""
    t(185).Enabled = False
    End If
  Case 149
    If c(149).Value = 1 Then
    t(184).Enabled = True
    Else
    t(184).Text = ""
    t(184).Enabled = False
    End If
  Case 148
    If c(148).Value = 1 Then
    t(183).Enabled = True
    Else
    t(183).Text = ""
    t(183).Enabled = False
    End If
  Case 147
    If c(147).Value = 1 Then
    t(182).Enabled = True
    Else
    t(182).Text = ""
    t(182).Enabled = False
    End If
  Case 146
    If c(146).Value = 1 Then
    dt(15).Enabled = True
    Else
    dt(15).Enabled = False
    End If
  Case 145
    If c(145).Value = 1 Then
    dt(14).Enabled = True
    Else
    dt(14).Enabled = False
    End If
  Case 144
    If c(144).Value = 1 Then
    dt(13).Enabled = True
    Else
    dt(13).Enabled = False
    End If
  Case 143
    If c(143).Value = 1 Then
    dt(12).Enabled = True
    Else
    dt(12).Enabled = False
    End If
  Case 142
    If c(142).Value = 1 Then
    dt(11).Enabled = True
    Else
    dt(11).Enabled = False
    End If
  Case 141
    If c(141).Value = 1 Then
    dt(10).Enabled = True
    Else
    dt(10).Enabled = False
    End If
  Case 140
    If c(140).Value = 1 Then
    t(172).Enabled = True
    Else
    t(172).Text = ""
    t(172).Enabled = False
    End If
  Case 139
    If c(139).Value = 1 Then
    t(171).Enabled = True
    Else
    t(171).Text = ""
    t(171).Enabled = False
    End If
  Case 138
    If c(138).Value = 1 Then
    t(170).Enabled = True
    Else
    t(170).Text = ""
    t(170).Enabled = False
    End If
  Case 137
    If c(137).Value = 1 Then
    t(169).Enabled = True
    Else
    t(169).Text = ""
    t(169).Enabled = False
    End If
  Case 136
    If c(136).Value = 1 Then
    t(168).Enabled = True
    Else
    t(168).Text = ""
    t(168).Enabled = False
    End If
  Case 135
    If c(135).Value = 1 Then
    t(167).Enabled = True
    Else
    t(167).Text = ""
    t(167).Enabled = False
    End If
  Case 134
    If c(134).Value = 1 Then
    t(166).Enabled = True
    Else
    t(166).Text = ""
    t(166).Enabled = False
    End If
  Case 133
    If c(133).Value = 1 Then
    t(165).Enabled = True
    Else
    t(165).Text = ""
    t(165).Enabled = False
    End If
  Case 132
    If c(132).Value = 1 Then
    t(164).Enabled = True
    Else
    t(164).Text = ""
    t(164).Enabled = False
    End If
  Case 108
    If c(108).Value = 1 Then
    For i = 98 To 99
    t(i).Enabled = True
    Next i
    Else
    For i = 98 To 99
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 109
    If c(102).Value = 1 Then
    For i = 6 To 7
    f(i).Enabled = True
    Next i
    Else
    For i = 6 To 7
    f(i).Enabled = False
    Next i
    End If
  Case 105
    If c(105).Value = 1 Then
    For i = 62 To 67
    t(i).Enabled = True
    Next i
    Else
    For i = 62 To 67
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 107
    If c(107).Value = 1 Then
    For i = 59 To 61
    t(i).Enabled = True
    Next i
    Else
    For i = 59 To 61
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 106
    If c(106).Value = 1 Then
    For i = 56 To 58
    t(i).Enabled = True
    Next i
    Else
    For i = 56 To 58
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 104
    If c(104).Value = 1 Then
    For i = 15 To 17
    t(i).Enabled = True
    Next i
    Else
    For i = 15 To 17
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 103
    If c(103).Value = 1 Then
    For i = 12 To 14
    t(i).Enabled = True
    Next i
    Else
    For i = 12 To 14
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 101
    If c(101).Value = 1 Then
       c(101).Caption = "Job Already Inactive!!"
    Else
       c(101).Caption = "Job Already Active!"
    End If
    If ADD_NEW_OK <> "" Then
    If c(101).Value = 1 Then
       c(101).Caption = "Job Already Inactive!!"
       Data2.Recordset.Edit
       Data2.Recordset("Inactivation_Date") = CStr(Now)
       Data2.Recordset.Update
       bb = App.Path & "\Users\" & Form1.Combo6.Text & "\Active Jobs\" & l(150).Caption & ".txt"
       Set fs = CreateObject("Scripting.FileSystemObject")
       a = App.Path & "\Users\" & Form1.Combo6.Text & "\Inactive Jobs\"
       fs.CopyFile bb, a
       'MsgBox bb
       Kill bb
    Else
       c(101).Caption = "Job Already Active!"
       Data2.Recordset.Edit
       Data2.Recordset("Inactivation_Date") = ""
       Data2.Recordset.Update
       bb = App.Path & "\Users\" & Form1.Combo6.Text & "\Inactive Jobs\" & l(150).Caption & ".txt"
       Set fs = CreateObject("Scripting.FileSystemObject")
       a = App.Path & "\Users\" & Form1.Combo6.Text & "\Active Jobs\"
       fs.CopyFile bb, a
       'MsgBox bb
       Kill bb
    End If
    End If
  Case 102
    If c(102).Value = 1 Then
    For i = 18 To 23
    t(i).Enabled = True
    Next i
    Else
    For i = 18 To 23
    t(i).Text = ""
    t(i).Enabled = False
    Next i
    End If
  Case 153
    If c(153).Value = 1 Then
       c(153).Caption = "Date Begin:"
       dt(16).Enabled = True
    Else
       c(153).Caption = "Date Start:"
       dt(16).Enabled = False
    End If
  Case 154
    If c(154).Value = 1 Then
       c(154).Caption = "Date End:"
       dt(17).Enabled = True
    Else
       c(154).Caption = "Today:"
       dt(17).Enabled = False
    End If

End Select
End Sub

Private Sub c_GotFocus(Index As Integer)
If ADD_NEW_OK = "ERROR" Then
If c(Index).Value = 0 Then
c(Index).Value = 0
'MsgBox "A"
Else
c(Index).Value = 1
'MsgBox "B"
End If
Command5.SetFocus
'MsgBox "C"
End If
End Sub

Private Sub co_Click(Index As Integer)
Dim nodx As Node
Dim k As String
Dim new_ As String
Dim bb As String
Dim a As String
Dim dou As Double
Dim n As Object
Dim fs As Object
Select Case Index
Case 0
cd.FileName = " "
cd.Filter = "Photoes Docs (*.gif)(*.jpeg)(*.bmp)|*.gif; *.jpg; *.bmp"
cd.ShowOpen
bb = cd.FileName
new_ = cd.FileTitle
If bb <> " " Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Photoes\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile bb, a
Set nodx = tv(0).Nodes.Add("Photoes", tvwChild, , new_)
End If
Case 1
cd.FileName = " "
cd.Filter = "Soundes Docs (*.wav)| *.wav"
cd.ShowOpen
bb = cd.FileName
new_ = cd.FileTitle
If bb <> " " Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Soundes\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile bb, a
Set nodx = tv(0).Nodes.Add("Soundes", tvwChild, , new_)
End If
Case 2
cd.FileName = " "
cd.Filter = "Movies Docs (*.avi)|*.avi"
cd.ShowOpen
bb = cd.FileName
new_ = cd.FileTitle
If bb <> " " Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Movies\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile bb, a
Set nodx = tv(0).Nodes.Add("Movies", tvwChild, , new_)
End If
Case 3
cd.FileName = " "
cd.Filter = "All files (*.*)|*.*"
cd.ShowOpen
bb = cd.FileName
new_ = cd.FileTitle
If bb <> " " Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Other\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile bb, a
Set nodx = tv(0).Nodes.Add("Other", tvwChild, , new_)
End If
Case 4
mm.Command = "Close"
Case 6
mm.PlayEnabled = True
mm.Command = "play"
Case 5
mm.Command = "pause"
Case 11
a = ""
If tv(0).SelectedItem.Parent = "Photoes" Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Photoes\"
a = a & tv(0).SelectedItem.Text
Kill a
tv(0).Nodes.Remove (tv(0).SelectedItem.Index)
End If
If tv(0).SelectedItem.Parent = "Soundes" Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Soundes\"
a = a & tv(0).SelectedItem.Text
Kill a
tv(0).Nodes.Remove (tv(0).SelectedItem.Index)
End If
If tv(0).SelectedItem.Parent = "Movies" Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Movies\"
a = a & tv(0).SelectedItem.Text
Kill a
tv(0).Nodes.Remove (tv(0).SelectedItem.Index)
End If
If tv(0).SelectedItem.Parent = "Other..." Then
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Other\"
a = a & tv(0).SelectedItem.Text
Kill a
tv(0).Nodes.Remove (tv(0).SelectedItem.Index)
End If
Case 8
Text4.Text = ""
Text4.SetFocus
If (Dir(App.Path & "\Jobs\" & l(150).Caption & "\Notes\Note.txt")) = "Note.txt" Then
Kill (App.Path & "\Jobs\" & l(150).Caption & "\Notes\Note.txt")
End If
Case 9
Text4.Text = Text4.Text & vbCrLf & "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\" & vbCrLf
Text4.SetFocus
Case 10
Dim i As Integer
Dim str, str1 As String
Dim intFileNum As Integer
intFileNum = FreeFile
Open App.Path & "\Jobs\" & l(150).Caption & "\Notes\Note.txt" For Append As #intFileNum
str = ""
str1 = ""
For i = 1 To Len(Text4.Text)
str1 = Mid(Text4.Text, i, 1)
If str1 = Chr(13) Or i = Len(Text4.Text) Then
   If i = Len(Text4.Text) Then
   Print #intFileNum, str & str1
   Else
   Print #intFileNum, str
   End If
str = ""
Else
    If str1 <> Chr(10) Then
    str = str & str1
    End If
End If
Next i
Close #intFileNum
End Select
End Sub





Private Sub Command1_Click()
Form7.Show
End Sub

Private Sub Command2_Click()
Form4.Show
End Sub

Private Sub Command3_Click()
Dim i As Integer
   For i = 0 To 13
     c(i).Value = False
   Next
End Sub

Private Sub Command4_Click()
Form4.Show
End Sub

Private Sub Command5_Click()
'MsgBox sb1.Panels(2).Text
If sb1.Panels(2).Text = "Status: Create New Job..." Then
Dim temp As String
'MsgBox "KK"
'Data2.Recordset.Update
Data3.Recordset.Update
Data4.Recordset.Update
Data5.Recordset.Update
Data6.Recordset.Update
Data7.Recordset.Update
Data8.Recordset.Update
Data9.Recordset.Update
Data10.Recordset.Update
sb1.Panels(2).Text = "Edit"
'MsgBox sb1.Panels(2).Text
sb1.Panels(2).AutoSize = sbrNoAutoSize
sb1.Panels(2).MinWidth = 1440
temp = sb1.Panels(3).Text
sb1.Panels(3).Text = "Accept"
'MsgBox sb1.Panels(3).Text
sb1.Panels(3).AutoSize = sbrNoAutoSize
sb1.Panels(3).MinWidth = 1440
sb1.Panels.Add
sb1.Panels(4).Text = "Cancel"
'MsgBox sb1.Panels(4).Text
sb1.Panels(4).AutoSize = sbrNoAutoSize
sb1.Panels(4).MinWidth = 1440
sb1.Panels.Add
sb1.Panels(5).Text = temp
'MsgBox sb1.Panels(5).Text
sb1.Panels(5).AutoSize = sbrContents
End If
'MsgBox "lll"
Unload Me
Form1.Show
End Sub


Private Sub dt_GotFocus(Index As Integer)
If Index = 18 Then Exit Sub
If ADD_NEW_OK = "ERROR" Then
Command5.SetFocus
SendKeys ("{Esc}")
End If
End Sub

Private Sub Form_Load()
Dim nodx As Node
Dim i As Integer
'Form1.sb.Panels(1).Text = "Came back from Job manager..."
On Error GoTo eff
Call note_pad_load
Call ShowFolderList55(App.Path & "\Mail Templates")
ADD_NEW_OK = "ERROR"
For i = 0 To 187
t(i).Locked = True
Next i
'For i = 101 To 152
'c(i).Enabled = False
'Next i
tv(0).LineStyle = tvwRootLines
Set nodx = tv(0).Nodes.Add(, , "r", "Related Documents")
Set nodx = tv(0).Nodes.Add("r", tvwChild, "Photoes", "Photoes")
Set nodx = tv(0).Nodes.Add("r", tvwChild, "Soundes", "Soundes")
Set nodx = tv(0).Nodes.Add("r", tvwChild, "Movies", "Movies")
Set nodx = tv(0).Nodes.Add("r", tvwChild, "Other", "Other...")
l(150).Caption = Form1.Label25.Caption
'If Form1.Command3.Caption = "&New Job" Then
'l(150).Caption = Form1.Label27.Caption
'End If
Label5.Caption = Form1.Label21.Caption
ShowFolderList_Photoes (App.Path & "\Jobs\" & l(150).Caption & "\Documents\Photoes")
ShowFolderList_Soundes (App.Path & "\Jobs\" & l(150).Caption & "\Documents\Soundes")
ShowFolderList_Movies (App.Path & "\Jobs\" & l(150).Caption & "\Documents\Movies")
ShowFolderList_Other (App.Path & "\Jobs\" & l(150).Caption & "\Documents\Other")
mm.Command = "open"
Call tv1_initialize
  '******************
  '******************
Call ref(l(150).Caption)
  '******************
Call ref111(l(150).Caption)
  '******************
Call ref2(l(150).Caption)
If Form1.Command3.Caption = "&New Job" Then
sb1.Panels(2).Text = "Status: Create New Job..."
sb1.Panels(2).AutoSize = sbrContents
sb1.Panels.Remove (3)
sb1.Panels.Remove (3)
'**********************
Data3.Recordset.AddNew
Data3.Recordset("Job_ID") = l(150).Caption
Data3.Recordset.Update
'Data3.Recordset.FindFirst l(150).Caption
Data4.Recordset.AddNew
Data4.Recordset("Job_ID") = l(150).Caption
Data4.Recordset.Update
'Data4.Recordset.FindFirst l(150).Caption
Data5.Recordset.AddNew
Data5.Recordset("Job_ID") = l(150).Caption
Data5.Recordset.Update
'Data5.Recordset.FindFirst l(150).Caption
Data6.Recordset.AddNew
Data6.Recordset("Job_ID") = l(150).Caption
Data6.Recordset.Update
'Data6.Recordset.FindFirst l(150).Caption
Data7.Recordset.AddNew
Data7.Recordset("Job_ID") = l(150).Caption
Data7.Recordset.Update
'Data7.Recordset.FindFirst l(150).Caption
Data8.Recordset.AddNew
Data8.Recordset("Job_ID") = l(150).Caption
Data8.Recordset.Update
'Data8.Recordset.FindFirst l(150).Caption
Data9.Recordset.AddNew
Data9.Recordset("Job_ID") = l(150).Caption
Data9.Recordset.Update
'Data9.Recordset.FindFirst l(150).Caption
Data10.Recordset.AddNew
Data10.Recordset("Job_ID") = l(150).Caption
Data10.Recordset.Update
'Data10.Recordset.FindFirst l(150).Caption
'*******************
ADD_NEW_OK = "OK"
For i = 0 To 187
t(i).Locked = False
Next i
'For i = 101 To 152
'c(i).Enabled = True
'Next i
Call ref111(l(150).Caption)
  '******************
Call ref2(l(150).Caption)
'Data2.Recordset.MoveLast  'Allalll
Data2.Recordset.Edit
Data3.Recordset.Edit
Data4.Recordset.Edit
Data5.Recordset.Edit
Data6.Recordset.Edit
Data7.Recordset.Edit
Data8.Recordset.Edit
Data9.Recordset.Edit
Data10.Recordset.Edit
End If
ADD_NEW_OK = ""
For i = 101 To 108
Call c_Click(i)
Next i
For i = 132 To 154
Call c_Click(i)
Next i
ADD_NEW_OK = "ERROR"
If Form1.Command3.Caption = "&New Job" Then
ADD_NEW_OK = "OK"
End If
Exit Sub
eff:
MsgBox (err.Description)
MsgBox (err.Number)
'*****
End Sub

Private Sub Form_Unload(Cancel As Integer)
'***************
Dim intFileNum As Integer
Dim Date_, Time_ As String
If ADD_NEW_OK = "OK" Then
i = MsgBox("We suppose that you already accepted the New record!", vbInformation, "Attorney Master O.K.")
ADD_NEW_OK = "ERROR"
'Data2.Recordset.Update
'Data3.Recordset.Update
'Data4.Recordset.Update
'Data5.Recordset.Update
'Data6.Recordset.Update
'Data7.Recordset.Update
'Data8.Recordset.Update
'Data9.Recordset.Update
'Data10.Recordset.Update
'Exit Sub
End If
intFileNum = FreeFile
If c(101).Value = 1 Then
Open App.Path & "\Users\" & Form1.Combo6.Text & "\Inactive Jobs\" & Form1.Label25.Caption & ".txt" For Append As #intFileNum
Else
Open App.Path & "\Users\" & Form1.Combo6.Text & "\Active Jobs\" & Form1.Label25.Caption & ".txt" For Append As #intFileNum
End If
Time_ = Time
Date_ = Date
Print #intFileNum, "Job closed at:"; Spc(3); Time_; Spc(3); Date_; Spc(3); "By_"; Form1.Combo6.Text
Close #intFileNum
'***************
mm.Command = "Close"
'***************
'Data1.Recordset.Close
'Data1.Database.Close
'Data2.Recordset.Close
'Data2.Database.Close
'***************
Call Form1.total
For i = 0 To 187
t(i).Locked = True
Next i
Form1.Show
Unload Form4
End Sub


Private Sub Label13_Click()
If Label13.Caption = "AM" Then
Label13.Caption = "PM"
Else
Label13.Caption = "AM"
End If
End Sub

Private Sub Label40_Click()
If Label40.Caption = "AM" Then
Label40.Caption = "PM"
Else
Label40.Caption = "AM"
End If
End Sub

Private Sub sb1_PanelClick(ByVal Panel As MSComctlLib.Panel)
'***************
Dim i As Integer
Dim intFileNum As Integer
Dim Date_, Time_ As String
intFileNum = FreeFile
Open App.Path & "\Users\" & Form1.Combo6.Text & "\Modified\" & Form1.Label25.Caption & "ML.txt" For Append As #intFileNum
Time_ = Time
Date_ = Date
If Panel.Text = "Status: Create New Job..." Then
If ADD_NEW_OK = "OK" Then
i = MsgBox("You already accepted the New record.", vbInformation, "Attorney Master O.K.")
ADD_NEW_OK = "ERROR"
Data2.Recordset.Update
Data3.Recordset.Update
Data4.Recordset.Update
Data5.Recordset.Update
Data6.Recordset.Update
Data7.Recordset.Update
Data8.Recordset.Update
Data9.Recordset.Update
Data10.Recordset.Update
Exit Sub
End If
End If
'***************
If Panel.Text = "Edit" Then
'Data2.Recordset.MoveLast  'Allalll
'Data2.Recordset.Edit
Data3.Recordset.Edit
Data4.Recordset.Edit
Data5.Recordset.Edit
Data6.Recordset.Edit
Data7.Recordset.Edit
Data8.Recordset.Edit
Data9.Recordset.Edit
Data10.Recordset.Edit
'********************
Print #intFileNum, "Job Edit Started at:"; Spc(4); Time_; Spc(3); Date_; Spc(3); "By_"; Form1.Combo6.Text
'*******************
Close #intFileNum
'*******************
For i = 0 To 187
t(i).Locked = False
Next i
'For i = 101 To 152
'c(i).Enabled = True
'Next i
ADD_NEW_OK = "OK"
MsgBox ("Edit Mode")
End If
If Panel.Text = "Accept" Then
If ADD_NEW_OK = "OK" Then
i = MsgBox("You already accepted the record.", vbInformation, "Attorney Master O.K.")
ADD_NEW_OK = "ERROR"
'Data2.Recordset.Update
Data3.Recordset.Update
Data4.Recordset.Update
Data5.Recordset.Update
Data6.Recordset.Update
Data7.Recordset.Update
Data8.Recordset.Update
Data9.Recordset.Update
Data10.Recordset.Update
Print #intFileNum, "Job Edit Accepted at:"; Spc(3); Time_; Spc(3); Date_; Spc(3); "By_"; Form1.Combo6.Text
'*******************
Close #intFileNum
'*******************
For i = 0 To 187
t(i).Locked = True
Next i
'For i = 101 To 152
'c(i).Enabled = False
'Next i
End If
End If
If Panel.Text = "Cancel" Then
If ADD_NEW_OK = "OK" Then
i = MsgBox("You already canceled the record.", vbInformation, "Attorney Master CANCEL")
ADD_NEW_OK = "ERROR"
'Data2.Recordset.CancelUpdate
Data3.Recordset.CancelUpdate
Data4.Recordset.CancelUpdate
Data5.Recordset.CancelUpdate
Data6.Recordset.CancelUpdate
Data7.Recordset.CancelUpdate
Data8.Recordset.CancelUpdate
Data9.Recordset.CancelUpdate
Data10.Recordset.CancelUpdate
'*************************
Print #intFileNum, "Job Edit Canceled at:"; Spc(3); Time_; Spc(3); Date_; Spc(3); "By_"; Form1.Combo6.Text
'*******************
Close #intFileNum
'*******************
For i = 0 To 187
t(i).Locked = True
Next i
'For i = 101 To 152
'c(i).Enabled = False
'Next i
End If
End If
End Sub


Private Sub st1_Click11(PreviousTab As Integer)
'Private Sub st1_Click(PreviousTab As Integer)
'*****
Select Case st1.Caption
  Case "P. 1"
'      Call ref1(l(150).Caption, "Interview")
  Case "P. 2"
'       Call ref1(l(150).Caption, "Defendant")
  Case "P. 3"
'      Call ref1(l(150).Caption, "Facts")
  Case "P. 4"
'      Call ref1(l(150).Caption, "Automobile_C1")
  Case "P. 5"
'      Call ref1(l(150).Caption, "Automobile_C2")
  Case "P. 6"
'      Call ref1(l(150).Caption, "Automobile_C3")
  Case "P. 7"
'      Call ref1(l(150).Caption, "D_Injuries")
  Case "P. 8"
'      Call ref1(l(150).Caption, "Wage_Loss")
End Select
End Sub


Private Sub st1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case st1.Caption
  Case "P. 1"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Interview..."
  Case "P. 2"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Defendant..."
  Case "P. 3"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Facts Relating to the Injury or Occurrence"
  Case "P. 4"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Automobile Collision (1) "
  Case "P. 5"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Automobile Collision (2) "
  Case "P. 6"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Automobile Collision (3) "
  Case "P. 7"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Description of Injuries,... & Medical Treatment"
  Case "P. 8"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Wage Loss"
  Case "P. 9"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Expense Account"
  Case "P. 10"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Relative Comments"
  Case "P. 11"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   Relative Documments,..."
  Case "P. 12"
     sb1.Panels.Item(1) = "(" & st1.Caption & ")   File Documents Diagrams"
End Select
End Sub


Private Sub Text5_Change()
'dt(14).Value = Text5.Text
End Sub

Private Sub t_Change(Index As Integer)
Dim i As Integer
If ADD_NEW_OK = "ERROR" Then Exit Sub
'***************
Select Case Index
'*************** Interview **************
Case 0
If Len(t(0).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 1
If Len(t(1).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 2
If Len(t(2).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 3
If Len(t(3).Text) = 3 Then
t(4).SetFocus
End If
If Len(t(3).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 4
If Len(t(4).Text) = 2 Then
t(5).SetFocus
End If
If Len(t(4).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 5
If Len(t(5).Text) = 4 Then
t(6).SetFocus
End If
If Len(t(5).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 6
If Len(t(6).Text) > 8 Then
i = MsgBox("Available digits are 8 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 7
If Len(t(7).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 8
If Len(t(8).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 9
If Len(t(9).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 10
If Len(t(10).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 11
If Len(t(11).Text) > 5 Then
i = MsgBox("Available digits are 5 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 12
If Len(t(12).Text) = 4 Then
c(104).SetFocus
End If
If Len(t(12).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 13
If Len(t(13).Text) = 3 Then
t(12).SetFocus
End If
If Len(t(13).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 14
If Len(t(14).Text) = 3 Then
t(13).SetFocus
End If
If Len(t(14).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 15
If Len(t(15).Text) = 4 Then
dt(1).SetFocus
End If
If Len(t(15).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 16
If Len(t(16).Text) = 3 Then
t(15).SetFocus
End If
If Len(t(16).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 17
If Len(t(17).Text) = 3 Then
t(16).SetFocus
End If
If Len(t(17).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 18
If Len(t(18).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 19
If Len(t(19).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 20
If Len(t(20).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 21
If Len(t(21).Text) = 3 Then
t(22).SetFocus
End If
If Len(t(21).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 22
If Len(t(22).Text) = 3 Then
t(23).SetFocus
End If
If Len(t(22).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 23
If Len(t(23).Text) = 4 Then
t(24).SetFocus
End If
If Len(t(23).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 24
If Len(t(24).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 25
If Len(t(25).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 26
If Len(t(26).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 27
If Len(t(27).Text) > 60 Then
i = MsgBox("Available digits are 60 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 28
If Len(t(28).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 29
If Len(t(29).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 30
If Len(t(30).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 31
If Len(t(31).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 32
If Len(t(32).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 33
If Len(t(33).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 34
If Len(t(34).Text) = 4 Then
t(37).SetFocus
End If
If Len(t(34).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 35
If Len(t(35).Text) = 3 Then
t(34).SetFocus
End If
If Len(t(35).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 36
If Len(t(36).Text) = 3 Then
t(35).SetFocus
End If
If Len(t(36).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 37
If Len(t(37).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 38
If Len(t(38).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 39
If Len(t(39).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 40
If Len(t(40).Text) = 4 Then
t(43).SetFocus
End If
If Len(t(40).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 41
If Len(t(41).Text) = 3 Then
t(40).SetFocus
End If
If Len(t(41).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 42
If Len(t(42).Text) = 3 Then
t(41).SetFocus
End If
If Len(t(42).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 43
If Len(t(43).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Defendant **************
Case 44
If Len(t(44).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 45
If Len(t(45).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 46
If Len(t(46).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 47
If Len(t(47).Text) = 3 Then
t(48).SetFocus
End If
If Len(t(47).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 48
If Len(t(48).Text) = 2 Then
t(49).SetFocus
End If
If Len(t(48).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 49
If Len(t(49).Text) = 4 Then
t(50).SetFocus
End If
If Len(t(49).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 50
If Len(t(50).Text) > 6 Then
i = MsgBox("Available digits are 6 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 51
If Len(t(51).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 52
If Len(t(52).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 53
If Len(t(53).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 54
If Len(t(54).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 55
If Len(t(55).Text) > 5 Then
i = MsgBox("Available digits are 5 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 58
If Len(t(58).Text) = 4 Then
c(107).SetFocus
End If
If Len(t(58).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 57
If Len(t(57).Text) = 3 Then
t(58).SetFocus
End If
If Len(t(57).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 56
If Len(t(56).Text) = 3 Then
t(57).SetFocus
End If
If Len(t(56).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 61
If Len(t(61).Text) = 4 Then
dt(2).SetFocus
End If
If Len(t(61).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 60
If Len(t(60).Text) = 3 Then
t(61).SetFocus
End If
If Len(t(60).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 59
If Len(t(59).Text) = 3 Then
t(60).SetFocus
End If
If Len(t(59).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 62
If Len(t(62).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 63
If Len(t(63).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 64
If Len(t(64).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 65
If Len(t(65).Text) = 3 Then
t(66).SetFocus
End If
If Len(t(65).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 66
If Len(t(66).Text) = 3 Then
t(67).SetFocus
End If
If Len(t(66).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 67
If Len(t(67).Text) = 4 Then
t(68).SetFocus
End If
If Len(t(67).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 68
If Len(t(68).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 69
If Len(t(69).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 70
If Len(t(70).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 71
If Len(t(71).Text) > 60 Then
i = MsgBox("Available digits are 60 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 72
If Len(t(72).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 73
If Len(t(73).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 74
If Len(t(74).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 75
If Len(t(75).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 76
If Len(t(76).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 77
If Len(t(77).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 80
If Len(t(80).Text) = 4 Then
t(81).SetFocus
End If
If Len(t(80).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 79
If Len(t(79).Text) = 3 Then
t(80).SetFocus
End If
If Len(t(79).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 78
If Len(t(78).Text) = 3 Then
t(79).SetFocus
End If
If Len(t(78).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 81
If Len(t(81).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 82
If Len(t(82).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 83
If Len(t(83).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 86
If Len(t(86).Text) = 4 Then
t(87).SetFocus
End If
If Len(t(86).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 85
If Len(t(85).Text) = 3 Then
t(86).SetFocus
End If
If Len(t(85).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 84
If Len(t(84).Text) = 3 Then
t(85).SetFocus
End If
If Len(t(84).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 87
If Len(t(87).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Facts **************
Case 88
If Len(t(88).Text) = 2 Then
t(89).SetFocus
End If
If Len(t(88).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 89
If Len(t(89).Text) = 2 Then
t(90).SetFocus
End If
If Len(t(89).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 90
If Len(t(90).Text) = 2 Then
t(91).SetFocus
End If
If Len(t(90).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 91
If Len(t(91).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 92
If Len(t(92).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 93
If Len(t(93).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 94
If Len(t(94).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 95
If Len(t(95).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 96
If Len(t(96).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 97
If Len(t(97).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 98
If Len(t(98).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 99
If Len(t(99).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 100
If Len(t(100).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 101
If Len(t(101).Text) > 200 Then
i = MsgBox("Available digits are 200 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 102
If Len(t(102).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 103
If Len(t(103).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 104
If Len(t(104).Text) > 200 Then
i = MsgBox("Available digits are 200 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Automobile_C1 **************
Case 105
If Len(t(105).Text) > 150 Then
i = MsgBox("Available digits are 150 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 106
If Len(t(106).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 107
If Len(t(107).Text) > 2 Then
i = MsgBox("Available digits are 2 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 108
If Len(t(108).Text) > 7 Then
i = MsgBox("Available digits are 7 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 109
If Len(t(109).Text) > 40 Then
i = MsgBox("Available digits are 40 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 110
If Len(t(110).Text) > 150 Then
i = MsgBox("Available digits are 150 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 111
If Len(t(111).Text) > 150 Then
i = MsgBox("Available digits are 150 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 112
If Len(t(112).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 113
If Len(t(113).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 114
If Len(t(114).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 115
If Len(t(115).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 116
If Len(t(116).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 117
If Len(t(117).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Automobile_C2 **************
Case 118
If Len(t(118).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 119
If Len(t(119).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 120
If Len(t(120).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 121
If Len(t(121).Text) > 40 Then
i = MsgBox("Available digits are 40 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 122
If Len(t(122).Text) = 3 Then
t(123).SetFocus
End If
If Len(t(122).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 123
If Len(t(123).Text) = 3 Then
t(123).SetFocus
End If
If Len(t(123).Text) > 3 Then
i = MsgBox("Available digits are 3 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 124
If Len(t(124).Text) = 4 Then
t(125).SetFocus
End If
If Len(t(124).Text) > 4 Then
i = MsgBox("Available digits are 4 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 125
If Len(t(125).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 126
If Len(t(126).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 127
If Len(t(127).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 128
If Len(t(128).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 129
If Len(t(129).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 130
If Len(t(130).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 131
If Len(t(131).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 132
If Len(t(132).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 133
If Len(t(133).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Automobile_C3 **************
Case 134
If Len(t(134).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 135
If Len(t(135).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 136
If Len(t(136).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 137
If Len(t(137).Text) > 200 Then
i = MsgBox("Available digits are 200 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 138
If Len(t(138).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 139
If Len(t(139).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 140
If Len(t(140).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 141
If Len(t(141).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 142
If Len(t(142).Text) > 200 Then
i = MsgBox("Available digits are 200 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** D_Injuries **************
Case 144
If Len(t(144).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 145
If Len(t(145).Text) > 20 Then
i = MsgBox("Available digits are 20 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 146
If Len(t(146).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 147
If Len(t(147).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 148
If Len(t(148).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 149
If Len(t(149).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 150
If Len(t(150).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 151
If Len(t(151).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 152
If Len(t(152).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 153
If Len(t(153).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 154
If Len(t(154).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 155
If Len(t(155).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 156
If Len(t(156).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 157
If Len(t(157).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 158
If Len(t(158).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 159
If Len(t(159).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 160
If Len(t(160).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Wage_Loss **************
Case 161
If Len(t(161).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 162
If Len(t(162).Text) > 30 Then
i = MsgBox("Available digits are 30 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 163
If Len(t(163).Text) > 100 Then
i = MsgBox("Available digits are 100 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 164
If Len(t(164).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 165
If Len(t(165).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 166
If Len(t(166).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 167
If Len(t(167).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 168
If Len(t(168).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 169
If Len(t(169).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 170
If Len(t(170).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 171
If Len(t(171).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 172
If Len(t(172).Text) > 15 Then
i = MsgBox("Available digits are 15 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 173
If Len(t(173).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 174
If Len(t(174).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 175
If Len(t(175).Text) > 50 Then
i = MsgBox("Available digits are 50 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
'*************** Expenses Account **************
Case 176
If Len(t(176).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 177
If Len(t(177).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 178
If Len(t(178).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 179
If Len(t(179).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 180
If Len(t(180).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 181
If Len(t(181).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 182
If Len(t(182).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 183
If Len(t(183).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 184
If Len(t(184).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 185
If Len(t(185).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 186
If Len(t(186).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
Case 187
If Len(t(187).Text) > 10 Then
i = MsgBox("Available digits are 10 for this field!", vbCritical, "Attorney Master Alert")
SendKeys ("{Backspace}")
SendKeys ("{Home}")
End If
End Select
End Sub

Private Sub tv_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
Dim a As String
mm.Command = "Close"
For i = 4 To 6
co(i).Enabled = False
Next i
Select Case Index
Case 0
If tv(0).SelectedItem.Text <> "Related Documents" Then
If tv(0).SelectedItem.Parent = "Photoes" Then
pi.Visible = False
p.Visible = True
'a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Photoes\"
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Photoes\"
a = a & tv(0).SelectedItem.Text
Set p.Picture = LoadPicture(a)
End If
If tv(0).SelectedItem.Parent = "Soundes" Then
pi.Visible = False
p.Visible = True
'a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Soundes\"
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Soundes\"
a = a & tv(0).SelectedItem.Text
For i = 4 To 6
co(i).Enabled = True
Next i
mm.DeviceType = "WaveAudio"
mm.FileName = a
mm.Command = "open"
End If
If tv(0).SelectedItem.Parent = "Movies" Then
p.Visible = False
pi.Visible = True
'a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Movies\"
a = App.Path & "\Jobs\" & l(150).Caption & "\Documents\Movies\"
a = a & tv(0).SelectedItem.Text
For i = 4 To 6
co(i).Enabled = True
Next i
mm.DeviceType = "AVIVideo"
mm.FileName = a
mm.Command = "open"
mm.hWndDisplay = pi.hWnd
End If
End If
End Select
End Sub
Sub ShowFolderList_1(folderspec)
 Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 In sf
        s = f1.Name
    
    'MsgBox (f1.DateCreated)
    'MsgBox (f1.DateLastModified)
    'MsgBox (f1.DateLastAccessed)
    Next
End Sub

Sub ShowFolderList_Photoes(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
Set nodx = tv(0).Nodes.Add("Photoes", tvwChild, , s)
    Next
End Sub
Sub ShowFolderList_Soundes(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
Set nodx = tv(0).Nodes.Add("Soundes", tvwChild, , s)
    Next
End Sub
Sub ShowFolderList_Movies(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
Set nodx = tv(0).Nodes.Add("Movies", tvwChild, , s)
    Next
End Sub
Sub ShowFolderList_Other(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
Set nodx = tv(0).Nodes.Add("Other", tvwChild, , s)
    Next
End Sub




Private Sub tv1_initialize()
Dim nodx As Node
tv(1).LineStyle = tvwRootLines
Set nodx = tv(1).Nodes.Add(, , "r", "INTERVIEW...")
Set nodx = tv(1).Nodes.Add("r", tvwChild, , "Last Name")
Set nodx = tv(1).Nodes.Add("r", tvwChild, , "First Name")
Set nodx = tv(1).Nodes.Add("r", tvwChild, , "MI")
Set nodx = tv(1).Nodes.Add("r", tvwChild, , "SS#")
Set nodx = tv(1).Nodes.Add("r", tvwChild, "ADD1", "Address")
Set nodx = tv(1).Nodes.Add("ADD1", tvwChild, , "Number")
Set nodx = tv(1).Nodes.Add("ADD1", tvwChild, , "Street name")
Set nodx = tv(1).Nodes.Add("ADD1", tvwChild, , "City")
Set nodx = tv(1).Nodes.Add("ADD1", tvwChild, , "County")
Set nodx = tv(1).Nodes.Add("ADD1", tvwChild, , "State")
Set nodx = tv(1).Nodes.Add("ADD1", tvwChild, , "Zip Code")
Set nodx = tv(1).Nodes.Add("r", tvwChild, "HB1", "Phone Number")
Set nodx = tv(1).Nodes.Add("HB1", tvwChild, , "Home")
Set nodx = tv(1).Nodes.Add("HB1", tvwChild, , "Business")
Set nodx = tv(1).Nodes.Add("r", tvwChild, "S1", "Spouse")
Set nodx = tv(1).Nodes.Add("S1", tvwChild, , "Last Name")
Set nodx = tv(1).Nodes.Add("S1", tvwChild, , "First Name")
Set nodx = tv(1).Nodes.Add("S1", tvwChild, , "Occupation")
Set nodx = tv(1).Nodes.Add("S1", tvwChild, "HB2", "Phone Number")
Set nodx = tv(1).Nodes.Add("HB2", tvwChild, , "Home")
Set nodx = tv(1).Nodes.Add("HB2", tvwChild, , "Business")
Set nodx = tv(1).Nodes.Add("r", tvwChild, , "Age")
Set nodx = tv(1).Nodes.Add("r", tvwChild, "OC1", "Occupation")
Set nodx = tv(1).Nodes.Add("OC1", tvwChild, , "By Whome Employed")
Set nodx = tv(1).Nodes.Add("OC1", tvwChild, , "Nature of Work")
Set nodx = tv(1).Nodes.Add("OC1", tvwChild, , "How Long Employed")
Set nodx = tv(1).Nodes.Add("OC1", tvwChild, , "Prior Employment History")
Set nodx = tv(1).Nodes.Add("r", tvwChild, "ICN1", "Insurance Companies Name")
Set nodx = tv(1).Nodes.Add("ICN1", tvwChild, , "Liability")
Set nodx = tv(1).Nodes.Add("ICN1", tvwChild, , "PIP,if applicable")
Set nodx = tv(1).Nodes.Add("ICN1", tvwChild, , "Medical")
Set nodx = tv(1).Nodes.Add("ICN1", tvwChild, , "Limits")
Set nodx = tv(1).Nodes.Add("ICN1", tvwChild, "BI1", "BI Adjusters")
Set nodx = tv(1).Nodes.Add("BI1", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("BI1", tvwChild, , "Phone number")
Set nodx = tv(1).Nodes.Add("BI1", tvwChild, , "Claim#")
Set nodx = tv(1).Nodes.Add("ICN1", tvwChild, "PD1", "PD Adjusters")
Set nodx = tv(1).Nodes.Add("PD1", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("PD1", tvwChild, , "Phone number")
Set nodx = tv(1).Nodes.Add("PD1", tvwChild, , "Claim#")
Set nodx = tv(1).Nodes.Add(, , "r1", "DEFENDANT...")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, , "Last Name")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, , "First Name")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, , "MI")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, , "SS#")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, "ADD2", "Address")
Set nodx = tv(1).Nodes.Add("ADD2", tvwChild, , "Number")
Set nodx = tv(1).Nodes.Add("ADD2", tvwChild, , "Street name")
Set nodx = tv(1).Nodes.Add("ADD2", tvwChild, , "City")
Set nodx = tv(1).Nodes.Add("ADD2", tvwChild, , "County")
Set nodx = tv(1).Nodes.Add("ADD2", tvwChild, , "State")
Set nodx = tv(1).Nodes.Add("ADD2", tvwChild, , "Zip Code")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, "HB3", "Phone Number")
Set nodx = tv(1).Nodes.Add("HB3", tvwChild, , "Home")
Set nodx = tv(1).Nodes.Add("HB3", tvwChild, , "Business")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, "S2", "Spouse")
Set nodx = tv(1).Nodes.Add("S2", tvwChild, , "Last Name")
Set nodx = tv(1).Nodes.Add("S2", tvwChild, , "First Name")
Set nodx = tv(1).Nodes.Add("S2", tvwChild, , "Occupation")
Set nodx = tv(1).Nodes.Add("S2", tvwChild, "HB4", "Phone Number")
Set nodx = tv(1).Nodes.Add("HB4", tvwChild, , "Home")
Set nodx = tv(1).Nodes.Add("HB4", tvwChild, , "Business")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, , "Age")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, "OC2", "Occupation")
Set nodx = tv(1).Nodes.Add("OC2", tvwChild, , "By Whome Employed")
Set nodx = tv(1).Nodes.Add("OC2", tvwChild, , "Nature of Work")
Set nodx = tv(1).Nodes.Add("OC2", tvwChild, , "How Long Employed")
Set nodx = tv(1).Nodes.Add("OC2", tvwChild, , "Prior Employment History")
Set nodx = tv(1).Nodes.Add("r1", tvwChild, "ICN2", "Insurance Companies Name")
Set nodx = tv(1).Nodes.Add("ICN2", tvwChild, , "Liability")
Set nodx = tv(1).Nodes.Add("ICN2", tvwChild, , "PIP,if applicable")
Set nodx = tv(1).Nodes.Add("ICN2", tvwChild, , "Medical")
Set nodx = tv(1).Nodes.Add("ICN2", tvwChild, , "Limits")
Set nodx = tv(1).Nodes.Add("ICN2", tvwChild, "BI2", "BI Adjusters")
Set nodx = tv(1).Nodes.Add("BI2", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("BI2", tvwChild, , "Phone number")
Set nodx = tv(1).Nodes.Add("BI2", tvwChild, , "Claim#")
Set nodx = tv(1).Nodes.Add("ICN2", tvwChild, "PD2", "PD Adjusters")
Set nodx = tv(1).Nodes.Add("PD2", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("PD2", tvwChild, , "Phone number")
Set nodx = tv(1).Nodes.Add("PD2", tvwChild, , "Claim#")
Set nodx = tv(1).Nodes.Add(, , "r3", "FACTS RELATING TO THE INJURY OR OCCURENCE...")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "W1", "When")
Set nodx = tv(1).Nodes.Add("W1", tvwChild, , "Date")
Set nodx = tv(1).Nodes.Add("W1", tvwChild, , "Time")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "LO1", "Location")
Set nodx = tv(1).Nodes.Add("LO1", tvwChild, , "City")
Set nodx = tv(1).Nodes.Add("LO1", tvwChild, , "County")
Set nodx = tv(1).Nodes.Add("LO1", tvwChild, , "State")
Set nodx = tv(1).Nodes.Add("LO1", tvwChild, "LO1ADD", "Address")
Set nodx = tv(1).Nodes.Add("LO1ADD", tvwChild, , "Street")
Set nodx = tv(1).Nodes.Add("LO1ADD", tvwChild, , "Intersecting Street")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "DI1", "Description of Injury or Occurrence")
Set nodx = tv(1).Nodes.Add("DI1", tvwChild, , "Narrative Description")
Set nodx = tv(1).Nodes.Add("DI1", tvwChild, , "Diagram of scene")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "IIR1", "Investigation of Incident Reports")
Set nodx = tv(1).Nodes.Add("IIR1", tvwChild, "IIR1I", "Investigated?")
Set nodx = tv(1).Nodes.Add("IIR1I", tvwChild, , "Names of Investigating Agency")
Set nodx = tv(1).Nodes.Add("IIR1I", tvwChild, , "Name of Investigating Officer")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "STS1", "Statements?")
Set nodx = tv(1).Nodes.Add("STS1", tvwChild, "STS1MP", "Made by the Plaintiff")
Set nodx = tv(1).Nodes.Add("STS1MP", tvwChild, , "Oral")
Set nodx = tv(1).Nodes.Add("STS1MP", tvwChild, , "Written")
Set nodx = tv(1).Nodes.Add("STS1", tvwChild, "STS1MD", "Made by the Defendant")
Set nodx = tv(1).Nodes.Add("STS1MD", tvwChild, , "Oral")
Set nodx = tv(1).Nodes.Add("STS1MD", tvwChild, , "Written")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "WIT1", "Witnesses")
Set nodx = tv(1).Nodes.Add("WIT1", tvwChild, , "Names")
Set nodx = tv(1).Nodes.Add("WIT1", tvwChild, , "Addresses")
Set nodx = tv(1).Nodes.Add("WIT1", tvwChild, , "Occupations")
Set nodx = tv(1).Nodes.Add("WIT1", tvwChild, , "Phone Numbers")
Set nodx = tv(1).Nodes.Add("WIT1", tvwChild, , "Subject Matter of Information")
Set nodx = tv(1).Nodes.Add("r3", tvwChild, "PHOT1", "Photographs")
Set nodx = tv(1).Nodes.Add("PHOT1", tvwChild, "PHOT1HT", "Have been taken")
Set nodx = tv(1).Nodes.Add("PHOT1HT", tvwChild, , "By whom")
Set nodx = tv(1).Nodes.Add("PHOT1HT", tvwChild, , "Of accident or incident scene")
Set nodx = tv(1).Nodes.Add("PHOT1HT", tvwChild, , "Of damage")
Set nodx = tv(1).Nodes.Add("PHOT1HT", tvwChild, , "Other")
Set nodx = tv(1).Nodes.Add("PHOT1", tvwChild, "PHOT1TST", "Photographs that should be taken")
Set nodx = tv(1).Nodes.Add("PHOT1TST", tvwChild, , "Plaintiff's Vehicle")
Set nodx = tv(1).Nodes.Add("PHOT1TST", tvwChild, , "Defendant's Vehicle")
Set nodx = tv(1).Nodes.Add("PHOT1TST", tvwChild, , "Scene")
Set nodx = tv(1).Nodes.Add("PHOT1TST", tvwChild, , "Other")
Set nodx = tv(1).Nodes.Add(, , "r4", "AUTOMOBILE COLLISION...")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, , "Nature of Weather")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, , "Visibility")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "DT1", "Direction of Travel")
Set nodx = tv(1).Nodes.Add("DT1", tvwChild, , "Plainttif")
Set nodx = tv(1).Nodes.Add("DT1", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("DT1", tvwChild, , "Other")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "SRI1", "Streets or Roads Involved")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "Names")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "Roadway surface")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "How many lanes did each roadway have?")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "One-way or Two-way?")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "Interchange")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "Description of raod")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, , "Locality")
Set nodx = tv(1).Nodes.Add("SRI1", tvwChild, "SRI1ATCD", "Applicable traffic control devices")
Set nodx = tv(1).Nodes.Add("SRI1ATCD", tvwChild, , "Stop signs")
Set nodx = tv(1).Nodes.Add("SRI1ATCD", tvwChild, , "Stop lights")
Set nodx = tv(1).Nodes.Add("SRI1ATCD", tvwChild, , "Warning light")
Set nodx = tv(1).Nodes.Add("SRI1ATCD", tvwChild, , "Walk lights")
Set nodx = tv(1).Nodes.Add("SRI1ATCD", tvwChild, , "Traffic signing")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "FWR1", "Familiarity with route")
Set nodx = tv(1).Nodes.Add("FWR1", tvwChild, , "Plaintiff")
Set nodx = tv(1).Nodes.Add("FWR1", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "OTV1", "Obstructions to view")
Set nodx = tv(1).Nodes.Add("OTV1", tvwChild, , "Artificial lighting")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "TT1", "The Trip")
Set nodx = tv(1).Nodes.Add("TT1", tvwChild, "TT1P", "Plaintiff")
Set nodx = tv(1).Nodes.Add("TT1P", tvwChild, , "Where did trip begin?")
Set nodx = tv(1).Nodes.Add("TT1P", tvwChild, , "Where was destination?")
Set nodx = tv(1).Nodes.Add("TT1P", tvwChild, , "Purpose")
Set nodx = tv(1).Nodes.Add("TT1", tvwChild, "TT1D", "Defendant")
Set nodx = tv(1).Nodes.Add("TT1D", tvwChild, , "Where did trip begin?")
Set nodx = tv(1).Nodes.Add("TT1D", tvwChild, , "Where was destination?")
Set nodx = tv(1).Nodes.Add("TT1D", tvwChild, , "Purpose")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "NOP1", "Number of Passengers")
Set nodx = tv(1).Nodes.Add("NOP1", tvwChild, "NOP1PV", "Plaintif Vehicle")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Address")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Occupation")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Phone number")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Where located in automobile")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Injuries")
Set nodx = tv(1).Nodes.Add("NOP1PV", tvwChild, , "Relationship to driver or other occupants")
Set nodx = tv(1).Nodes.Add("NOP1", tvwChild, "NOP1DV", "Defendant Vehicle")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Address")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Occupation")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Phone number")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Where located in automobile")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Injuries")
Set nodx = tv(1).Nodes.Add("NOP1DV", tvwChild, , "Relationship to driver or other occupants")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "SPD", "Speed")
Set nodx = tv(1).Nodes.Add("SPD", tvwChild, , "Speed limit of roadways")
Set nodx = tv(1).Nodes.Add("SPD", tvwChild, , "Speed of all vehicles prior to braking")
Set nodx = tv(1).Nodes.Add("SPD", tvwChild, , "Speed of vehicle at impact")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "FKDC", "First knowledge of danger of collision")
Set nodx = tv(1).Nodes.Add("FKDC", tvwChild, , "What gave rise to notice of danger?")
Set nodx = tv(1).Nodes.Add("FKDC", tvwChild, , "Relative position of allvehicle at that time")
Set nodx = tv(1).Nodes.Add("FKDC", tvwChild, , "Approximate time")
Set nodx = tv(1).Nodes.Add("FKDC", tvwChild, , "Impact")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "EA1", "Evasive action")
Set nodx = tv(1).Nodes.Add("EA1", tvwChild, "EA1BA", "Brakes applied")
Set nodx = tv(1).Nodes.Add("EA1BA", tvwChild, , "Plaintiff")
Set nodx = tv(1).Nodes.Add("EA1BA", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("EA1", tvwChild, "EA1S", "Skidding")
Set nodx = tv(1).Nodes.Add("EA1S", tvwChild, , "Plaintiff")
Set nodx = tv(1).Nodes.Add("EA1S", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("EA1", tvwChild, "EA1H", "Horn")
Set nodx = tv(1).Nodes.Add("EA1H", tvwChild, , "Plaintiff")
Set nodx = tv(1).Nodes.Add("EA1H", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "SAL", "Signals and Lights")
Set nodx = tv(1).Nodes.Add("SAL", tvwChild, , "Brake lights")
Set nodx = tv(1).Nodes.Add("SAL", tvwChild, , "Tail Lights")
Set nodx = tv(1).Nodes.Add("SAL", tvwChild, , "Directional signal lights")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "ID1", "Impact Description")
Set nodx = tv(1).Nodes.Add("ID1", tvwChild, "ID1WPVC", "What parts of vehicle collided?")
Set nodx = tv(1).Nodes.Add("ID1WPVC", tvwChild, , "Where did each vehicle end up?")
Set nodx = tv(1).Nodes.Add("ID1WPVC", tvwChild, , "How far apart were the vehicles?")
Set nodx = tv(1).Nodes.Add("ID1WPVC", tvwChild, , "Describe any spinning or other action after impact?")
Set nodx = tv(1).Nodes.Add("ID1", tvwChild, "ID1EDV", "Exterior damage to vehicles")
Set nodx = tv(1).Nodes.Add("ID1EDV", tvwChild, "ID1EDVPV", "Plaintiff's vehicle")
Set nodx = tv(1).Nodes.Add("ID1EDVPV", tvwChild, , "Description of damage")
Set nodx = tv(1).Nodes.Add("ID1EDVPV", tvwChild, , "Cost of repair")
Set nodx = tv(1).Nodes.Add("ID1EDV", tvwChild, "ID1EDVDV", "Defendant's vehicle")
Set nodx = tv(1).Nodes.Add("ID1EDVDV", tvwChild, , "Description of damage")
Set nodx = tv(1).Nodes.Add("ID1EDVDV", tvwChild, , "Cost of repair")
Set nodx = tv(1).Nodes.Add("ID1", tvwChild, "ID1IDV", "Interior damage to vehicles")
Set nodx = tv(1).Nodes.Add("ID1IDV", tvwChild, "ID1IDD", "Describe damage to interior damage of vehicle")
Set nodx = tv(1).Nodes.Add("ID1IDD", tvwChild, , "Bent steering wheel?")
Set nodx = tv(1).Nodes.Add("ID1IDD", tvwChild, , "Broken seat?")
Set nodx = tv(1).Nodes.Add("ID1IDD", tvwChild, , "Dents in dash or interior of vehicle?")
Set nodx = tv(1).Nodes.Add("ID1IDV", tvwChild, "ID1IDO", "Other events inside vehicle")
Set nodx = tv(1).Nodes.Add("ID1IDO", tvwChild, , "Glasses come off?")
Set nodx = tv(1).Nodes.Add("ID1IDO", tvwChild, , "Things fly around inside vehicle?")
Set nodx = tv(1).Nodes.Add("ID1IDO", tvwChild, , "Other?")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "ROV11", "Removal of vehicles")
Set nodx = tv(1).Nodes.Add("ROV11", tvwChild, "ROV11H", "How was each vehicle removed from the scene?")
Set nodx = tv(1).Nodes.Add("ROV11H", tvwChild, , "Plaintiff")
Set nodx = tv(1).Nodes.Add("ROV11H", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("ROV11H", tvwChild, , "Other")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "CAS1", "Conversation at scene")
Set nodx = tv(1).Nodes.Add("CAS1", tvwChild, "CAS1D", "Detail each statement made by:")
Set nodx = tv(1).Nodes.Add("CAS1D", tvwChild, , "Plaintiff")
Set nodx = tv(1).Nodes.Add("CAS1D", tvwChild, , "Defendant")
Set nodx = tv(1).Nodes.Add("CAS1D", tvwChild, , "Each witness")
Set nodx = tv(1).Nodes.Add("CAS1D", tvwChild, , "Investigating officer")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "FOBAI1", "Force on body at impact")
Set nodx = tv(1).Nodes.Add("FOBAI1", tvwChild, , "Seatbelt?")
Set nodx = tv(1).Nodes.Add("FOBAI1", tvwChild, , "Describe what happened to body inside vehicle at moment of impact?")
Set nodx = tv(1).Nodes.Add("FOBAI1", tvwChild, , "Did body strike interior of vehicle?")
Set nodx = tv(1).Nodes.Add("FOBAI1", tvwChild, , "Loss of consciousness?")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "TAS12", "Treatment at scene")
Set nodx = tv(1).Nodes.Add("TAS12", tvwChild, , "Aid car")
Set nodx = tv(1).Nodes.Add("TAS12", tvwChild, , "Ambulance")
Set nodx = tv(1).Nodes.Add("TAS12", tvwChild, , "Other...")
Set nodx = tv(1).Nodes.Add("r4", tvwChild, "ROPFS1", "Removal of plaintiff from scene")
Set nodx = tv(1).Nodes.Add("ROPFS1", tvwChild, , "By self")
Set nodx = tv(1).Nodes.Add("ROPFS1", tvwChild, , "By Aid car")
Set nodx = tv(1).Nodes.Add("ROPFS1", tvwChild, , "By Ambulance")
Set nodx = tv(1).Nodes.Add("ROPFS1", tvwChild, , "Other...")
Set nodx = tv(1).Nodes.Add(, , "r5", "DESCRIPTION OF INJURIES,starting at the top of the tip of the toes...")
Set nodx = tv(1).Nodes.Add("r5", tvwChild, , "Immediately after injury")
Set nodx = tv(1).Nodes.Add("r5", tvwChild, , "When first seen by health care provider")
Set nodx = tv(1).Nodes.Add("r5", tvwChild, , "Several days later when all injuries could be appreciated")
Set nodx = tv(1).Nodes.Add(, , "r6", "MEDICAL TREATMENT...")
Set nodx = tv(1).Nodes.Add("r6", tvwChild, "r6H", "Hospitals")
Set nodx = tv(1).Nodes.Add("r6H", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("r6H", tvwChild, , "Address")
Set nodx = tv(1).Nodes.Add("r6H", tvwChild, , "Dates")
Set nodx = tv(1).Nodes.Add("r6", tvwChild, "r6D", "Doctors")
Set nodx = tv(1).Nodes.Add("r6D", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("r6D", tvwChild, , "Address")
Set nodx = tv(1).Nodes.Add("r6D", tvwChild, , "Dates")
Set nodx = tv(1).Nodes.Add("r6", tvwChild, "r6T", "Therapists")
Set nodx = tv(1).Nodes.Add("r6T", tvwChild, , "Name")
Set nodx = tv(1).Nodes.Add("r6T", tvwChild, , "Address")
Set nodx = tv(1).Nodes.Add("r6T", tvwChild, , "Dates")
Set nodx = tv(1).Nodes.Add(, , "r7", "WAGE LOSS...")
Set nodx = tv(1).Nodes.Add("r7", tvwChild, "r7D", "Dates missed")
Set nodx = tv(1).Nodes.Add("r7D", tvwChild, , "Rate of compensation")
Set nodx = tv(1).Nodes.Add("r7D", tvwChild, , "Reason for wage loss")
Set nodx = tv(1).Nodes.Add(, , "r8", "OUT-OF-POCKET EXPENSES...")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Physicians")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Hospitals")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Ambulance")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Drugs")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Crutches,appliances")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Domestic help")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Auto repair")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Lost wages")
Set nodx = tv(1).Nodes.Add("r8", tvwChild, , "Other")
Set nodx = tv(1).Nodes.Add(, , "r9", "PLAINTIFF'S PERIOR MEDICAL HISTORY...")
Set nodx = tv(1).Nodes.Add("r9", tvwChild, , "Past hospitalizations")
Set nodx = tv(1).Nodes.Add("r9", tvwChild, , "Past serious injuries")
Set nodx = tv(1).Nodes.Add("r9", tvwChild, , "Past accidents of broken bones or other serious injuries")
Set nodx = tv(1).Nodes.Add("r9", tvwChild, , "Past auto accidents")
Set nodx = tv(1).Nodes.Add("r9", tvwChild, , "Past claims")
Set nodx = tv(1).Nodes.Add("r9", tvwChild, , "Past injuries or symptoms to same area of body injured in injury or incident")
Set nodx = tv(1).Nodes.Add(, , , "IMPRESSION OF CLIENT...")
Set nodx = tv(1).Nodes.Add(, , , "SETTELMENT OFFERS...")
Set nodx = tv(1).Nodes.Add(, , , "TOTAL MEDICAL SPECIALS...")
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
Private Sub ref1(Job_ID_Incoming As String, Tabel As String)
Dim db As Database
Dim sqlq As String
Dim d_b As String
d_b = App.Path & "\Shabshab.mdb"
Set db = OpenDatabase(d_b)
sqlq = "Select * from" & " " & Tabel & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
'MsgBox sqlq
Select Case Tabel
Case "Interview"
    Set Data3.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "Defendant"
    Set Data4.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "Facts"
    Set Data5.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "Automobile_C1"
    Set Data6.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "Automobile_C2"
    Set Data7.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "Automobile_C3"
    Set Data8.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "D_Injuries"
    Set Data9.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
Case "Wage_Loss"
    Set Data10.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
End Select
End Sub

Private Sub ref2(Job_ID_Incoming)
Dim db As Database
Dim sqlq As String
Dim d_b As String
d_b = App.Path & "\Shabshab.mdb"
Set db = OpenDatabase(d_b)
sqlq = "Select * from Expenses_Account Where Job_ID = '" _
& Job_ID_Incoming & "'"
'MsgBox sqlq
Set Data1.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
End Sub


Private Sub ref111(Job_ID_Incoming As String)
Dim db As Database
Dim sqlq As String
Dim d_b As String
d_b = App.Path & "\Shabshab.mdb"
Set db = OpenDatabase(d_b)
sqlq = "Select * from" & " " & "Interview" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data3.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "Defendant" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data4.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "Facts" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data5.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "Automobile_C1" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data6.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "Automobile_C2" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data7.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "Automobile_C3" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data8.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "D_Injuries" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data9.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
sqlq = "Select * from" & " " & "Wage_Loss" & " " & "Where Job_ID = '" _
& Job_ID_Incoming & "'"
    Set Data10.Recordset = db.OpenRecordset(sqlq, dbOpenDynaset)
End Sub


Private Sub note_pad_load()
'*****************
Dim i As Integer
If (Dir(App.Path & "\Jobs\" & Form1.Label25.Caption & "\Notes\Note.txt")) <> "Note.txt" Then
'PWait.Label3.Caption = "No Notes was set for Job ID: " & Form1.Label25.Caption
Else
Dim ss, sss As String
Dim intFileNum As Integer
i = MsgBox("There are some related Notes available about the Job ID: " & Form1.Label25.Caption, vbInformation, "Attorney Master Alert")
intFileNum = FreeFile
ss = ""
sss = ""
Open App.Path & "\Jobs\" & Form1.Label25.Caption & "\Notes\Note.txt" For Input As #intFileNum
Do
Line Input #intFileNum, ss
'sss = sss & ss
sss = sss & ss & vbCrLf
Loop Until EOF(intFileNum)
Close #intFileNum
'*******************
Text4.Text = sss
End If
End Sub

Sub ShowFolderList55(folderspec)
    Dim fs, f, f1, fc, s
    Dim i As Integer
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    i = 0
    For Each f1 In fc
        s = f1.Name
        If i <= 13 Then
        c(i).Caption = s
        i = i + 1
        End If
    Next
End Sub

