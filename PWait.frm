VERSION 5.00
Begin VB.Form PWait 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1905
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   3945
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "PWait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   3615
      Begin VB.Label Label3 
         Caption         =   "Connecting to Word ,Pleseae Wait..."
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
         Left            =   210
         TabIndex        =   2
         Top             =   840
         Width           =   3165
      End
      Begin VB.Label Label2 
         Caption         =   "Attorney Master Status:"
         Height          =   225
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   2955
      End
   End
End
Attribute VB_Name = "PWait"
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

