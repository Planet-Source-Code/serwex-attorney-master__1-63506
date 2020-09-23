VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3780
   Icon            =   "ChangePass.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   3780
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Reject"
      Height          =   330
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1575
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Accept"
      Height          =   330
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1575
      Width           =   1275
   End
   Begin VB.TextBox Text3 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1050
      Width           =   2430
   End
   Begin VB.TextBox Text2 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1260
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   630
      Width           =   2430
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   210
      Width           =   2430
   End
   Begin VB.Label Label3 
      Caption         =   "Confirm:"
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "New Password:"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   735
      Width           =   1170
   End
   Begin VB.Label Label1 
      Caption         =   "User Name:"
      Height          =   225
      Left            =   105
      TabIndex        =   5
      Top             =   315
      Width           =   960
   End
End
Attribute VB_Name = "Form8"
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

Private Sub Command1_Click()
If Text2.Text = Text3.Text Then
Form1.Data3.Recordset.Edit
Form1.Data3.Recordset("Password").Value = Text2.Text
Form1.Data3.Recordset.Update
Unload Me
Form1.Show
Else
Text2.Text = ""
Text3.Text = ""
Dim i As Integer
i = MsgBox("Try the password again.", vbInformation, "Attorney Master Alert")
Text2.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
Dim pass As String
Dim db As Database
Text1.Text = Form1.Combo6.Text
Set db = OpenDatabase(App.Path & "\ShabShab.mdb")
Set Form1.Data3.Recordset = db.OpenRecordset("select * from Users where Name='" & Form1.Combo6.Text & "'", dbOpenDynaset)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Data3.Recordset.Close
End Sub
