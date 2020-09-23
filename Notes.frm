VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master - [Notes Area]"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "Notes.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "A&ccept"
      Height          =   435
      Left            =   2940
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4725
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Append"
      Height          =   435
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4725
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Quit"
      Height          =   435
      Left            =   4305
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4725
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&New"
      Height          =   435
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4725
      Width           =   1170
   End
   Begin VB.TextBox Text1 
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
      Height          =   4425
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   105
      Width           =   5370
   End
End
Attribute VB_Name = "Form5"
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
Text1.Text = ""
Text1.SetFocus
If (Dir(App.Path & "\Users\" & Form1.Combo6.Text & "\Notes\Note.txt")) = "Note.txt" Then
Kill (App.Path & "\Users\" & Form1.Combo6.Text & "\Notes\Note.txt")
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Text1.Text = Text1.Text & vbCrLf & "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\" & vbCrLf
Text1.SetFocus
End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim str, str1 As String
Dim intFileNum As Integer
intFileNum = FreeFile
Open App.Path & "\Users\" & Form1.Combo6.Text & "\Notes\Note.txt" For Append As #intFileNum
str = ""
str1 = ""
For i = 1 To Len(Text1.Text)
str1 = Mid(Text1.Text, i, 1)
If str1 = Chr(13) Or i = Len(Text1.Text) Then
   If i = Len(Text1.Text) Then
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
End Sub

Private Sub Form_Load()
'*****************
Dim i As Integer
If (Dir(App.Path & "\Users\" & Form1.Combo6.Text & "\Notes\Note.txt")) <> "Note.txt" Then
i = MsgBox("There is no Notes!", vbInformation, "Attorney Master Alert")
Else
Dim ss, sss As String
Dim intFileNum As Integer
intFileNum = FreeFile
ss = ""
sss = ""
Open App.Path & "\Users\" & Form1.Combo6.Text & "\Notes\Note.txt" For Input As #intFileNum
Do
Line Input #intFileNum, ss
'sss = sss & ss
sss = sss & ss & vbCrLf
Loop Until EOF(intFileNum)
Close #intFileNum
'*******************
Text1.Text = sss
End If
End Sub
