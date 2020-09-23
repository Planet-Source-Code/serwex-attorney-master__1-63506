VERSION 5.00
Begin VB.Form Form7 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master - [Transfer Jobs]"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   Icon            =   "Your Transfered.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   4410
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Transfer"
      Height          =   330
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2625
      Width           =   1590
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2205
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2100
      Width           =   2115
   End
   Begin VB.ListBox List2 
      Height          =   1230
      ItemData        =   "Your Transfered.frx":030A
      Left            =   105
      List            =   "Your Transfered.frx":030C
      TabIndex        =   2
      Top             =   2100
      Width           =   2010
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Quit"
      Height          =   330
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3045
      Width           =   1590
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Accept"
      Height          =   330
      Left            =   2730
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   1590
   End
   Begin VB.ListBox List1 
      Height          =   1185
      ItemData        =   "Your Transfered.frx":030E
      Left            =   105
      List            =   "Your Transfered.frx":0310
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   420
      Width           =   2010
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   105
      X2              =   4305
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label3 
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
      Left            =   2205
      TabIndex        =   8
      Top             =   1785
      Width           =   330
   End
   Begin VB.Label Label2 
      Caption         =   "Active Jobs:"
      Height          =   225
      Left            =   105
      TabIndex        =   7
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Received Jobs:"
      Height          =   225
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   1485
   End
End
Attribute VB_Name = "Form7"
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
Dim i As Integer
For i = 0 To List1.ListCount - 1
If List1.Selected(i) = True Then
Kill (App.Path & "\Users\" & Form1.Combo6.Text & "\Transfered Job\" & List1.List(i))
List1.RemoveItem (i)
End If
Next i
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim i As Integer
Dim intFileNum As Integer
If Combo1.Text = "" Then
i = MsgBox("Don't forgot 'Your desiered person' no ever!", vbExclamation, "Attorney Master Alert")
Exit Sub
End If
intFileNum = FreeFile
Open App.Path & "\Transfered Jobs\" & List2.Text For Append As #intFileNum
Print #intFileNum, Form1.Combo6.Text
Print #intFileNum, Combo1.Text
Print #intFileNum, Now
Close #intFileNum
End Sub

Private Sub Form_Load()
Call ShowsubFolderList(App.Path & "\Users")
Call Active_job_file5(App.Path & "\Users\" & Form1.Combo6.Text & "\Active Jobs")
Call Active_job_file6(App.Path & "\Users\" & Form1.Combo6.Text & "\Transfered Job")
End Sub
Sub ShowsubFolderList(folderspec)
    Dim fs, f, f1, s, sf
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
    For Each f1 In sf
        s = f1.Name
        Combo1.AddItem (s)
    Next
End Sub


Sub Active_job_file5(folderspec)
    Dim i As Integer
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
       s = f1.Name
       List2.AddItem (s)
    Next
End Sub

Sub Active_job_file6(folderspec)
    Dim i As Integer
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
       s = f1.Name
       List1.AddItem (s)
    Next
End Sub


