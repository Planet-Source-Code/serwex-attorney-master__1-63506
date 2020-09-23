VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master - [Messages!]"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "Messages.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Send a Message..."
      Height          =   5160
      Left            =   5040
      TabIndex        =   17
      Top             =   105
      Width           =   4950
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&New"
         Height          =   330
         Left            =   3780
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   315
         Width           =   1065
      End
      Begin VB.TextBox Text8 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   945
         TabIndex        =   9
         Top             =   1260
         Width           =   3900
      End
      Begin VB.TextBox Text6 
         ForeColor       =   &H00FF0000&
         Height          =   2745
         Left            =   105
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1680
         Width           =   4740
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   945
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   630
         Width           =   2745
      End
      Begin VB.TextBox Text7 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   945
         Width           =   2745
      End
      Begin VB.TextBox Text5 
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   315
         Width           =   2745
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Quit"
         Height          =   435
         Left            =   3570
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   4620
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Send"
         Height          =   435
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   4620
         Width           =   1275
      End
      Begin VB.Label Label8 
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
         Left            =   1470
         TabIndex        =   23
         Top             =   4620
         Width           =   2010
      End
      Begin VB.Label Label7 
         Caption         =   "Subject:"
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
         TabIndex        =   22
         Top             =   1365
         Width           =   1065
      End
      Begin VB.Label Label6 
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
         Left            =   105
         TabIndex        =   20
         Top             =   1050
         Width           =   645
      End
      Begin VB.Label Label5 
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
         Left            =   105
         TabIndex        =   19
         Top             =   735
         Width           =   540
      End
      Begin VB.Label Label4 
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
         Left            =   105
         TabIndex        =   18
         Top             =   420
         Width           =   540
      End
   End
   Begin VB.TextBox Text4 
      ForeColor       =   &H00FF0000&
      Height          =   1905
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3360
      Width           =   4845
   End
   Begin VB.TextBox Text3 
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2940
      Width           =   2745
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2625
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2310
      Width           =   2745
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Refresh"
      Height          =   330
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   4845
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Delete"
      Height          =   330
      Left            =   3780
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2310
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   1635
      ItemData        =   "Messages.frx":030A
      Left            =   105
      List            =   "Messages.frx":030C
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   525
      Width           =   4845
   End
   Begin VB.Label Label3 
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
      Left            =   105
      TabIndex        =   16
      Top             =   3045
      Width           =   645
   End
   Begin VB.Label Label2 
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
      Left            =   105
      TabIndex        =   15
      Top             =   2730
      Width           =   540
   End
   Begin VB.Label Label1 
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
      Left            =   105
      TabIndex        =   14
      Top             =   2415
      Width           =   540
   End
End
Attribute VB_Name = "Form6"
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
Dim str, str1 As String
Dim intFileNum As Integer
If Combo1.Text = "" Then
i = MsgBox("Don't forgot 'Your desiered person' no ever!", vbExclamation, "Attorney Master Alert")
Exit Sub
End If
If Text6.Text = "" Then
i = MsgBox("Don't forgot the 'Body of Message' no ever!", vbExclamation, "Attorney Master Alert")
Exit Sub
End If
If Label8.Caption <> "" Then
i = MsgBox("Don't forgot the 'New Button' no ever!", vbExclamation, "Attorney Master Alert")
Exit Sub
End If
If Text8.Text = "" Then
i = MsgBox("Don't forgot the subject no ever!", vbExclamation, "Attorney Master Alert")
Exit Sub
End If
intFileNum = FreeFile
'Sub OpenTextFileTest()
'    Const ForReading = 1, ForWriting = 2, ForAppending = 3
'    Dim fs, f
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.OpenTextFile("c:\testfile.txt", ForAppending, TristateFalse)
'    f.Write "Hello world!"
'    f.Close
'End Sub
'Set fs = CreateObject("Scripting.FileSystemObject")
'Set a = fs.CreateTextFile("c:\testfile.txt", True)
'a.WriteLine ("This is a test.")
'a.Close
Open App.Path & "\Users\" & Combo1.Text & "\Messages\" & Text8.Text For Append As #intFileNum
str = ""
str1 = ""
Print #intFileNum, "/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\/\"
Print #intFileNum, Text5.Text
Print #intFileNum, Combo1.Text
Text7.Text = Now
'str = "'" & CStr(Date) & "'" & "   " & "'" & CStr(Time) & "'"
'Text7.Text = str
'str = ""
Print #intFileNum, Text7.Text
For i = 1 To Len(Text6.Text)
str1 = Mid(Text6.Text, i, 1)
If str1 = Chr(13) Or i = Len(Text6.Text) Then
   If i = Len(Text6.Text) Then
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
Label8.Caption = "Your message has been sent..."
End Sub

Private Sub Command5_Click()
Call ShowsubFolderList1(App.Path & "\Users\" & Form1.Combo6.Text & "\Messages")
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Dim i As Integer
i = MsgBox("Realy remove the messages?", vbYesNo + vbQuestion + vbDefaultButton2, "Attorney Master Alert")
If i = vbNo Then
Exit Sub
End If
For i = 0 To List1.ListCount - 1
   If List1.Selected(i) = True Then
      Kill (App.Path & "\Users\" & Form1.Combo6.Text & "\Messages\" & List1.List(i))
   End If
Next i
Call ShowsubFolderList1(App.Path & "\Users\" & Form1.Combo6.Text & "\Messages")
End Sub


Private Sub Command6_Click()
Label8.Caption = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
End Sub

Private Sub Form_Load()
Text5.Text = Form1.Combo6.Text
Call ShowsubFolderList(App.Path & "\Users")
Call ShowsubFolderList1(App.Path & "\Users\" & Form1.Combo6.Text & "\Messages")
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
Sub ShowsubFolderList1(folderspec)
    Dim i As Integer
    Dim fs, f, f1, fc, s
List1.Clear
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
       s = f1.Name
       List1.AddItem (s)
    Next
End Sub


Private Sub List1_Click()
'Sub OpenTextFileTest()
'    Const ForReading = 1, ForWriting = 2, ForAppending = 3
'    Dim fs, f
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.OpenTextFile("c:\testfile.txt", ForAppending, TristateFalse)
'    f.Write "Hello world!"
'    f.Close
'End Sub
'Sub ShowFolderSize(filespec)
'    Dim fs, f, s
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    Set f = fs.GetFolder(filespec)
'    s = UCase(f.Name) & " uses " & f.Size & " bytes."
'    MsgBox s, 0, "Folder Size Info"
'End Sub
'Sub TextStreamTest()
'    Const ForReading = 1, ForWriting = 2, ForAppending = 3
'    Const TristateUseDefault = -2, TristateTrue = -1,
'TristateFalse = 0
'    Dim fs, f, ts, s
'    Set fs = CreateObject("Scripting.FileSystemObject")
'    fs.CreateTextFile "test1.txt"            'Create a file
'    Set f = fs.GetFile("test1.txt")
'    Set ts = f.OpenAsTextStream(ForWriting, TristateUseDefault)
'    ts.Write "Hello World"
'    ts.Close
'    Set ts = f.OpenAsTextStream(ForReading, TristateUseDefault)
'    s = ts.ReadLine
'    MsgBox s
'    ts.Close
'End Sub
'*****************
Dim i As Integer
Dim ss, sss As String
Dim intFileNum As Integer
intFileNum = FreeFile
ss = ""
sss = ""
Open App.Path & "\Users\" & Form1.Combo6.Text & "\Messages\" & List1.Text For Input As #intFileNum
'Line Input #
Line Input #intFileNum, ss
Line Input #intFileNum, ss
Text1.Text = ss
Line Input #intFileNum, ss
Text2.Text = ss
Line Input #intFileNum, ss
Text3.Text = ss
ss = ""
Do
Line Input #intFileNum, ss
'sss = sss & ss
sss = sss & ss & vbCrLf
Loop Until EOF(intFileNum)
Close #intFileNum
'*******************
Text4.Text = sss
End Sub
