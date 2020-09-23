VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Attorney Master [Mail Manager]"
   ClientHeight    =   8625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Discard Changes"
      Height          =   330
      Left            =   10185
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "This buttons are managing the available Jobs."
      Top             =   7770
      Width           =   1590
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H0080C0FF&
      Caption         =   "&Save Changes"
      Height          =   330
      Left            =   8400
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "This buttons are managing the available Jobs."
      Top             =   7770
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Font Style"
      Height          =   330
      Left            =   210
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Select the Font Style for word outputs and Print."
      Top             =   1575
      Width           =   1380
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View &Mode"
      Height          =   225
      Left            =   10605
      TabIndex        =   38
      Top             =   105
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog com 
      Left            =   3570
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add New File Style:"
      Height          =   6000
      Left            =   105
      TabIndex        =   5
      Top             =   2100
      Width           =   5580
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Add &New"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5460
         Width           =   1275
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Edit"
         Height          =   330
         Left            =   1470
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5460
         Width           =   1275
      End
      Begin VB.ComboBox c 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   7
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   3885
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   6
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   3465
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   5
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   3045
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   4
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   2625
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   2
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   1785
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   1
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1365
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H0080C0FF&
         Height          =   315
         Index           =   0
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   945
         Width           =   2850
      End
      Begin VB.ComboBox c 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Index           =   3
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2205
         Width           =   2850
      End
      Begin VB.TextBox Text3 
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
         Left            =   1890
         TabIndex        =   27
         Top             =   4725
         Width           =   2745
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Reject"
         Height          =   330
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5460
         Width           =   1275
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   7
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3885
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   6
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3465
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   5
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3045
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   4
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2625
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   3
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2205
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   2
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1785
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   1
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1365
         Width           =   645
      End
      Begin VB.CommandButton co 
         BackColor       =   &H0080C0FF&
         Caption         =   "Add >"
         Height          =   330
         Index           =   0
         Left            =   4830
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   945
         Width           =   645
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "A&ccept"
         Height          =   330
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   5460
         Width           =   1275
      End
      Begin VB.Label Label14 
         Caption         =   "New File Name:"
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   4830
         Width           =   1695
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         Caption         =   ".Text"
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
         Left            =   4725
         TabIndex        =   24
         Top             =   4725
         Width           =   645
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   210
         X2              =   5355
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label Label11 
         Caption         =   "Wage Loss:"
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   3990
         Width           =   1800
      End
      Begin VB.Label Label10 
         Caption         =   "Related Facts:"
         Height          =   225
         Left            =   210
         TabIndex        =   13
         Top             =   1890
         Width           =   1800
      End
      Begin VB.Label Label9 
         Caption         =   "Description of Injuries:"
         Height          =   225
         Left            =   210
         TabIndex        =   12
         Top             =   3570
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Automobile C(3):"
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   3150
         Width           =   1485
      End
      Begin VB.Label Label7 
         Caption         =   "Automobile C(2):"
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   2730
         Width           =   1380
      End
      Begin VB.Label Label6 
         Caption         =   "Automobile C(1):"
         Height          =   225
         Left            =   210
         TabIndex        =   9
         Top             =   2310
         Width           =   1485
      End
      Begin VB.Label Label5 
         Caption         =   "Defendant:"
         Height          =   225
         Left            =   210
         TabIndex        =   8
         Top             =   1470
         Width           =   1380
      End
      Begin VB.Label Label4 
         Caption         =   "Interview:"
         Height          =   225
         Left            =   210
         TabIndex        =   7
         Top             =   1050
         Width           =   1380
      End
      Begin VB.Label Label3 
         Caption         =   "Select your desiered fields from the colections below:"
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
         Left            =   210
         TabIndex        =   6
         Top             =   315
         Width           =   4740
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Related mails:"
      Height          =   1275
      Left            =   105
      TabIndex        =   4
      ToolTipText     =   "All sample Mails."
      Top             =   105
      Width           =   2745
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Remove"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   840
         Width           =   1065
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Add to Job"
         Height          =   330
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   840
         Width           =   1065
      End
      Begin VB.ComboBox Combo10 
         BackColor       =   &H00C0FFC0&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   420
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " View Available Files:"
      Height          =   1275
      Left            =   2940
      TabIndex        =   1
      ToolTipText     =   "Only mails for this Job."
      Top             =   105
      Width           =   2745
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Delete"
         Height          =   330
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Word"
         Height          =   330
         Left            =   945
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Print"
         Height          =   330
         Left            =   1785
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   840
         Width           =   855
      End
      Begin VB.ComboBox Combo11 
         BackColor       =   &H00FFFFC0&
         Height          =   315
         Left            =   105
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   420
         Width           =   2535
      End
   End
   Begin VB.TextBox Text2 
      ForeColor       =   &H00FF0000&
      Height          =   7260
      Left            =   5775
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   420
      Width           =   6000
   End
   Begin VB.Label Label1 
      Caption         =   "Mail Content:"
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
      Left            =   5880
      TabIndex        =   41
      Top             =   105
      Width           =   1905
   End
   Begin VB.Line Line4 
      X1              =   105
      X2              =   105
      Y1              =   1470
      Y2              =   1995
   End
   Begin VB.Line Line2 
      X1              =   105
      X2              =   1680
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line Line5 
      X1              =   1680
      X2              =   1680
      Y1              =   1470
      Y2              =   1995
   End
   Begin VB.Line Line3 
      X1              =   105
      X2              =   1680
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Label Label15 
      BackColor       =   &H80000014&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "This is your output default!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   1785
      TabIndex        =   40
      Top             =   1470
      Width           =   3900
   End
End
Attribute VB_Name = "Form4"
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
Check1.Caption = "&Edit Mode"
Text2.Locked = False
Call Combo11_Click
'*************************
Else
Check1.Caption = "&View Mode"
Text2.Locked = True
Call fix_text
End If
End Sub

Private Sub fix_text()
Dim str, str1, chk As String
Dim cond As Boolean
Dim i As Integer
Dim ii As Integer
Dim iii As Integer
Dim j As Integer
Dim jj As Integer
Dim k As Integer
'/\/\/\/\/\/\/\/\/\/\/\//\
str1 = ""
jj = 0
For i = 1 To Len(Text2.Text)
str1 = Mid(Text2.Text, i, 1)
If str1 = "¬" Then jj = jj + 1
Next i
'/\/\/\/\/\/\/\/\/\/\/\//\
For k = 1 To jj / 2
str1 = ""
str = ""
cond = False
iii = 15
ii = 0
'i = 0
For i = 1 To Len(Text2.Text)
str1 = Mid(Text2.Text, i, 1)
'*******
If cond = True And i = ii + 1 Then
   iii = CInt(str1)
End If
If str1 = "¬" Then
   If cond = False Then
   cond = True
   ii = i
   Else
   cond = False
       For j = 0 To c(iii).ListCount
       chk = c(iii).List(j)
           If Mid(str, 2, Len(str)) = chk Then
           Text2.SelStart = ii - 1
           Text2.SelLength = i - ii + 1
'           MsgBox chk & "  " & j & "   " & iii & "   " & c(iii).ListCount
           Call from_database(iii, j)
           End If
       Next j
'   str1 = ""
'   str = ""
'   iii = 15
'   ii = 0
   Exit For
   End If
End If
If cond = True And iii <> 15 Then
   str = str & str1
End If
'*********
Next i
Next k
End Sub



Private Sub co_Click(Index As Integer)
Dim i As Integer
If c(Index).Text = "" Then
i = MsgBox("No fields selected!", vbExclamation, "Attorney Master Alert")
Else
Clipboard.Clear
'¬ 170
Clipboard.SetText "¬" & Index & c(Index).Text & "¬"
Text2.SetFocus
SendKeys "^V"
End If
End Sub

Private Sub Combo10_Change()
'******************


End Sub


Private Sub Combo11_Click()
'*****************
Dim i As Integer
If Combo11.Text = "" Then
i = MsgBox("There is no available file!", vbInformation, "Attorney Master Alert")
Else
Dim ss, sss As String
Dim intFileNum As Integer
intFileNum = FreeFile
Text3.Text = ""
ss = ""
sss = ""
Open App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt\" & Combo11.Text For Input As #intFileNum
Do
Line Input #intFileNum, ss
'sss = sss & ss
sss = sss & ss & vbCrLf
Loop Until EOF(intFileNum)
Close #intFileNum
'*******************
Text2.Text = sss
If Check1.Value = 0 Then
 Call fix_text
End If
Text2.SetFocus
End If
End Sub

Private Sub Command1_Click()
Dim j As Integer
Dim fname As String
Dim word_s As Object
On Error GoTo eee
If (Dir(App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Doc\" & Form2.l(150).Caption & Combo11.Text & ".Doc")) = Form2.l(150).Caption & Combo11.Text & ".Doc" Then
PWait.Label3.Caption = "Connecting to Word for Printing..."
PWait.Show
Set word_s = CreateObject("word.application")
fname = App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Doc\" & Form2.l(150).Caption & Combo11.Text & ".Doc"
word_s.Documents.Open (fname)
word_s.Documents(1).PrintOut
word_s.Documents(1).Close
word_s.quit
Unload PWait
Else
j = MsgBox("At first Create the Word file.", vbInformation, "Attorney Master Alert")
End If
Exit Sub
eee:
j = MsgBox(err.Description, vbExclamation, "Attorney master Alert")
End Sub

Private Sub Command10_Click()
Text2.Locked = False
Text2.Text = ""
Text3.Text = ""
Check1.Enabled = False
Text2.SetFocus
End Sub

Private Sub Command14_Click()
Text2.Locked = True
Text2.Text = ""
Text3.Text = ""
Check1.Enabled = True
End Sub

Private Sub Command15_Click()
Dim a, bb As String
Dim fs As Object
Dim i As Integer
If Combo10.Text = "" Then
i = MsgBox("At first select a template!", vbInformation, "Attorney Master")
Else
bb = App.Path & "\Mail Templates\" & Combo10.Text
a = App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt\"
Set fs = CreateObject("Scripting.FileSystemObject")
fs.CopyFile bb, a
Call ShowFolderList1(App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt")
End If
End Sub

Private Sub Command2_Click()
com.Flags = cdlCFBoth Or cdlCFEffects
com.ShowFont
Label15.Font.Name = com.FontName
Label15.Font.Size = com.FontSize
Label15.Font.Bold = com.FontBold
Label15.Font.Italic = com.FontItalic
End Sub


Private Sub Command3_Click()
Dim i As Integer
If Text3.Text = "" Then
i = MsgBox("Enter the file name above.", vbInformation, "Attorney Master Alert")
Else
Dim str, str1 As String
Dim intFileNum As Integer
intFileNum = FreeFile
Text2.Locked = True
Check1.Enabled = True
Open App.Path & "\Mail Templates\" & Text3.Text For Output As #intFileNum
str = ""
str1 = ""
For i = 1 To Len(Text2.Text)
str1 = Mid(Text2.Text, i, 1)
If str1 = Chr(13) Or i = Len(Text2.Text) Then
   If i = Len(Text2.Text) Then
   Print #intFileNum, str & str1
   Else
   Print #intFileNum, str
   End If
'Print #intFileNum, str
str = ""
Else
    If str1 <> Chr(10) Then
    str = str & str1
    End If
End If
Next i
Close #intFileNum
Call ShowFolderList(App.Path & "\Mail Templates")
End If
End Sub

Private Sub Command4_Click()
Dim i, j As Integer
Dim str, str1, fname As String
'************
Dim word_s As Object
If Combo11.Text = "" Then
j = MsgBox("You must select a template to create the word document.", vbInformation, "Attorney Master Alert")
Exit Sub
End If
If Check1.Value = 1 Then
j = MsgBox("You must be in View Mode to create the word document.", vbInformation, "Attorney Master Alert")
Exit Sub
End If
On Error GoTo eee
PWait.Show
PWait.Label3.Caption = "Connecting to Word ,Pleseae Wait..."
Set word_s = CreateObject("word.application")
word_s.Documents.Add
word_s.Documents(1).content.Font.Name = Label15.FontName
word_s.Documents(1).content.Font.Size = Label15.FontSize
word_s.Documents(1).content.Font.Bold = Label15.FontBold
'**********************
str = ""
str1 = ""
For i = 1 To Len(Text2.Text)
str1 = Mid(Text2.Text, i, 1)
If str1 = Chr(13) Or i = Len(Text2.Text) Then
word_s.Documents(1).content.insertafter _
Text:=str '& Chr(13)
str = ""
Else
str = str & str1
End If
Next i
fname = App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Doc\" & Form2.l(150).Caption & Combo11.Text & ".Doc"
word_s.Documents(1).SaveAs fname
PWait.Label3.Caption = "Connection Complete."
'MsgBox word_s.Documents(1).Name
'Set myRange = word_s.Documents(1).Range(Start:=0, End:=10)
'myRange.Bold = True
'word_s.Documents(1).Range(Start:=0, End:=0).InsertBefore Text:="Hi" & vbCrLf
'word_s.Documents(1).Save
'word_s.PrintPreview = True
'word_s.Documents(1).PrintOut
word_s.Documents(1).Close
word_s.quit
Unload PWait
Exit Sub
eee:
j = MsgBox(err.Description, vbExclamation, "Attorney Master Word Connection Error!")
End Sub


Private Sub Command5_Click()
Dim i As Integer
If Combo11.Text = "" Then
i = MsgBox("At first select a Template!", vbInformation, "Attorney Master Alert")
Else
i = MsgBox("Realy remove the file: " & Combo11.Text, vbYesNoCancel + vbQuestion + vbDefaultButton2, "Attorney Master Question")
If i = vbYes Then
Kill (App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt\" & Combo11.Text)
Call ShowFolderList1(App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt")
End If
End If
End Sub

Private Sub Command6_Click()
Dim i As Integer
If Combo10.Text = "" Then
i = MsgBox("At first select a Template!", vbInformation, "Attorney Master Alert")
Else
i = MsgBox("Realy remove the file: " & Combo10.Text, vbYesNoCancel + vbQuestion + vbDefaultButton2, "Attorney Master Question")
If i = vbYes Then
Kill (App.Path & "\Mail Templates\" & Combo10.Text)
Call ShowFolderList(App.Path & "\Mail Templates")
End If
End If
End Sub

Private Sub Command7_Click()
Dim i As Integer
If Check1.Value = 0 Then
i = MsgBox("You must be in Edit Mode to save changes.", vbInformation, "Attorney Master Alert")
Exit Sub
End If
If Combo11.Text = "" Then
i = MsgBox("Select a template file to save changes.", vbInformation, "Attorney Master Alert")
Else
Dim str, str1 As String
Dim intFileNum As Integer
intFileNum = FreeFile
Open App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt\" & Combo11.Text For Output As #intFileNum
str = ""
str1 = ""
For i = 1 To Len(Text2.Text)
str1 = Mid(Text2.Text, i, 1)
If str1 = Chr(13) Or i = Len(Text2.Text) Then
   If i = Len(Text2.Text) Then
   Print #intFileNum, str & str1
   Else
   Print #intFileNum, str
   End If
'Print #intFileNum, str
str = ""
Else
    If str1 <> Chr(10) Then
    str = str & str1
    End If
End If
Next i
Close #intFileNum
Call ShowFolderList1(App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt")
End If
End Sub

Private Sub Command8_Click()
Dim i As Integer
If Combo11.Text = "" Then
i = MsgBox("There is no template to discard its changes!", vbInformation, "Attorney Master Alert")
Else
Call Combo11_Click
End If
End Sub

Private Sub Command9_Click()
Dim i As Integer
If Combo10.Text = "" Then
i = MsgBox("At first select a Template!", vbInformation, "Attorney Master Alert")
Else
Check1.Enabled = False
Text2.Locked = False
Text3.Text = Combo10.Text
'*****************
Dim ss, sss As String
Dim intFileNum As Integer
intFileNum = FreeFile
ss = ""
sss = ""
Open App.Path & "\Mail Templates\" & Text3.Text For Input As #intFileNum
Do
Line Input #intFileNum, ss
'sss = sss & ss
sss = sss & ss & vbCrLf
Loop Until EOF(intFileNum)
Close #intFileNum
'*******************
Text2.Text = sss
Text2.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim i, j As Integer
Form4.Caption = "Attorney Master [Mail Manager]-[Job ID: " & Form2.l(150).Caption & "]"
'j = Form2.Data3.Recordset.Fields.Count
'For i = 0 To j - 1
'Combo1.AddItem Form2.Data3.Recordset.Fields(i).Name
'Next i
'*************************
c(0).AddItem "Last Name"
c(0).AddItem "First Name"
c(0).AddItem "MI"
c(0).AddItem "SS#"
c(0).AddItem "ADD_Number"
c(0).AddItem "ADD_Street"
c(0).AddItem "ADD_City"
c(0).AddItem "ADD_County"
c(0).AddItem "ADD_State"
c(0).AddItem "ADD_ZipCode" '9  ---   11
c(0).AddItem "Home_Phone"
c(0).AddItem "Business_Phone"
c(0).AddItem "Age"         '12  ---   dt(1)
c(0).AddItem "Spouse_Last Name"  '13  ---   18
c(0).AddItem "Spouse_First Name"
c(0).AddItem "Spouse_Occupation"
c(0).AddItem "Spouse_Phone"   '16  ---   23
c(0).AddItem "Occupation_By Whom Employed"  '17  ----  24
c(0).AddItem "Occupation_Nature of Work"
c(0).AddItem "Occupation_How Long Employed"
c(0).AddItem "Occupation_Prior Employment History"
c(0).AddItem "Insurance_Liability"
c(0).AddItem "Insurance_PIP"
c(0).AddItem "Insurance_Medical"
c(0).AddItem "Insurance_Limits"
c(0).AddItem "Insurance_BI Adjusters"
c(0).AddItem "Insurance_BI Name"        '25  ----   33
c(0).AddItem "Insurance_BI Phone Number"
c(0).AddItem "Insurance_BI Claim#"
c(0).AddItem "Insurance_PD Adjusters"
c(0).AddItem "Insurance_PD Name"
c(0).AddItem "Insurance_PD Phone Number"
c(0).AddItem "Insurance_PD Claim#"
c(1).AddItem "Last Name"   '0 ----
c(1).AddItem "First Name"
c(1).AddItem "MI"
c(1).AddItem "SS#"
c(1).AddItem "ADD_Number"
c(1).AddItem "ADD_Street"
c(1).AddItem "ADD_City"
c(1).AddItem "ADD_County"
c(1).AddItem "ADD_State"
c(1).AddItem "ADD_ZipCode"
c(1).AddItem "Home_Phone"
c(1).AddItem "Business_Phone"
c(1).AddItem "Age"
c(1).AddItem "Spouse_Last Name"
c(1).AddItem "Spouse_First Name"
c(1).AddItem "Spouse_Occupation"
c(1).AddItem "Spouse_Phone"
c(1).AddItem "Occupation_By Whom Employed"
c(1).AddItem "Occupation_Nature of Work"
c(1).AddItem "Occupation_How Long Employed"
c(1).AddItem "Occupation_Prior Employment History"
c(1).AddItem "Insurance_Liability"
c(1).AddItem "Insurance_PIP"
c(1).AddItem "Insurance_Medical"
c(1).AddItem "Insurance_BI Adjusters"
c(1).AddItem "Insurance_BI Name"
c(1).AddItem "Insurance_BI Phone Number"
c(1).AddItem "Insurance_BI Claim#"
c(1).AddItem "Insurance_PD Adjusters"
c(1).AddItem "Insurance_PD Name"
c(1).AddItem "Insurance_PD Phone Number"
c(1).AddItem "Insurance_PD Claim#"
c(1).AddItem "Insurance_Limits"
c(2).AddItem "When_Date"    '0 ---
c(2).AddItem "When_Time"    '1 --- 88 89 90
c(2).AddItem "Location_City" '2 --- 91
c(2).AddItem "Location_County"
c(2).AddItem "Location_State"
c(2).AddItem "Location_Address"
c(2).AddItem "Location_Street"
c(2).AddItem "Location_Intersecting Street"
c(2).AddItem "Narrative Description"
c(2).AddItem "Diagram of Scene"
c(2).AddItem "Names of Investigating Agency"
c(2).AddItem "Name of Investigating Officer"
c(2).AddItem "Witnesses_Names"
c(2).AddItem "Witnesses_Addresses"
c(2).AddItem "Witnesses_Occupations"
c(2).AddItem "Witnesses_Phone Numbers"
c(2).AddItem "Witnesses_Subject Matter of Information"
c(3).AddItem "Involved_Names"
c(3).AddItem "Involved_Roadway surface"
c(3).AddItem "Involved_Roadway Lanes?"
c(3).AddItem "Involved_One or Two-way?"
c(3).AddItem "Involved_Interchange"
c(3).AddItem "Involved_Description of Road"
c(3).AddItem "Involved_Locality"
c(3).AddItem "Speed limit of roadways"
c(3).AddItem "Speeds of all vehicles prior to braking"
c(3).AddItem "Speeds of vehicles at impact"
c(3).AddItem "Plaintiff_trip begin?"
c(3).AddItem "Plaintiff_destination?"
c(3).AddItem "Plaintiff_Purpose?"
c(3).AddItem "Defendant_trip begin?"
c(3).AddItem "Defendant_destination?"
c(3).AddItem "Defendant_Purpose?"
c(4).AddItem "Plaintiff_Passenger Number"  '0 --- 118
c(4).AddItem "Plaintiff_Passenger Name"   '1 --- 119
c(4).AddItem "Plaintiff_Passenger Address"  '2 --- 120
c(4).AddItem "Plaintiff_Passenger Occupation"  '3 --- 121
c(4).AddItem "Plaintiff_Passenger Phone Number" '4 --- 122 123 124
c(4).AddItem "Plaintiff_Passenger Location?"    '5 --- 125
c(4).AddItem "Plaintiff_Passenger Injuries?"
c(4).AddItem "Plaintiff_Passenger Relationship"
c(4).AddItem "Defendant_Passenger Number"
c(4).AddItem "Defendant_Passenger Name"
c(4).AddItem "Defendant_Passenger Address"
c(4).AddItem "Defendant_Passenger Occupation"
c(4).AddItem "Defendant_Passenger Phone Number"
c(4).AddItem "Defendant_Passenger Location?"
c(4).AddItem "Defendant_Passenger Injuries?"
c(4).AddItem "Defendant_Passenger Relationship"
c(4).AddItem "What gave rise to notice of danger?"
c(4).AddItem "Relative position of all vehicles at that Time."
c(4).AddItem "Approximate time."
c(4).AddItem "Removal Plaintiff Vehicles"
c(4).AddItem "Removal Defendant Vehicles"
c(5).AddItem "What parts of vehicle collided?"
c(5).AddItem "Where did each vehicle end up?"
c(5).AddItem "How far apart were the vehicles?"
c(5).AddItem "Describe after impact:"
c(5).AddItem "Plaintiff's exterior damage"
c(5).AddItem "Plaintiff's cost of repair"
c(5).AddItem "Plaintiff's cost of repair"
c(5).AddItem "Defendant's exterior damage"
c(5).AddItem "Defendant's cost of repair"
c(5).AddItem "What happened to body inside?"
c(6).AddItem "When first seen by health care provider"  '0 --- 144
c(6).AddItem "Several days later when all injuries could be appreciated"
c(6).AddItem "Hospitals Name"   '2 --- 146
c(6).AddItem "Hospitals Street Name"  '3  --- 147
c(6).AddItem "Hospitals City"   '4  ----  148
c(6).AddItem "Hospitals County"  '5  ----  149
c(6).AddItem "Hospitals State"  '6  ----  150
c(6).AddItem "Hospitals Date Begin"  '7  ----  dt 4
c(6).AddItem "Hospitals Date End"  '8 ----  dt 5
c(6).AddItem "Doctors Name"         '9  ----  151
c(6).AddItem "Doctors Street Name"    '10  ----  152
c(6).AddItem "Doctors City"    '11 ----  153
c(6).AddItem "Doctors County"    '12 ----  154
c(6).AddItem "Doctors State"    '13 ----  155
c(6).AddItem "Doctors Date Begin"    '14 ----  dt 6
c(6).AddItem "Doctors Date End"    '15 ----  dt 7
c(6).AddItem "Therapists Name"      '16 ----  156
c(6).AddItem "Therapists Street Name"      '17 ----  157
c(6).AddItem "Therapists City"      '18 ----  158
c(6).AddItem "Therapists County"      '19 ----  159
c(6).AddItem "Therapists State"      '20 ----  160
c(6).AddItem "Therapists Date Begin"      '21 ----  dt 8
c(6).AddItem "Therapists Date End"      '22 ----  dt 9
c(7).AddItem "Dates missed"     '0 ----  161
c(7).AddItem "Rate of compensation"     '1 ----  162
c(7).AddItem "Reason for wage loss"     '2 ----  163
c(7).AddItem "Out-of-Pocket Physicians"     '3 ----  164
c(7).AddItem "Out-of-Pocket Hospitals"     '4 ----  165
c(7).AddItem "Out-of-Pocket Ambulance"
c(7).AddItem "Out-of-Pocket Drugs"
c(7).AddItem "Out-of-Pocket Crutches,Appliances"
c(7).AddItem "Out-of-Pocket Domestic help"
c(7).AddItem "Out-of-Pocket Auto repair"
c(7).AddItem "Out-of-Pocket Lost wages"
c(7).AddItem "Out-of-Pocket Other"
c(7).AddItem "Impression of Client"
c(7).AddItem "Settlement Offers"
c(7).AddItem "Total Medical specials"
'***************************
Call ShowFolderList(App.Path & "\Mail Templates")
Call ShowFolderList1(App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Txt")
'Call ShowFolderList2(App.Path & "\Jobs\" & Form2.l(150).Caption & "\Mails\Doc")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload PWait
Form2.Show
End Sub

Sub ShowFolderList(folderspec)
    Dim fs, f, f1, fc, s
    Combo10.Clear
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        Combo10.AddItem s
    Next
End Sub

Sub ShowFolderList1(folderspec)
    Dim fs, f, f1, fc, s
    Combo11.Clear
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
        Combo11.AddItem s
    Next
End Sub

Sub ShowFolderList2(folderspec)
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = f1.Name
    Next
End Sub

Private Sub from_database(combonum As Integer, combolistnum As Integer)
'MsgBox "Goto the select"
Select Case combonum
Case 0
If combolistnum <= 2 Then
'MsgBox Form2.t(combolistnum).Text
Text2.SelText = Form2.t(combolistnum).Text
End If
If combolistnum = 3 Then
'MsgBox Form2.t(combolistnum).Text & "-" & Form2.t(combolistnum + 1).Text & "-" & Form2.t(combolistnum + 2).Text
Text2.SelText = Form2.t(combolistnum).Text & "-" & Form2.t(combolistnum + 1).Text & "-" & Form2.t(combolistnum + 2).Text
End If
If combolistnum >= 4 And combolistnum <= 9 Then
'MsgBox Form2.t(combolistnum + 2).Text
Text2.SelText = Form2.t(combolistnum + 2).Text
End If
If combolistnum = 10 Then
Text2.SelText = Form2.t(combolistnum + 4).Text & "-" & Form2.t(combolistnum + 3).Text & "-" & Form2.t(combolistnum + 2).Text
End If
If combolistnum = 11 Then
Text2.SelText = Form2.t(combolistnum + 6).Text & "-" & Form2.t(combolistnum + 5).Text & "-" & Form2.t(combolistnum + 4).Text
End If
If combolistnum = 12 Then
Text2.SelText = Form2.dt(1).Value
End If
If combolistnum = 13 Then
Text2.SelText = Form2.t(combolistnum + 5).Text
End If
If combolistnum = 14 Then
Text2.SelText = Form2.t(combolistnum + 5).Text
End If
If combolistnum = 15 Then
Text2.SelText = Form2.t(combolistnum + 5).Text
End If
If combolistnum = 16 Then
Text2.SelText = Form2.t(combolistnum + 5).Text & "-" & Form2.t(combolistnum + 6).Text & "-" & Form2.t(combolistnum + 7).Text
End If
If combolistnum >= 17 And combolistnum <= 26 Then
Text2.SelText = Form2.t(combolistnum + 7).Text
End If
If combolistnum = 27 Then
Text2.SelText = Form2.t(36).Text & "-" & Form2.t(35).Text & "-" & Form2.t(34).Text
End If
If combolistnum = 28 Then
Text2.SelText = Form2.t(37).Text
End If
If combolistnum = 29 Then
Text2.SelText = Form2.t(38).Text
End If
If combolistnum = 30 Then
Text2.SelText = Form2.t(39).Text
End If
If combolistnum = 31 Then
Text2.SelText = Form2.t(42).Text & "-" & Form2.t(41).Text & "-" & Form2.t(40).Text
End If
If combolistnum = 32 Then
Text2.SelText = Form2.t(43).Text
End If
Case 1
If combolistnum <= 2 Then
Text2.SelText = Form2.t(combolistnum + 44).Text
End If
If combolistnum = 3 Then
Text2.SelText = Form2.t(combolistnum + 44).Text & "-" & Form2.t(combolistnum + 1 + 44).Text & "-" & Form2.t(combolistnum + 2 + 44).Text
End If
If combolistnum >= 4 And combolistnum <= 9 Then
Text2.SelText = Form2.t(combolistnum + 2 + 44).Text
End If
If combolistnum = 10 Then
Text2.SelText = Form2.t(combolistnum + 2 + 44).Text & "-" & Form2.t(combolistnum + 3 + 44).Text & "-" & Form2.t(combolistnum + 4 + 44).Text
End If
If combolistnum = 11 Then
Text2.SelText = Form2.t(combolistnum + 4 + 44).Text & "-" & Form2.t(combolistnum + 5 + 44).Text & "-" & Form2.t(combolistnum + 6 + 44).Text
End If
If combolistnum = 12 Then
Text2.SelText = Form2.dt(2).Value
End If
If combolistnum = 13 Then
Text2.SelText = Form2.t(combolistnum + 5 + 44).Text
End If
If combolistnum = 14 Then
Text2.SelText = Form2.t(combolistnum + 5 + 44).Text
End If
If combolistnum = 15 Then
Text2.SelText = Form2.t(combolistnum + 5 + 44).Text
End If
If combolistnum = 16 Then
Text2.SelText = Form2.t(combolistnum + 5 + 44).Text & "-" & Form2.t(combolistnum + 6 + 44).Text & "-" & Form2.t(combolistnum + 7 + 44).Text
End If
If combolistnum >= 17 And combolistnum <= 25 Then
Text2.SelText = Form2.t(combolistnum + 7 + 44).Text
End If
If combolistnum = 26 Then
Text2.SelText = Form2.t(34 + 44).Text & "-" & Form2.t(35 + 44).Text & "-" & Form2.t(36 + 44).Text
End If
If combolistnum = 27 Then
Text2.SelText = Form2.t(37 + 44).Text
End If
If combolistnum = 28 Then
Text2.SelText = Form2.t(38 + 44).Text
End If
If combolistnum = 29 Then
Text2.SelText = Form2.t(39 + 44).Text
End If
If combolistnum = 30 Then
Text2.SelText = Form2.t(40 + 44).Text & "-" & Form2.t(41 + 44).Text & "-" & Form2.t(42 + 44).Text
End If
If combolistnum = 31 Then
Text2.SelText = Form2.t(43 + 44).Text
End If
Case 2
If combolistnum = 0 Then
Text2.SelText = Form2.dt(3).Value
End If
If combolistnum = 1 Then
Text2.SelText = Form2.t(88).Text & ":" & Form2.t(89).Text & ":" & Form2.t(90).Text & Form2.Label13.Caption
End If
If combolistnum > 1 Then
Text2.SelText = Form2.t(89 + combolostnum).Text
End If
Case 3
Text2.SelText = Form2.t(105 + combolostnum).Text
Case 4
If combolistnum < 4 Then
Text2.SelText = Form2.t(118 + combolostnum).Text
End If
If combolistnum = 4 Then
Text2.SelText = Form2.t(122).Text & "-" & Form2.t(123).Text & "-" & Form2.t(124).Text
End If
If combolistnum > 4 Then
Text2.SelText = Form2.t(120 + combolistnum).Text
End If
Case 5
Text2.SelText = Form2.t(134 + combolostnum).Text
Case 6
If combolistnum <= 6 Then
Text2.SelText = Form2.t(144 + combolistnum).Text
End If
If combolistnum = 7 Then
Text2.SelText = Form2.dt(4).Value
End If
If combolistnum = 8 Then
Text2.SelText = Form2.dt(5).Value
End If
If combolistnum >= 9 And combolistnum <= 13 Then
Text2.SelText = Form2.t(142 + combolistnum).Text
End If
If combolistnum = 14 Then
Text2.SelText = Form2.dt(6).Value
End If
If combolistnum = 15 Then
Text2.SelText = Form2.dt(7).Value
End If
If combolistnum >= 16 And combolistnum <= 20 Then
Text2.SelText = Form2.t(140 + combolistnum).Text
End If
If combolistnum = 21 Then
Text2.SelText = Form2.dt(8).Value
End If
If combolistnum = 22 Then
Text2.SelText = Form2.dt(9).Value
End If
Case 7
Text2.SelText = Form2.t(161 + combolostnum).Text
End Select
End Sub
