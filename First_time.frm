VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form First_time 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6585
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "First_time.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   6585
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar sb 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   15
      Top             =   6855
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8996
            Text            =   "Status: Fill the Name & Company to continue..."
            TextSave        =   "Status: Fill the Name & Company to continue..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Text            =   "Please Wait..."
            TextSave        =   "Please Wait..."
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
   Begin VB.Frame Frame1 
      Height          =   6510
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   6165
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1785
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4935
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   4410
         Width           =   2010
      End
      Begin VB.TextBox Text4 
         DataField       =   "Superviser_Password"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1680
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   3990
         Width           =   2010
      End
      Begin VB.Timer t1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   3570
         Top             =   5565
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   5985
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Max             =   20
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   210
         ScaleHeight     =   165
         ScaleWidth      =   3735
         TabIndex        =   18
         Top             =   6195
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox txtSearchSpec 
            Height          =   285
            Left            =   2040
            TabIndex        =   22
            Text            =   "Iexplore.exe"
            Top             =   120
            Width           =   1575
         End
         Begin VB.FileListBox filList 
            Height          =   285
            Left            =   120
            TabIndex        =   21
            Top             =   105
            Width           =   1815
         End
         Begin VB.DirListBox dirList 
            Height          =   315
            Left            =   2040
            TabIndex        =   20
            Top             =   525
            Width           =   1575
         End
         Begin VB.ListBox lstFoundFiles 
            Height          =   255
            Left            =   210
            TabIndex        =   19
            Top             =   525
            Width           =   1485
         End
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   540
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   17
         Text            =   "First_time.frx":000C
         Top             =   5355
         Width           =   3690
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Visit our site!"
         Height          =   330
         Left            =   4200
         TabIndex        =   7
         ToolTipText     =   "www.4LawSupport.com"
         Top             =   3570
         Width           =   1800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "E&xit setup"
         Height          =   330
         Left            =   4200
         TabIndex        =   6
         Top             =   3150
         Width           =   1800
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Start"
         Height          =   330
         Left            =   4200
         TabIndex        =   5
         Top             =   2730
         Width           =   1800
      End
      Begin VB.TextBox Text2 
         DataField       =   "Company_Name"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         TabIndex        =   2
         Top             =   3570
         Width           =   2010
      End
      Begin VB.TextBox Text1 
         DataField       =   "Superviser_Name"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   3150
         Width           =   2010
      End
      Begin VB.Image Image1 
         Height          =   750
         Left            =   1995
         Top             =   945
         Width           =   4080
      End
      Begin VB.Image Pic 
         Height          =   540
         Left            =   4620
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label13 
         DataField       =   "Last_Job_ID"
         DataSource      =   "Data1"
         Height          =   225
         Left            =   4305
         TabIndex        =   27
         Top             =   1995
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label12 
         Caption         =   "Your product serial number:      "
         Height          =   225
         Left            =   210
         TabIndex        =   26
         Top             =   2835
         Width           =   2010
      End
      Begin VB.Label Label11 
         Caption         =   "Confirm Password:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   210
         TabIndex        =   25
         Top             =   4515
         Width           =   1380
      End
      Begin VB.Label Label10 
         Caption         =   "Password:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   210
         TabIndex        =   24
         Top             =   4095
         Width           =   960
      End
      Begin VB.Line Line3 
         X1              =   210
         X2              =   1470
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Label Label9 
         Caption         =   "Description:"
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
         TabIndex        =   16
         Top             =   5040
         Width           =   1590
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Product_Serial"
         DataSource      =   "Data1"
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2520
         TabIndex        =   14
         Top             =   2730
         Width           =   1170
      End
      Begin VB.Label Label7 
         Caption         =   "Label7"
         Height          =   120
         Left            =   210
         TabIndex        =   13
         Top             =   2835
         Width           =   15
      End
      Begin VB.Label Label3 
         Caption         =   "Company:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   210
         TabIndex        =   12
         Top             =   3675
         Width           =   960
      End
      Begin VB.Label Label2 
         Caption         =   "Name:"
         Enabled         =   0   'False
         Height          =   225
         Left            =   210
         TabIndex        =   11
         Top             =   3255
         Width           =   855
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   3150
         X2              =   3885
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   420
         X2              =   4200
         Y1              =   5775
         Y2              =   5775
      End
      Begin VB.Image logo1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2280
         Left            =   4200
         Stretch         =   -1  'True
         ToolTipText     =   "About Attorney Master..."
         Top             =   4095
         Width           =   1815
      End
      Begin VB.Image Logo 
         BorderStyle     =   1  'Fixed Single
         Height          =   2280
         Left            =   105
         Stretch         =   -1  'True
         ToolTipText     =   "About Attorney Master..."
         Top             =   210
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Attorney Master"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   420
         TabIndex        =   10
         Top             =   5145
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label Label4 
         Caption         =   "Copyright 2000-2001 MixofTix Developer Network"
         Height          =   225
         Left            =   2100
         TabIndex        =   9
         Top             =   2310
         Width           =   3660
      End
      Begin VB.Label Label6 
         Caption         =   "Version:  1.0.0"
         Height          =   225
         Left            =   2100
         TabIndex        =   8
         Top             =   2100
         Width           =   1485
      End
   End
End
Attribute VB_Name = "First_time"
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
Dim ok_ As String
Dim iepath As String
Dim r As Double

Private Sub Command1_Click()
On Error GoTo err:
If Command1.Caption = "&Start" Then
Command1.Caption = "&Next >"
'MsgBox Data1.DatabaseName
Data1.DatabaseName = App.Path & "\Shabshab.mdb"
Data1.RecordSource = "Fixed_IDS"
Data1.Refresh
'MsgBox Data1.RecordSource
'MsgBox Data1.RecordSource
'MsgBox App.Path
'MsgBox Data1.DatabaseName
'MsgBox Data1.Recordset.RecordCount
'Data1.Recordset.MoveFirst
Data1.Recordset.AddNew
Label13.Caption = "111100000"
Label8.Caption = "1111-111111"
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text1.SetFocus
Label2.Enabled = True
Label3.Enabled = True
Label10.Enabled = True
Label11.Enabled = True
Exit Sub
End If
If Command1.Caption = "&Finish" Then
Unload Me
Splash.Show
Exit Sub
End If
Command3.Enabled = False
If Text1.Text <> "" And Text2.Text <> "" _
And Text4.Text <> "" And Text5.Text <> "" _
And Text4.Text = Text5.Text _
Then
Data1.Recordset.Update
Data1.Recordset.MoveFirst
sb.Panels(1).Text = "Status: Initialize the system for the first time..."
Command1.Enabled = False
MkDir (App.Path & "\Jobs")
MkDir (App.Path & "\New Users")
MkDir (App.Path & "\Transfered Jobs")
MkDir (App.Path & "\Mail Templates")
MkDir (App.Path & "\Users")
MkDir (App.Path & "\Users\" & Text1.Text)
MkDir (App.Path & "\Users\" & Text1.Text & "\Active Jobs")
MkDir (App.Path & "\Users\" & Text1.Text & "\Inactive Jobs")
MkDir (App.Path & "\Users\" & Text1.Text & "\Messages")
MkDir (App.Path & "\Users\" & Text1.Text & "\Modified")
MkDir (App.Path & "\Users\" & Text1.Text & "\Notes")
MkDir (App.Path & "\Users\" & Text1.Text & "\Transfered Job")
Command1.Enabled = True
Command1.Caption = "&Finish"
Command2.Enabled = False
Text1.Enabled = False
Text2.Enabled = False
Text4.Enabled = False
Text5.Enabled = False
sb.Panels(1).Text = "Status: Initialize the system Finished."
sb.Panels(2).Text = "Job Complete."
Else
Text5.Text = ""
Text4.Text = ""
Text4.SetFocus
sb.Panels(1).Text = "Status: Fill the Name & Company to continue..."
End If
Exit Sub
err:
MsgBox (err.Description)
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Dim i As Integer
'Dim DirCount1 As Integer
Dim f As Double
Dim remm As String
On Error GoTo IE_not_found
'**********************
remm = sb.Panels(1).Text
sb.Panels(1).Text = "Status: You are planning to visit our site..."
If (UCase(Dir("C:\Program Files\Internet Explorer\Iexplore.exe"))) = UCase("Iexplore.exe") Then
f = Shell("C:\Program Files\Internet Explorer\Iexplore.exe www.MixofTix.net", vbMaximizedFocus)
Else
t1.Enabled = True
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
    dirList.Path = "C:"
    dirList.Path = "\"
    ResetSearch
    filList.Pattern = txtSearchSpec.Text
    FirstPath = dirList.Path
    DirCount = dirList.ListCount
    NumFiles = 0                       ' Reset found files indicator.
    r = 0
    result = DirDiver(FirstPath, DirCount, "")
    filList.Path = dirList.Path
'**********************
'dirList.Path = "C:"
'ChDir "\"
'DirCount1 = dirList.ListCount
'filList.Pattern = "Iexplore.exe"
'path_IE = find_IE("Iexplore.exe", DirCount1, "")
'"C:\Program Files\Internet Explorer\Iexplore.exe www.MixofTix.net"
If iepath = "" Then GoTo IE_not_found
iepath = iepath & " www.MixofTix.net"
f = Shell(iepath, vbMaximizedFocus)
t1.Enabled = False
pb.Value = 0
End If
sb.Panels(1).Text = remm
Exit Sub
IE_not_found:
i = MsgBox("Internet Explorer not found.Goto site ,www.MixofTix.net,by your self!", vbInformation, "IE Not Found")
t1.Enabled = False
pb.Value = 0
sb.Panels(1).Text = remm
End Sub



Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer
'  Recursively search directories from NewPath down...
'  NewPath is searched on this recursion.
'  BackUp is origin of this recursion.
'  DirCount is number of subdirectories in this directory.
Static FirstErr As Integer
Dim ii As Double
Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
Dim OldPath As String, ThePath As String, entry As String
Dim retval As Integer
    DirDiver = False            ' Set to True if there is an error.
    retval = DoEvents()         ' Check for events (for instance, if the user chooses Cancel).
    On Local Error GoTo DirDriverHandler
    DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            r = r + 1
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            lstFoundFiles.AddItem entry
            iepath = entry
            For ii = 1 To r
            Exit Function
            Next ii
            'lblCount.Caption = Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    Exit Function
DirDriverHandler:
    If err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If
End Function

Private Sub DirList_Change()
    ' Update the file list box to synchronize with the directory list box.
    filList.Path = dirList.Path
End Sub

Private Sub DirList_LostFocus()
    dirList.Path = dirList.List(dirList.ListIndex)
End Sub
Private Sub ResetSearch()
    ' Reinitialize before starting a new search.
    lstFoundFiles.Clear
'    lblCount.Caption = 0
'    Picture2.Visible = False
'    cmdSearch.Caption = "&Search"
'    Picture1.Visible = True
'    dirList.Path = CurDir: drvList.Drive = dirList.Path ' Reset the path.
End Sub

Private Sub Form_Load()
On Error GoTo err:
ok_ = "error"
Call ShowsubFolderList11(App.Path)
If ok_ = "ll" Then
Pic.Picture = LoadPicture(App.Path & "\logo.gif")
Logo.Picture = LoadPicture(App.Path & "\logo.jpg")
logo1.Picture = LoadPicture(App.Path & "\file_s.jpg")
Image1.Picture = LoadPicture(App.Path & "\name.jpg")
End If
Exit Sub
err:
MsgBox (err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'If ok_ = "ll" And Command1.Caption <> "&Start" Then
'Data1.Recordset.Close
'End If
End Sub

Private Sub t1_Timer()
Dim i As Integer
If pb.Value < pb.Max Then
pb.Value = pb.Value + 1
Else
pb.Value = 1
End If
End Sub
Private Sub Text2_GotFocus()
    Text2.SelStart = 0
    Text2.SelLength = Len(Text2.Text)
End Sub
Private Sub Text3_GotFocus()
    Text5.SelStart = 0
    Text5.SelLength = Len(Text5.Text)
End Sub
Private Sub Text4_GotFocus()
    Text4.SelStart = 0
    Text4.SelLength = Len(Text4.Text)
End Sub


Private Sub Text1_GotFocus()
    Text1.SelStart = 0
    Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub txtSearchSpec_Change()
    ' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.Text
End Sub

Sub ShowsubFolderList11(folderspec)
    Dim j, U As String
    Dim fs, f, f1, s, sf
On Error GoTo err:
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set sf = f.SubFolders
  j = "Error"
  U = "Error"
    For Each f1 In sf
        s = f1.Name
        If s = "Jobs" Then j = "OK"
        If s = "Users" Then U = "OK"
    Next
If j = "OK" And U = "OK" Then
ok_ = "kk"
Unload Me
Splash.Show
Else
ok_ = "ll"
End If
Exit Sub
err:
MsgBox (err.Description)
End Sub

