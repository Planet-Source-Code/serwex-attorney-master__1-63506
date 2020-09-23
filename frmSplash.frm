VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6105
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   3045
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   210
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Timer t1 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   4200
         Top             =   210
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   165
         Left            =   0
         ScaleHeight     =   165
         ScaleWidth      =   3735
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   3735
         Begin VB.ListBox lstFoundFiles 
            Height          =   255
            Left            =   210
            TabIndex        =   19
            Top             =   525
            Width           =   1485
         End
         Begin VB.DirListBox dirList 
            Height          =   315
            Left            =   2040
            TabIndex        =   18
            Top             =   525
            Width           =   1575
         End
         Begin VB.FileListBox filList 
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   105
            Width           =   1815
         End
         Begin VB.TextBox txtSearchSpec 
            Height          =   285
            Left            =   2040
            TabIndex        =   16
            Text            =   "Iexplore.exe"
            Top             =   120
            Width           =   1575
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&About..."
         Height          =   330
         Left            =   4830
         TabIndex        =   14
         Top             =   2310
         Width           =   1170
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Visit our site!"
         Height          =   330
         Left            =   4830
         TabIndex        =   13
         ToolTipText     =   "www.4LawSupport.com"
         Top             =   2730
         Width           =   1170
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Quit"
         Height          =   435
         Left            =   4830
         TabIndex        =   10
         Top             =   3255
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Users"
         Height          =   435
         Left            =   2310
         TabIndex        =   8
         Top             =   3255
         Width           =   1170
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Superviser"
         Height          =   435
         Left            =   3570
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3255
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00000000&
         Height          =   750
         Left            =   1575
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   2310
         Width           =   3165
      End
      Begin VB.Frame Frame2 
         Height          =   120
         Left            =   105
         TabIndex        =   4
         Top             =   3045
         Width           =   5790
      End
      Begin MSComctlLib.ProgressBar pb 
         Height          =   225
         Left            =   2310
         TabIndex        =   20
         Top             =   3780
         Visible         =   0   'False
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
         Max             =   20
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   1515
         Stretch         =   -1  'True
         Top             =   765
         Width           =   3660
      End
      Begin VB.Image Pic 
         Height          =   540
         Left            =   4725
         Stretch         =   -1  'True
         Top             =   210
         Width           =   1275
      End
      Begin VB.Label l 
         DataField       =   "Product_Serial"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   2
         Left            =   315
         TabIndex        =   23
         Top             =   1260
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label l 
         DataField       =   "Company_Name"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   1
         Left            =   315
         TabIndex        =   22
         Top             =   945
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label l 
         DataField       =   "Superviser_Name"
         DataSource      =   "Data1"
         Height          =   225
         Index           =   0
         Left            =   315
         TabIndex        =   21
         Top             =   1995
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line2 
         Visible         =   0   'False
         X1              =   4305
         X2              =   5040
         Y1              =   1365
         Y2              =   1365
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         Visible         =   0   'False
         X1              =   1575
         X2              =   5145
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label Label6 
         Caption         =   "Version:  1.0.0"
         Height          =   225
         Left            =   1575
         TabIndex        =   12
         Top             =   1470
         Width           =   1485
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Login as user"
         Height          =   225
         Left            =   2310
         TabIndex        =   11
         Top             =   3780
         Width           =   3690
      End
      Begin VB.Label Label5 
         Caption         =   "Welcome to:"
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
         Left            =   1575
         TabIndex        =   9
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label Label4 
         Caption         =   "Copyright 2000-2001 MixofTix Developers Network"
         Height          =   225
         Left            =   1575
         TabIndex        =   6
         Top             =   1680
         Width           =   3795
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
         Left            =   1575
         TabIndex        =   5
         Top             =   630
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.Label Label2 
         Caption         =   "This program is licensed to:"
         Height          =   225
         Left            =   1575
         TabIndex        =   2
         Top             =   1995
         Width           =   2010
      End
      Begin VB.Image Logo 
         BorderStyle     =   1  'Fixed Single
         Height          =   2700
         Left            =   210
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1185
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning: This product is protected by copyright law and international treaties."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   210
         TabIndex        =   1
         Top             =   3255
         Width           =   2025
      End
   End
End
Attribute VB_Name = "Splash"
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
Dim iepath As String
Dim r As Double

Private Sub Command1_Click()
    Me.Hide
    Login.Show
End Sub

Private Sub Command2_Click()
    Me.Hide
    Form1.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
pb.Visible = True
pb.Value = 0
Label3.Visible = False
'**************
Dim i As Integer
'Dim DirCount1 As Integer
Dim f As Double
'Dim path_IE As String
On Error GoTo IE_not_found
'**********************
Dim FirstPath As String, DirCount As Integer, NumFiles As Integer
Dim result As Integer
If (UCase(Dir("C:\Program Files\Internet Explorer\Iexplore.exe"))) = UCase("Iexplore.exe") Then
f = Shell("C:\Program Files\Internet Explorer\Iexplore.exe www.MixofTix.net", vbMaximizedFocus)
Else
t1.Enabled = True
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
End If
pb.Value = 0
t1.Enabled = False
pb.Visible = False
Label3.Visible = True
Exit Sub
IE_not_found:
i = MsgBox("Internet Explorer not found.Goto site ,www.MixofTix.net,by your self.", vbInformation, "IE Not Found")
'**************
pb.Value = 0
t1.Enabled = False
pb.Visible = False
Label3.Visible = True
'End of file
End Sub

Private Sub Command5_Click()
About.Show

End Sub

Private Sub Form_Initialize()
Pic.Picture = LoadPicture(App.Path & "\logo.gif")
Logo.Picture = LoadPicture(App.Path & "\logo.jpg")
Image1.Picture = LoadPicture(App.Path & "\name.jpg")
Command2.Default = True
End Sub

Private Sub Form_Load()
  Data1.DatabaseName = App.Path & "\Shabshab.mdb"
  Data1.RecordSource = "Fixed_IDS"
  Data1.Refresh
  Data1.Recordset.MoveFirst
  Text1.Text = l(0).Caption & vbCrLf & l(1).Caption & vbCrLf & "Product Serial Number: " & l(2).Caption
End Sub



Private Sub Logo_Click()
MsgBox Format(Now, "d-mmmm h:nn:ss")
End Sub

Private Sub Text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command2.SetFocus
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

Private Sub t1_Timer()
If pb.Value < pb.Max Then
pb.Value = pb.Value + 1
Else
pb.Value = 1
End If
End Sub

Private Sub txtSearchSpec_Change()
    ' Update file list box if user changes pattern.
    filList.Pattern = txtSearchSpec.Text
End Sub


