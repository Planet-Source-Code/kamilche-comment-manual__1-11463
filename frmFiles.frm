VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFiles 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Comment Manual"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4710
      Left            =   75
      TabIndex        =   1
      Top             =   90
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   8308
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabMaxWidth     =   2999
      TabCaption(0)   =   "Files To Print"
      TabPicture(0)   =   "frmFiles.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "picParent(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Text Preview"
      TabPicture(1)   =   "frmFiles.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picParent(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Formatted Preview"
      TabPicture(2)   =   "frmFiles.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "picParent(2)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdPrint"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin VB.CommandButton cmdPrint 
         Height          =   285
         Left            =   5235
         Picture         =   "frmFiles.frx":0054
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         Width           =   315
      End
      Begin VB.PictureBox picParent 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3870
         Index           =   1
         Left            =   -74655
         ScaleHeight     =   3870
         ScaleWidth      =   6960
         TabIndex        =   7
         Top             =   555
         Visible         =   0   'False
         Width           =   6960
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   750
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   1323
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   3
            RightMargin     =   12000
            TextRTF         =   $"frmFiles.frx":0156
         End
      End
      Begin VB.PictureBox picParent 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   3870
         Index           =   2
         Left            =   195
         ScaleHeight     =   3870
         ScaleWidth      =   6960
         TabIndex        =   6
         Top             =   540
         Visible         =   0   'False
         Width           =   6960
      End
      Begin VB.PictureBox picParent 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3870
         Index           =   0
         Left            =   -74700
         ScaleHeight     =   3870
         ScaleWidth      =   6960
         TabIndex        =   2
         Top             =   555
         Visible         =   0   'False
         Width           =   6960
         Begin VB.FileListBox File1 
            Height          =   2820
            Left            =   2025
            MultiSelect     =   2  'Extended
            Pattern         =   "*.frm;*.cls;*.bas"
            TabIndex        =   5
            Top             =   0
            Width           =   2700
         End
         Begin VB.DirListBox Dir1 
            Height          =   2340
            Left            =   0
            TabIndex        =   4
            Top             =   345
            Width           =   2040
         End
         Begin VB.DriveListBox Drive1 
            Height          =   315
            Left            =   0
            TabIndex        =   3
            Top             =   15
            Width           =   2040
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4830
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPrint_Click()
    'Print all the selected files.
    Dim i As Long
    MyPrinter.SetDestination Printer
    For i = 1 To UBound(Files, 1)
        MyPrinter.PrintFile Files(i)
    Next i
End Sub

Private Sub Drive1_Change()
    'User changed the drive - update the directory listing to reflect new choice.
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Dir1_Change()
    'User changed the directory - update the files listing to reflect new  choice.
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    'User selected new files - update the file count and file array to
    'reflect the new choices.
    RefreshCount
End Sub

Private Sub Form_Load()
    'Load 'last pathname' setting, set directory defaults, choose first tab.
    Dim Pathname As String, i As Long
    ReDim Files(0 To 0)
    Pathname = GetSetting(App.Title, "Preferences", "Pathname", App.Path)
    Drive1.Drive = Pathname
    Dir1.Path = Pathname
    For i = 0 To picParent.Count - 1
        picParent(i).BackColor = &HC0C0C0
    Next i
    picParent(2).BackColor = vbWhite
    SSTab1.Tab = 0
    SSTab1_Click 0
    RefreshCount
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Save 'last pathname' setting
    Dim Pathname As String
    Pathname = Dir1.Path
    SaveSetting App.Title, "Preferences", "Pathname", Pathname
    ShutDown
End Sub

Private Sub Form_Resize()
    'Each tab has its own 'parent' picturebox, which contains all the
    'controls for that tab. Resize the tabs, and move all the 'parent'
    'pictureboxes to fit the new screen size.
    Dim i As Long, max As Long
    On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        SSTab1.Move 50, 50, ScaleWidth - 100, ScaleHeight - StatusBar1.Height - 100
        For i = 0 To picParent.Count - 1
            picParent(i).Move 150, 500, SSTab1.Width - 300, SSTab1.Height - 650
        Next i
        If SSTab1.Tab = 2 Then
            Preview
        End If
    End If
End Sub

Private Sub picParent_Resize(Index As Integer)
    'Rearrange all the controls to fit the new 'parent picturebox' size.
    With picParent(Index)
        If Index = 0 Then
            Drive1.Move 0, 0, .ScaleWidth
            Dir1.Move 0, Drive1.Height, .ScaleWidth / 2, .ScaleHeight - Drive1.Height
            File1.Move Dir1.Width, Dir1.Top, .ScaleWidth - Dir1.Width, Dir1.Height
        ElseIf Index = 1 Then
            RichTextBox1.Move 0, 0, .ScaleWidth, .ScaleHeight
        ElseIf Index = 2 Then
            picParent(Index).Cls
        End If
    End With
End Sub

Private Sub Status(ByVal s As String)
    'Display text in the status bar at the bottom of the screen.
    StatusBar1.SimpleText = s
    StatusBar1.Refresh
End Sub

Private Sub RefreshCount()
    'Displays the count of selected files, and
    'refreshes the 'files' array with that count.
    Dim i As Long, max As Long, ctr As Long
    max = File1.ListCount - 1
    ctr = 0
    ReDim Files(0 To 0)
    For i = 0 To max
        If File1.Selected(i) = True Then
            ctr = ctr + 1
            ReDim Preserve Files(0 To ctr)
            Files(ctr) = File1.Path & "\" & File1.List(i)
        End If
    Next i
    If ctr = 0 Then
        Status File1.ListCount & " files."
    ElseIf ctr = 1 Then
        Status ctr & " file selected."
    Else
        Status ctr & " files selected."
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    'Display the parent picturebox appropriate for this tab,
    'and hide all the other ones. If the 'Preview' tab
    'was selected, perform the preview as well.
    Dim i As Long
    For i = 0 To picParent.Count - 1
        picParent(i).Visible = False
    Next i
    picParent(SSTab1.Tab).Visible = True
    
End Sub

Private Sub Preview()
    'Preview the comment manual.
    Dim i As Long, max As Long
    
    If CheckForFiles = False Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    
    MyPrinter.PrintPreview
End Sub

Private Sub TextPreview()
    'Preview the comment manual.
    Dim i As Long, max As Long
    
    If CheckForFiles = False Then
        SSTab1.Tab = 0
        Exit Sub
    End If
    
    'Load the file into the textbox
    RichTextBox1.Text = LoadFile(Files(1))
End Sub

Private Function CheckForFiles() As Boolean
    'Checks to see if there's at least one file selected.
    'Returns true if so, false otherwise.
    Dim max As Long
    max = UBound(Files, 1)
    If max = 0 Then
        MsgBox "There are no files selected! Please go to the 'Files' tab and choose which files to print.", vbCritical
        CheckForFiles = False
    Else
        CheckForFiles = True
    End If
End Function

Private Sub SSTab1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Perform the action associated with the current tab (used by text and formatted preview tabs)
    If SSTab1.Tab = 1 Then
        TextPreview
    ElseIf SSTab1.Tab = 2 Then
        Preview
    End If
End Sub
