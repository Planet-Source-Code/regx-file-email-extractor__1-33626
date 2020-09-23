VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DGS Email Extractor"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11115
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11115
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picstatus1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00808080&
      ForeColor       =   &H8000000E&
      Height          =   300
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   11055
      TabIndex        =   23
      Top             =   6810
      Width           =   11115
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   6825
      TabIndex        =   22
      Top             =   45
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   10398
      View            =   3
      SortOrder       =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Email-Address"
         Object.Tag             =   "Email-Address"
         Text            =   "Email-Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Source"
         Object.Tag             =   "Source"
         Text            =   "Source"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open File"
      Filter          =   "Text Files(*.txt,*.rtf)|*.txt;*.rtf|MS Office(*.doc,*.mdb,*.xls) |*.doc;*.mdb;*.xls|All Files(*.*)|*.*"
   End
   Begin RichTextLib.RichTextBox RTF1 
      Height          =   4620
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8149
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0442
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5925
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10451
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabMaxWidth     =   3881
      MouseIcon       =   "Form1.frx":0523
      TabCaption(0)   =   "Extract From File"
      TabPicture(0)   =   "Form1.frx":053F
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblscanfiletarget"
      Tab(0).Control(1)=   "cmdSelectFile"
      Tab(0).Control(2)=   "cmdScanFile"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Extract From Folder"
      TabPicture(1)   =   "Form1.frx":0A81
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblscanfoldertarget"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdRemoveExt"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdAddExt"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txtext"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lstExt"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdScanFolder"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmbRecursionlvl"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdStopScanningFolder"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cmdSelectFolder"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CommandButton cmdScanFile 
         BackColor       =   &H00008000&
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -69360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   840
         Width           =   975
      End
      Begin VB.CommandButton cmdSelectFile 
         BackColor       =   &H0000FFFF&
         Caption         =   "Select File"
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
         Left            =   -74880
         TabIndex        =   18
         Top             =   480
         Width           =   1395
      End
      Begin VB.CommandButton cmdSelectFolder 
         BackColor       =   &H0000FFFF&
         Caption         =   "Select Folder"
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
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   1395
      End
      Begin VB.CommandButton cmdStopScanningFolder 
         BackColor       =   &H000000C0&
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   855
         Width           =   735
      End
      Begin VB.ComboBox cmbRecursionlvl 
         Height          =   315
         ItemData        =   "Form1.frx":0FC3
         Left            =   1560
         List            =   "Form1.frx":0FD9
         TabIndex        =   13
         Text            =   "0 - unlimited"
         Top             =   855
         Width           =   1200
      End
      Begin VB.CommandButton cmdScanFolder 
         BackColor       =   &H00008000&
         Caption         =   "Scan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   855
         Width           =   855
      End
      Begin VB.ListBox lstExt 
         Height          =   3765
         ItemData        =   "Form1.frx":0FFB
         Left            =   5280
         List            =   "Form1.frx":1002
         MultiSelect     =   2  'Extended
         TabIndex        =   10
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtext 
         Height          =   285
         Left            =   5280
         TabIndex        =   9
         Top             =   5160
         Width           =   795
      End
      Begin VB.CommandButton cmdAddExt 
         Caption         =   "Add"
         Height          =   255
         Left            =   6120
         TabIndex        =   8
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton cmdRemoveExt 
         Caption         =   "Remove"
         Height          =   255
         Left            =   5280
         TabIndex        =   7
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label lblscanfiletarget 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   -73440
         TabIndex        =   19
         Top             =   480
         Width           =   5070
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recursion Level"
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   855
         Width           =   1395
      End
      Begin VB.Label lblscanfoldertarget 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1545
         TabIndex        =   16
         Top             =   480
         Width           =   5070
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Extensions"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   5280
         TabIndex        =   11
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.CommandButton cmdremoveselectedemails 
      Caption         =   "Remove Selected"
      Height          =   360
      Left            =   6840
      TabIndex        =   5
      Top             =   6360
      Width           =   2055
   End
   Begin VB.CommandButton cmdsavelist 
      Caption         =   "Save emails to File"
      Height          =   360
      Left            =   9120
      TabIndex        =   4
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearEmailList 
      Caption         =   "Clear List"
      Height          =   255
      Left            =   10080
      TabIndex        =   2
      Top             =   6000
      Width           =   975
   End
   Begin VB.PictureBox picstatus2 
      Align           =   2  'Align Bottom
      BackColor       =   &H00404040&
      ForeColor       =   &H8000000E&
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   11055
      TabIndex        =   1
      Top             =   7110
      Width           =   11115
   End
   Begin VB.Label lblscancount 
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
      Left            =   120
      TabIndex        =   21
      Top             =   6000
      Width           =   5055
   End
   Begin VB.Label lbllistcount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6840
      TabIndex        =   3
      Top             =   6000
      Width           =   3075
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' VeryFast File-Email Extractor
' Copyright 2002 DGS
'Written by Gary Varnell
'=============================================
'Needs reference to:
'Microsoft Scripting Runtime
'Microsoft VBScript Regular Expressions 5.5
'download at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
'=============================================
Option Explicit
Dim extensions As Dictionary
Dim regx1 As RegExp
Dim Matches As MatchCollection
Dim Match As Match
Public WithEvents DGSDirScan1 As DGSDirScanner
Attribute DGSDirScan1.VB_VarHelpID = -1
Dim scancount As Long


Private Sub cmdScanFolder_Click()
' make sure user selected a folder
If Me.lblscanfoldertarget & "" = "" Then
    MsgBox "Please select a folder to scan", vbOKOnly, "Nothing to do"
    Exit Sub
End If

' make sure folder exist
Dim fs As FileSystemObject
Set fs = New FileSystemObject
If fs.FolderExists(Me.lblscanfoldertarget) = False Then
    MsgBox "The folder you selected does not exist!", vbCritical, "Error"
    Exit Sub
End If

' create a dictionary object and add all the extensions to be scanned
Set extensions = New Dictionary
Dim a As Long
extensions.RemoveAll
For a = 0 To lstExt.ListCount - 1
    extensions.Add lstExt.List(a), lstExt.List(a)
Next

scancount = 0 = 0
' Start recursive directory scan
' Fires the new folder and new file event (below) for every file/folder
DGSDirScan1.Scan Me.lblscanfoldertarget
status1 "All Done =)"
status2 ""
Beep

End Sub

Private Sub DGSDirScan1_newdir(d As Scripting.Folder)
status1 "Scanning Dir " & d.path
End Sub

Private Sub DGSDirScan1_newfile(f As Scripting.IFile)
status2 f.Name
Dim ext As String
ext = Right(f.path, 4)
If extensions.Exists(ext) = True Then
    status2 "Opening File " & f.path
    Me.RTF1.LoadFile f.path
    ExtractEmail f.path
End If
End Sub

Private Sub cmdremoveext_Click()
On Error Resume Next
If Me.lstExt.SelCount = 0 Then
    MsgBox "Please select an extension to remove", vbOKOnly, "Nothing to do!"
End If
Dim x As Long
x = 0
While lstExt.SelCount > 0
If lstExt.Selected(x) = True Then
    lstExt.RemoveItem x
Else
    x = x + 1
End If
Wend

End Sub

Private Sub cmdremoveselectedemails_Click()
On Error Resume Next
Dim x As Long
x = 1
While x < ListView1.ListItems.Count + 1
Debug.Print x
    If ListView1.ListItems.Item(x).Selected = True Then
        ListView1.ListItems.Remove x
    Else
        x = x + 1
    End If
Wend
End Sub

Private Sub cmdsavelist_Click()
Form2.Show vbModal
End Sub

Private Sub cmbRecursionlvl_Change()
Me.DGSDirScan1.Scandepth = cmbRecursionlvl
End Sub

Private Sub cmbRecursionlvl_Click()
Me.DGSDirScan1.Scandepth = Mid(cmbRecursionlvl, 1, 1)
End Sub

Private Sub cmdSelectFolder_Click()
Me.lblscanfoldertarget = getFolder & ""
lblscancount.Caption = ""
status1 " Press the scan button to scan the selected folder."
status2 ""
End Sub

Private Sub cmdStopScanningFolder_Click()
Me.DGSDirScan1.Cancel
End Sub

Private Sub cmdAddExt_Click()
If Len(txtext) > 4 Then
     MsgBox "Extension can only be 4 characters in length including the dot." & vbCrLf & "For longer extensions simply omit all but the last 4 characters", vbOKOnly, "Invalid Extension"
    txtext = Right(txtext, 4)
ElseIf Len(txtext) < 4 Then
    MsgBox "Extension must be 4 characters long." & vbCrLf & "The extension has been shortened for you.", vbOKOnly, "Invalid Extension"
    Exit Sub
End If
lstExt.AddItem txtext
txtext = ""
End Sub

Private Sub cmdSelectFile_Click()
On Error GoTo bail
Me.CommonDialog1.ShowOpen
If Me.CommonDialog1.FileName & "" <> "" Then
    ' check that file exist
    Dim fs As FileSystemObject
    Set fs = New FileSystemObject

    If fs.FileExists(CommonDialog1.FileName) Then
        status2 "Opening File " & CommonDialog1.FileName
        Me.RTF1.LoadFile CommonDialog1.FileName
        status1 "Press the scan button to scan this file."
        status2 ""
        lblscancount.Caption = ""
        lblscanfiletarget.Caption = CommonDialog1.FileName
    Else
    MsgBox CommonDialog1.FileName & " doesn't exist", vbCritical, "File not found"
    End If

End If
Exit Sub
bail:
MsgBox Err.Number & " " & Err.Description, vbCritical, "Unexpected error"
Err.Clear
End Sub

Private Sub cmdClearEmailList_Click()
ListView1.ListItems.Clear
End Sub

Private Sub cmdScanFile_Click()
scancount = 0
ExtractEmail lblscanfiletarget.Caption
    status1 "All Done =)"
    status2 ""
    Beep
End Sub

Private Sub ExtractEmail(path As String)
        On Error Resume Next ' or we could implicitly handle duplicates
        status2 "Extracting Email Adresses"
        Set Matches = regx1.Execute(RTF1.Text)    ' Execute search.
        For Each Match In Matches
          
            ' add to listview
            Me.ListView1.ListItems.Add , Match.Value, Match.Value
            Me.ListView1.ListItems(Match.Value).ListSubItems.Add , Match.Value, path
            ' update listcounter
            lbllistcount.Caption = ListView1.ListItems.Count & " emails in list"
        Next
        Set Matches = Nothing
        scancount = scancount + 1
        lblscancount.Caption = scancount & " files scanned"
End Sub

Private Sub Form_Load()
Set DGSDirScan1 = New DGSDirScanner
'regx for emails
    Set regx1 = New RegExp   ' Create Regular expresion to extract valid email addresses
    regx1.Pattern = "[a-zA-Z0-9-_.]+@[a-zA-Z0-9-_.]+\.[a-zA-Z0-9]+"   ' Set pattern.
    regx1.IgnoreCase = False   ' Set case insensitivity.
    regx1.Global = True        ' Set global applicability.
End Sub


Private Sub Form_Unload(Cancel As Integer)
Me.DGSDirScan1.Cancel
End Sub
Private Sub status2(msg As String)
picstatus2.Cls
picstatus2.Print msg
End Sub
Private Sub status1(msg As String)
picstatus1.Cls
picstatus1.Print msg
End Sub


Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
' sort column

' reverse sort order if current column is sortkey
If ListView1.SortKey = ColumnHeader.Index - 1 Then
    If ListView1.SortOrder = lvwAscending Then
        ListView1.SortOrder = lvwDescending
    Else
        ListView1.SortOrder = lvwAscending
    End If
Else ' set selected column as sort column
    ListView1.SortOrder = lvwAscending
    ListView1.SortKey = ColumnHeader.Index - 1
End If
ListView1.Sorted = True
End Sub
