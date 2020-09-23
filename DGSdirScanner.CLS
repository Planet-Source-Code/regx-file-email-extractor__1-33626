VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DGSDirScanner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"walkdir"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
' DirScanner Class Module by RegX
' Copyright 2002 DGS
'=============================================
'Needs reference to:
'Microsoft VBScript Regular Expressions 5.5
'download at http://msdn.microsoft.com/downloads/default.asp?URL=/downloads/sample.asp?url=/msdn-files/027/001/733/msdncompositedoc.xml
'=============================================
'
'Usage
'Dim WithEvents ds As dirScanner
'Set ds = New dirScanner
'ds.Scan Dir1.Path

Public Event newdir(d As Folder)
Public Event newfile(f As File)
'local variable(s) to hold property value(s)
Private mvarScandepth As Long 'local copy
Private hault As Boolean
Public Property Let Scandepth(ByVal vData As Long)
Attribute Scandepth.VB_Description = "Sets the recursion depth for scan"
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Scandepth = 5
    mvarScandepth = vData
End Property

Public Property Get Scandepth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Scandepth
    Scandepth = mvarScandepth
End Property
Public Sub Cancel()
    hault = True
End Sub
Public Sub Scan(startdir)
' dim variables
Dim fs As New FileSystemObject
Dim fld1, fld2 As Folder
Dim fil As File
Dim tmpdepth As Long
hault = False

Dim dirs As New Collection ' create collection for dirs
dirs.Add startdir
Do While dirs.Count > 0 And hault = False
    ' get sub folders
    Set fld1 = fs.getFolder(dirs.Item(1))
    tmpdepth = tmpdepth + 1
    If Scandepth = 0 Or tmpdepth < Scandepth Then
        For Each fld2 In fld1.SubFolders
            If hault = True Then Exit Sub
            RaiseEvent newdir(fld2)
            dirs.Add fld2.path
            DoEvents
        Next
    End If
    ' get files
    For Each fil In fld1.Files
        If hault = True Then Exit Sub
        RaiseEvent newfile(fil)
            DoEvents
    Next
    dirs.Remove 1
Loop

Set fs = Nothing
Set dirs = Nothing
End Sub

