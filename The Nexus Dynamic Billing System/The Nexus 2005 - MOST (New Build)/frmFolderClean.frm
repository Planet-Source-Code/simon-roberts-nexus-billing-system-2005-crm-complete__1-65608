VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFolderClean 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Checking and Removing Files"
   ClientHeight    =   5910
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   Begin VB.Timer timKiller 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2700
      Top             =   2730
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   225
      Left            =   180
      TabIndex        =   2
      Top             =   5610
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Files to Remove"
      Height          =   2655
      Left            =   150
      TabIndex        =   1
      Top             =   2880
      Width           =   6285
      Begin MSComctlLib.ListView lvFiles 
         Height          =   2205
         Left            =   120
         TabIndex        =   4
         Top             =   270
         Width           =   6045
         _ExtentX        =   10663
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Name"
            Object.Width           =   10557
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Folders to Remove"
      Height          =   2655
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   6285
      Begin MSComctlLib.ListView lvFolders 
         Height          =   2205
         Left            =   150
         TabIndex        =   3
         Top             =   300
         Width           =   5985
         _ExtentX        =   10557
         _ExtentY        =   3889
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Folder Name"
            Object.Width           =   10557
         EndProperty
      End
   End
End
Attribute VB_Name = "frmFolderClean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Pattern As String
Public UNCPath As String
Public NewFileType As String
Public nFiles As Long

Public FormState As enumFormState


Public Function FindFilesAPI(ByVal Path As String, ByVal SearchStr As String, ByRef outFiles() As String, ByVal SubDirs As Boolean, Optional ByRef nFiles As Long = 0) As Long
    Dim hSearch As Long, WFD As WIN32_FIND_DATA
    Dim Result As Long, CurItem As String
    
    Path = NormalizeDir(Path)
    Screen.MousePointer = vbHourglass
    'Walk through this directory and get matching files.
    hSearch = FindFirstFile(Path & "*", WFD)
    If Not hSearch = INVALID_HANDLE_VALUE Then
        Result = True
        Do While Result
            CurItem = StripNulls(WFD.cFileName)
            If Not CurItem = "." And Not CurItem = ".." Then
                'Valid item.
                If SubDirs And (GetFileAttributes(Path & CurItem) And FILE_ATTRIBUTE_DIRECTORY) Then
                    'Item is a sub-directory, read it recursivly.
                    FindFilesAPI Path & CurItem, SearchStr, outFiles(), True, nFiles
                Else
                    'Item is a file which we're searching for.
                    ReDim Preserve outFiles(nFiles)
                    outFiles(nFiles) = Path & CurItem
                    nFiles = nFiles + 1
                End If
            End If
            'Get next item
            Result = FindNextFile(hSearch, WFD)
            gSleep
        Loop
        FindClose hSearch
    End If
    'Return the number of files in this directory (well, is also stored in the ByRef parameter nFiles).
    FindFilesAPI = nFiles
    
End Function

Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStr(OriginalStr, Chr$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function

Private Sub AddFiles(ByRef FileList() As String, ByRef FileCnt As Long, sPath As String, uPath As String)
    Dim i As Long
    Dim k As Long
    
    'Add the new files to the track list.
    If FileCnt = 0 Then Exit Sub
    
    Me.MousePointer = 11
    Me.Enabled = False
    Dim bFound As Boolean
    Dim sDir As String
    
    For i = 0 To FileCnt - 1
        With lvFiles.ListItems.Add(, , GetFile(FileList(i)))
            .Tag = FileList(i)
            sDir = GetDir(.Tag)
        End With
        bFound = False
        For k = lvFolders.ListItems.Count To 1 Step -1
            If LCase(lvFolders.ListItems(k).Tag) = LCase(sDir) Then
                bFound = True
                Exit For
            End If
        Next
        If bFound = False Then
            With lvFolders.ListItems.Add(, , GetSubDir(sDir))
                .Tag = sDir
            End With
        End If
    Next
        
    Me.Enabled = True
    Me.MousePointer = 0
'    lblStatus.Caption = "Ready."
End Sub
Private Sub Form_Load()

'    Me.Show
    On Error Resume Next
    
    FormState = Loading
    
    Screen.MousePointer = vbHourglass
    
    With lvFolders.ListItems.Add(, , GetSubDir(NormalizeDir(UNCPath)))
        .Tag = NormalizeDir(UNCPath)
    End With
    
    Dim fList() As String
    Dim nFiles As Long
        
    nFiles = FindFilesAPI(UNCPath, Pattern, fList(), True, 0)
    Call AddFiles(fList(), nFiles, UNCPath, UNCPath)
    gSleep
    Me.Refresh
    
    
    If lvFolders.ListItems.Count = 1 And lvFiles.ListItems.Count = 0 Then
        SetAttr NormalizeDir(UNCPath), vbNormal
        RmDir NormalizeDir(UNCPath)
        Unload Me
    End If
    
    FormState = Waiting
    
    Screen.MousePointer = vbDefault
    pb.Value = 0
    pb.Max = lvFiles.ListItems.Count + lvFolders.ListItems.Count
    timKiller.Enabled = True
    
    
End Sub

Private Sub timKiller_Timer()

    On Error Resume Next
    
    If lvFiles.ListItems.Count > 0 Then
        SetAttr lvFiles.ListItems(lvFiles.ListItems.Count).Tag, vbNormal
        Kill lvFiles.ListItems(lvFiles.ListItems.Count).Tag
        lvFiles.ListItems.Remove lvFiles.ListItems.Count
        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
    ElseIf lvFolders.ListItems.Count > 0 Then
        Dim i As Long
        Dim DirtoKill As Long
        Dim ilen As Long
        For i = lvFolders.ListItems.Count To 1 Step -1
            If Len(lvFolders.ListItems(lvFolders.ListItems.Count).Tag) > ilen Then
                ilen = Len(lvFolders.ListItems(lvFolders.ListItems.Count).Tag)
                DirtoKill = i
            End If
        Next
'        MsgBox lvFolders.ListItems(DirtoKill).Tag
        SetAttr lvFolders.ListItems(DirtoKill).Tag, vbNormal
        RmDir lvFolders.ListItems(DirtoKill).Tag
        lvFolders.ListItems.Remove DirtoKill
        PopulateResellers = PopulateResellers + 1: pb.Value = pb.Value + 1
    Else
        timKiller.Enabled = False
        timKiller.Interval = 0
        FormState = Finished
    End If
    pb.Refresh
    
End Sub
