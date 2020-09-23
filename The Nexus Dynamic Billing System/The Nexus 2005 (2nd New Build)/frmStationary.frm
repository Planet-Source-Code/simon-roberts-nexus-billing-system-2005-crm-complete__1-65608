VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStationary 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stationary"
   ClientHeight    =   8685
   ClientLeft      =   3045
   ClientTop       =   2625
   ClientWidth     =   14220
   Icon            =   "frmStationary.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cd 
      Left            =   6570
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Stationary"
      Height          =   8475
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4305
      Begin MSComctlLib.ListView lvPaper 
         Height          =   7575
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   13361
         View            =   3
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
            Text            =   "Code"
            Object.Width           =   6879
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create a New Stationary Item"
         Height          =   345
         Left            =   90
         TabIndex        =   1
         Top             =   7980
         Width           =   4005
      End
   End
   Begin VB.Frame framTS 
      BorderStyle     =   0  'None
      Caption         =   "HTML"
      Height          =   8025
      Index           =   0
      Left            =   4560
      TabIndex        =   4
      Top             =   510
      Width           =   9435
      Begin VB.Frame Frame4 
         Caption         =   "Document History SQL Statement (DO NOT ALTER)"
         Height          =   1245
         Index           =   3
         Left            =   150
         TabIndex        =   20
         Top             =   6210
         Width           =   9255
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   180
            TabIndex        =   24
            Tag             =   "HistorySQL"
            Top             =   270
            Width           =   6795
         End
         Begin VB.ComboBox cmbFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   7050
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Tag             =   "HistoryTable"
            Top             =   270
            Width           =   2085
         End
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   180
            MaxLength       =   100
            TabIndex        =   22
            Tag             =   "HistoryTitle"
            Top             =   720
            Width           =   6795
         End
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   7050
            MaxLength       =   100
            TabIndex        =   21
            Tag             =   "formcode"
            Top             =   720
            Width           =   2055
         End
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "&Preview"
         Height          =   285
         Left            =   2610
         TabIndex        =   19
         Top             =   7680
         Width           =   1995
      End
      Begin VB.TextBox HTML 
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   6
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3675
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   60
         Width           =   9255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load HTML from File"
         Height          =   285
         Left            =   7290
         TabIndex        =   16
         Top             =   7620
         Width           =   2085
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save Changes To Database"
         Height          =   285
         Left            =   90
         TabIndex        =   15
         Top             =   7680
         Width           =   2445
      End
      Begin VB.Frame Frame3 
         Caption         =   "Main Record that is Pointer"
         Height          =   795
         Left            =   150
         TabIndex        =   12
         Top             =   3780
         Width           =   5325
         Begin VB.ComboBox cmbFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   2700
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Tag             =   "Fieldname"
            Top             =   270
            Width           =   2475
         End
         Begin VB.ComboBox cmbFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Tag             =   "Tablename"
            Top             =   270
            Width           =   2475
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "DOC Identifier"
         Height          =   795
         Index           =   0
         Left            =   5520
         TabIndex        =   10
         Top             =   3780
         Width           =   3855
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   150
            MaxLength       =   64
            TabIndex        =   11
            Tag             =   "DOCShortname"
            Top             =   270
            Width           =   3585
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Listing Directory Search Engine SQL Statement (DO NOT ALTER)"
         Height          =   795
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   4590
         Width           =   9255
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   180
            MaxLength       =   255
            TabIndex        =   9
            Tag             =   "SelectStatement"
            Top             =   270
            Width           =   6795
         End
         Begin VB.ComboBox cmbFields 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   7050
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Tag             =   "PrimaryTable"
            Top             =   270
            Width           =   2085
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "OP Function Display String"
         Height          =   795
         Index           =   2
         Left            =   150
         TabIndex        =   5
         Top             =   5400
         Width           =   9255
         Begin VB.TextBox txtField 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   180
            MaxLength       =   255
            TabIndex        =   6
            Tag             =   "DisplayFormat"
            Top             =   270
            Width           =   8895
         End
      End
   End
   Begin VB.Frame framTS 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   2
      Left            =   7740
      TabIndex        =   18
      Top             =   4320
      Visible         =   0   'False
      Width           =   4365
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   8535
      Left            =   4500
      TabIndex        =   3
      Top             =   90
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   15055
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "HTML Code"
            Object.ToolTipText     =   "HTML code for the selected stationary"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmStationary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTables_Click()


End Sub

Private Sub cmbFields_Click(Index As Integer)

    Select Case Index
    Case 1
    
        If odb.colDBObjects.Count > 0 Then
    
        cmbFields(0).Clear
        
        Dim I As Long
        
        For I = 1 To odb.colDBObjects.Count
            If cmbFields(Index).List(cmbFields(Index).ListIndex) = odb.colDBObjects(I).Tablename Then
                cmbFields(0).AddItem odb.colDBObjects(I).FieldName
            End If
        Next I
        
        End If
    End Select
End Sub

Private Sub cmdPreview_Click()

    Dim iFile As Integer
    
    iFile = FreeFile
    
    If Dir(IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "tmpstationary.html", vbNormal) <> "" Then
        Kill IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "tmpstationary.html"
    End If
    
    Open IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "tmpstationary.html" For Output As #iFile
    Print #iFile, HTML.Text
    Close #iFile

    ShellExecute Me.hwnd, vbNullString, IIf(Right(App.path, 1) = "\", App.path, App.path + "\") & "tmpstationary.html", 0, "", -1
    
End Sub

Private Sub Command1_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command1_Click"
    Const ContainerName = "frmStationary"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    txt = InputBox("Enter in a new Stationary Code for the system, maximum of 24 Characters", "StationaryCode")
    
    If Trim(txt) = "" Then Exit Sub
    
    
    Call MySQL.Execute(ADOConn, "insert into stationary (RecID, StationaryCode) VALUES ('" & MySQL.GetTMPRecID("stationary", ADOConn) & "','" & MySQL.ESC(txt) & "')")
    
    Call Form_Load
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Command2_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command2_Click"
    Const ContainerName = "frmStationary"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    cd.Filter = "(*.HTM;*.HTML) HTML File|*.htm;*.html"
    cd.FilterIndex = 1
    cd.Filename = ""
    cd.ShowOpen
    
    If cd.Filename <> "" Then
    
        Open cd.Filename For Input As #32
        HTML = ""
        Do
            Line Input #32, templine
            HTML = HTML + templine + vbCrLf
        Loop Until EOF(32) Or Err.Number <> 0
    
        Close #32
        
    Else
        
    End If
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Command3_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Command3_Click"
    Const ContainerName = "frmStationary"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    If lvPaper.SelectedItem Is Nothing Then
    
    Else
    
        'On Error GoTo 0
        
        Call MySQL.Execute(ADOConn, "Update stationary set HTML = '" & MySQL.ESC(HTML) & "' where RecID = " & Mid(lvPaper.SelectedItem.Key, 2))
        
        Dim I As Long
        
        For I = cmbFields.LBound To cmbFields.UBound
            Call MySQL.Execute(ADOConn, "Update stationary set `" + cmbFields(I).Tag + "` = '" & MySQL.ESC(cmbFields(I).Text) & "' where RecID = " & Mid(lvPaper.SelectedItem.Key, 2))
        Next
        
        
        For I = txtField.LBound To txtField.UBound
            Call MySQL.Execute(ADOConn, "Update stationary set `" + txtField(I).Tag + "` = '" & MySQL.ESC(txtField(I).Text) & "' where RecID = " & Mid(lvPaper.SelectedItem.Key, 2))
        Next
    
    End If
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub Form_Load()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Load"
    Const ContainerName = "frmStationary"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    Dim rsLoad As ADODB.Recordset
    
    lvPaper.ListItems.Clear
    
    Call MySQL.OpenTable(ADOConn, rsLoad, , "select StationaryCode, RecID from stationary")
    
    If rsLoad.RecordCount > 0 Then
        
        Dim itmX As ListItem
        
        While Not rsLoad.EOF And Err.Number = 0
            Set itmX = lvPaper.ListItems.Add(, "p" & rsLoad!RecID, rsLoad!StationaryCode)
            rsLoad.MoveNext
        Wend
        
    End If
    
    
    
    If odb.colDBObjects.Count > 0 Then
    
        Dim I As Long
        
        For I = 1 To odb.colDBObjects.Count
            If cmbFields(1).ListCount = 0 Then
                cmbFields(1).AddItem odb.colDBObjects(I).Tablename
                cmbFields(2).AddItem odb.colDBObjects(I).Tablename
                cmbFields(3).AddItem odb.colDBObjects(I).Tablename
            Else
                If odb.colDBObjects(I).Tablename <> cmbFields(1).List(cmbFields(1).ListCount - 1) Then
                    cmbFields(1).AddItem odb.colDBObjects(I).Tablename
                    cmbFields(2).AddItem odb.colDBObjects(I).Tablename
                    cmbFields(3).AddItem odb.colDBObjects(I).Tablename
                End If
            End If
        Next I
        
    End If
    
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub lvPaper_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvPaper_ItemClick"
    Const ContainerName = "frmStationary"
    '***************************************************************************************************************


'
'***********************************************************************************************
'**  Project Alpha ® 2003, 2004 +                                                             **
'***********************************************************************************************
'**  This code is not to be distributed, reverse engineered or simulated in any way without   **
'**  Premission from the author. The authors of this code is as follows: Simon Antony Roberts **                                                     **
'**Jarrett Cliff Costi, these two are the only people you can communicate with about this code**
'***********************************************************************************************
'**  Project Alpha is a product of Exitstencil Press Australia                                **
'***********************************************************************************************
'**                                                                                           **
'**  Routine:                                                                                 **
'**  Arguments:                                                                               **
'**  Description:    Subroutine, Function or Property of project alpha                        **
'**  Author:         Simon Roberts                                                            **
'**  Date Last mod:  19-01-2004                                                               **
'**                                                                                           **
'********************************************** Copyright © 2004 Exitstencil Press Australia ***
'
'
'
    If bDebug = -1 Then
        On Error GoTo 0
    ElseIf bDebug = 1 Then
        On Error Resume Next
    Else
        On Error GoTo ErrorOccur
    End If


    Dim rsLoad As ADODB.Recordset
    Call MySQL.OpenTable(ADOConn, rsLoad, , "select * from stationary where RecID = " & Mid(Item.Key, 2))
    
    Dim I As Long
    
    HTML = ""
    cmbFields(0).ListIndex = -1
    cmbFields(1).ListIndex = -1
    cmbFields(2).ListIndex = -1
    cmbFields(0).Clear
        
    Dim t As Long
    
    If IsNull(rsLoad!HTML) Then
    
    Else
        HTML = rsLoad!HTML
        
        For t = cmbFields.UBound To cmbFields.LBound Step -1
            For I = 0 To cmbFields(t).ListCount - 1
                If cmbFields(t).List(I) = rsLoad(cmbFields(t).Tag) Then
                    cmbFields(t).ListIndex = I
                    If t = 1 Then Call cmbFields_Click(Val(t))
                    Exit For
                End If
            Next I
        Next t
        
        For I = txtField.LBound To txtField.UBound
            If Not IsNull(rsLoad(txtField(I).Tag)) Then
                txtField(I) = rsLoad(txtField(I).Tag)
            Else
                txtField(I) = ""
            End If
        Next I
    End If
            
Exit Sub



ErrorOccur:
Select Case oErr.chkError(Val(Err.Number), Err.Description, RoutineName, ContainerName)
Case vbResume
    Resume
Case vbExit
    
Case vbResumeNext
    Resume Next
End Select

End Sub

Private Sub ts_Click()

    Dim I As Long
    
    For I = framTS.LBound To framTS.UBound
        framTS(I).Move ts.clientLeft, ts.clientTop, ts.clientWidth, ts.clientHeight
        framTS(I).Visible = IIf(ts.SelectedItem.Index = I - 1, True, False)
        framTS(I).ZOrder 0
    Next
    
End Sub
