VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmSupplier 
   Caption         =   "Suppliers"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13425
   Icon            =   "frmSupplier.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   13425
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   1
      Left            =   6240
      ScaleHeight     =   2715
      ScaleWidth      =   2865
      TabIndex        =   4
      Top             =   4710
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.PictureBox picTS 
      BorderStyle     =   0  'None
      Height          =   8955
      Index           =   0
      Left            =   5370
      ScaleHeight     =   8955
      ScaleWidth      =   7905
      TabIndex        =   3
      Top             =   450
      Width           =   7905
      Begin VB.Frame frmField 
         Caption         =   "Comment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Index           =   8
         Left            =   90
         TabIndex        =   21
         Top             =   6840
         Width           =   7725
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
            Height          =   1545
            Index           =   8
            Left            =   180
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   22
            Tag             =   "Comment"
            Top             =   270
            Width           =   7365
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "Business Fax"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   7
         Left            =   90
         TabIndex        =   19
         Top             =   6000
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            Index           =   7
            Left            =   180
            TabIndex        =   20
            Tag             =   "Fax"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "Business Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   6
         Left            =   90
         TabIndex        =   17
         Top             =   5160
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            Index           =   6
            Left            =   180
            TabIndex        =   18
            Tag             =   "Phone"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "Purchase Order Email Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   4320
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            Left            =   180
            TabIndex        =   16
            Tag             =   "PurchaseOrderEmail"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "Contact Email Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   3450
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            TabIndex        =   14
            Tag             =   "ContactEmail"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "Contact Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   2610
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            TabIndex        =   12
            Tag             =   "ContactName"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "ABN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   1770
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            TabIndex        =   10
            Tag             =   "ABN"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "ACN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   930
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            TabIndex        =   8
            Tag             =   "ACN"
            Top             =   270
            Width           =   7425
         End
      End
      Begin VB.Frame frmField 
         Caption         =   "Company Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Tag             =   "0"
         Top             =   90
         Width           =   7725
         Begin VB.TextBox txtField 
            Alignment       =   2  'Center
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
            Left            =   180
            TabIndex        =   6
            Tag             =   "CompanyName"
            Top             =   270
            Width           =   7425
         End
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   9405
      Left            =   5310
      TabIndex        =   2
      Top             =   120
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   16589
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Profile"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameSupplier 
      Caption         =   "Suppliers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   5055
      Begin VB.CommandButton cmdSupplier 
         Caption         =   "Create Supplier"
         Height          =   345
         Index           =   0
         Left            =   3210
         TabIndex        =   24
         Top             =   8880
         Width           =   1635
      End
      Begin VB.CommandButton cmdSupplier 
         Caption         =   "Save Current Supplier"
         Height          =   345
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   8880
         Width           =   1935
      End
      Begin MSComctlLib.ListView lvSupplier 
         Height          =   8535
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   4725
         _ExtentX        =   8334
         _ExtentY        =   15055
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         PictureAlignment=   5
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Company name"
            Object.Width           =   7937
         EndProperty
         Picture         =   "frmSupplier.frx":0442
      End
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSupplier_Click(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdSupplier_Click"
    Const ContainerName = "frmSupplier"
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


    Dim ix As Byte, SQL As String, sql1 As String
    
    Select Case Index
    Case 0
        For ix = txtField.LBound To txtField.UBound
            txtField(ix) = ""
        Next
        frmField(0).Tag = 0
    Case 1
        
        Select Case frmField(0).Tag
        Case 0
            SQL = "("
            sql1 = "("
            For ix = txtField.LBound To txtField.UBound
                SQL = SQL + txtField(ix).Tag
                sql1 = sql1 + "'" + MySQL.ESC(txtField(ix).Text) + "'"
                If ix < txtField.UBound Then
                    SQL = SQL + "', "
                    sql1 = sql1 + ", "
                End If
            Next
            SQL = SQL + ")"
            sql1 = sql1 + ")"
            MySQL.Execute ADOConn, "INSERT INTO supplier " & SQL & " VALUES " & sql1
            
            Dim rsOpen As ADODB.Recordset
            
            SQL = ""
            
            For ix = txtField.LBound To txtField.UBound
                SQL = SQL + txtField(ix).Tag + " = " + "'" + MySQL.ESC(txtField(ix).Text) + "'"
                If ix < txtField.UBound Then
                    SQL = SQL + " AND "
                End If
            Next
            
            
            If MySQL.OpenTable(ADOConn, rsOpen, , "select RecId from supplier where " & SQL) = True Then
                If rsOpen.RecordCount > 0 Then
                    
                    Call lvSupplier.ListItems.Add(, "s" & rsOpen!RecID, txtField(0).Text)
                    frmField(0).Tag = rsOpen!RecID
                End If
            End If
            
        Case Else
        
            SQL = ""
            
            For ix = txtField.LBound To txtField.UBound
                SQL = SQL + txtField(ix).Tag + "=" + "'" + MySQL.ESC(txtField(ix).Text) + "'"
                If ix < txtField.UBound Then
                    SQL = SQL + ","
                End If
            Next
            
            MySQL.Execute ADOConn, "UPDATE supplier SET " & SQL & " Where RecID = " & frmField(0).Tag
            
        End Select
    End Select
    
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
    Const ContainerName = "frmSupplier"
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



    Dim rsOpen As ADODB.Recordset
    
    
    If MySQL.OpenTable(ADOConn, rsOpen, , "select RecID, CompanyName from supplier") = True Then
        If rsOpen.RecordCount > 0 Then
            Dim itmX As ListItem
            
            While Not rsOpen.EOF And Err.Number = 0
                Set itmX = lvSupplier.ListItems.Add(, "s" & rsOpen!RecID, rsOpen!CompanyName)
                rsOpen.MoveNext
            Wend
       
        End If
    
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

Private Sub Form_Resize()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "Form_Resize"
    Const ContainerName = "frmSupplier"
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


    frameSupplier.Move frameSupplier.Left, frameSupplier.Top, frameSupplier.Width, Me.ScaleHeight - frameSupplier.Top * 2
    lvSupplier.Move lvSupplier.Left, lvSupplier.Top, lvSupplier.Width, frameSupplier.Height - lvSupplier.Top - cmdSupplier(0).Height - 60 * 2
    cmdSupplier(1).Move lvSupplier.Left, lvSupplier.Top + lvSupplier.Height + 60
    cmdSupplier(0).Move lvSupplier.Width - cmdSupplier(0).Width + lvSupplier.Left, lvSupplier.Top + lvSupplier.Height + 60
    
    If Me.ScaleWidth > 5400 Then ts.Move frameSupplier.Left * 2 + frameSupplier.Width, ts.Top, Me.ScaleWidth - frameSupplier.Left * 3 - frameSupplier.Width, Me.ScaleHeight - ts.Top * 2
    
    Call ts_Click
    
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

Private Sub lvSupplier_ItemClick(ByVal Item As MSComctlLib.ListItem)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "lvSupplier_ItemClick"
    Const ContainerName = "frmSupplier"
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



    Dim rsOpen As ADODB.Recordset
    
    
    If MySQL.OpenTable(ADOConn, rsOpen, , "select * from supplier where RecID = " & Mid(Item.Key, 2)) = True Then
        If rsOpen.RecordCount > 0 Then
            Dim ix As Byte
            For ix = txtField.LBound To txtField.UBound
                txtField(ix) = IIf(IsNull(rsOpen(txtField(ix).Tag)), "", rsOpen(txtField(ix).Tag))
            Next
            frmField(0).Tag = rsOpen!RecID
        End If
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

Private Sub picTS_Resize(Index As Integer)


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "picTS_Resize"
    Const ContainerName = "frmSupplier"
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


    Select Case Index
    Case 0
        Dim ix As Byte
        For ix = frmField.LBound To frmField.UBound
            If Me.Width > 6200 Then frmField(ix).Move frmField(ix).Left, frmField(ix).Top, picTS(Index).Width - frmField(ix).Left * 2
            If Me.Width > 6200 Then txtField(ix).Move txtField(ix).Left, txtField(ix).Top, frmField(ix).Width - txtField(ix).Left * 2
        Next ix
    End Select
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


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "ts_Click"
    Const ContainerName = "frmSupplier"
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


    Dim ix As Byte
    
    For ix = 1 To ts.Tabs.Count
        If ix <> ts.SelectedItem.Index Then
            picTS(ix - 1).Visible = False
        Else
            picTS(ix - 1).Move ts.clientLeft, ts.clientTop, ts.clientWidth, ts.clientHeight
            picTS(ix - 1).ZOrder 0
            picTS(ix - 1).Visible = True
        End If
    Next
    
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
