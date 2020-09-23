VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmProducts 
   Caption         =   "List of all products on the system for this vendor"
   ClientHeight    =   10875
   ClientLeft      =   1650
   ClientTop       =   1755
   ClientWidth     =   11280
   Icon            =   "frmProducts.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10875
   ScaleWidth      =   11280
   Begin MSComDlg.CommonDialog cd 
      Left            =   5400
      Top             =   5190
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Height          =   405
      Left            =   120
      TabIndex        =   1
      Top             =   10410
      Width           =   2175
   End
   Begin MSComctlLib.ListView lvProd 
      Height          =   10185
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   17965
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
End
Attribute VB_Name = "frmProducts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VendorID As Double

Private Sub cmdExport_Click()


    '*[ Error Checking Variables ]**********************************************************************************
    
    
    Const RoutineName = "cmdExport_Click"
    Const ContainerName = "frmProducts"
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


    If lvProd.ListItems.Count = 0 Then Exit Sub
    
    cd.Filter = "(*.CSV) Comma Seperate File|*.CSV"
    cd.FilterIndex = 1
    cd.Filename = ""
    cd.ShowSave
    
    If cd.Filename = "" Then Exit Sub
    
    If Dir(cd.Filename, vbNormal) <> "" Then
        Select Case MsgBox("The file that you have selected already exists, press Yes to append, No to Delete.", vbCritical + vbYesNoCancel)
        Case vbCancel
            Exit Sub
        Case vbNo
            Kill cd.Filename
        End Select
    End If
    
    Call GUI.LV2CSV(cd.Filename, lvProd, "")
    
    
    MsgBox "Done Exporting"
    Exit Sub
    
ErrExport:
    
    MsgBox "An error occured please try again later: " & Err.Description
    Err.Clear
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
    Const ContainerName = "frmProducts"
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


    Dim SQL As String
    
    SQL = "select plantemplates.RecID, plantemplates.VendorPartID, plantemplates.SubPartID, servicetypes.ServiceKey, servicetypes.Description as stDescription, plantemplates.BillImmediately, plantemplates.Description as PlanDescription, plantemplates.PeriodFee as CyclePeriodFee, plantemplates.MBPerPeriod as MBPerCycle, plantemplates.MBBlockSize, plantemplates.FeePerBlock, plantemplates.HoursPerPeriod, plantemplates.ExtraPerHour, plantemplates.SessionTimeout, plantemplates.IdleTimeout, plantemplates.SessionsAllowed, plantemplates.CostPrice, plantemplates.MBCostPrice as MBBlockCost, plantemplates.PeriodCostPrice as CostPerHour, plantemplates.Hidden, plantemplates.MBQuota from plantemplates, servicetypes where servicetypes.RecID = plantemplates.ServiceID and plantemplates.VendorID = '" & Me.VendorID & "' order by ServiceKey, PlanDescription"
    
    Dim rsload As ADODB.Recordset
    
    Call MySQL.OpenTable(ADOConn, rsload, , SQL)
    
    
    If rsload.State = adStateOpen Then
        lvProd.ListItems.Clear
        If rsload.RecordCount > 0 Then
            Dim itmX As ListItem
            
            lvProd.ColumnHeaders.Clear
            
            Dim ix As Long
            
            For ix = 0 To rsload.Fields.Count - 1
                lvProd.ColumnHeaders.Add , "r" & ix, rsload.Fields(ix).Name
            Next ix
            
            
            
            Do
                Set itmX = lvProd.ListItems.Add(, "r" & rsload!RecID, "")
                For ix = 0 To lvProd.ColumnHeaders.Count - 1
                    If ix = 0 Then
                        itmX.Text = IIf(IsNull(rsload(lvProd.ColumnHeaders(ix + 1).Text)), "", rsload(lvProd.ColumnHeaders(ix + 1).Text))
                    Else
                        itmX.SubItems(ix) = IIf(IsNull(rsload(lvProd.ColumnHeaders(ix + 1).Text)), "", rsload(lvProd.ColumnHeaders(ix + 1).Text))
                    End If
                    
                Next ix
                
                rsload.MoveNext
                
            Loop Until rsload.EOF
            
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
    Const ContainerName = "frmProducts"
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


    
    If Me.WindowState = vbMinimized Then Exit Sub
        cmdExport.Top = Me.ScaleHeight - cmdExport.Height - 120
    lvProd.Move 120, 120, Me.ScaleWidth - 240, cmdExport.Top - 240
    
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
