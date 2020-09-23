VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"colError"
Private mvarcolError As colError


Public Property Get colError() As colError
    If mvarcolError Is Nothing Then
        Set mvarcolError = New colError
    End If


    Set colError = mvarcolError
End Property


Public Property Set colError(vData As colError)
    Set mvarcolError = vData
End Property
Private Sub Class_Terminate()
    Set mvarcolError = Nothing
End Sub



Public Function chkError(errno As Long, errdesc As String, ByVal RoutineName As String, ByVal ContainerName As String) As Byte

    Static serr As Long
    Static LoopCount As Byte
    Dim objBuffer As New clsErrBuffer
        
    Select Case errno
    'Case 13 ' Type mismatch
    '    chkError = vbResumeNext
    Case 3420 ' Object invalid or no longer set.
        chkError = vbExit
        
    Case -2147217865 ' [MySQL][ODBC 3.51 Driver][mysqld-4.0.12-log]Table 'projectalpha.virtualISP' doesn't exist
        chkError = vbIgnore
        
    Case 91 ' Object variable or With block variable not set
        chkError = vbIgnore
        
    Case 35600 ' Index out of bounds
        chkError = vbExit
        
    Case -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-3.23.55-log]Can't create/write to file '/var/tmp/#sql12c85_4873_1.MYD' (Errcode: 28)
        chkError = vbExit
        
    Case 3704 ' Operation is not allowed when the object is closed.
        chkError = vbExit
    
    Case 20 ' Resume without error
        chkError = vbExit
        
    Case 35602 ' Key is not unique in collection
        chkError = vbIgnore
        
    Case 3021 ' Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.
        chkError = vbIgnore
        
    Case Else
        If serr = 0 Then
            serr = errno
            chkError = vbResume
            LoopCount = 0
        ElseIf serr <> errno Then
            serr = errno
            chkError = vbResume
            LoopCount = 0
        Else
            LoopCount = LoopCount + 1
        End If
        
        If LoopCount = 1 Then
                
            Dim txtErr As String
            
            txtErr = "Congradulation you have just trapped an error in the sub kernelling. But don't panic you can submit your persepective on what caused it when you have finished using this program because it will send an email back to the development team of any minor flaws in the code that exist."
            txtErr = txtErr + vbCrLf & vbCrLf
            txtErr = txtErr + "Error Number: " & errno
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "Description: " & errdesc
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "Result Code: " & Hex(Rnd * 255 * 255 * 255)
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "Routine Name: " & RoutineName
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "Container Name: " & ContainerName
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "You now currently have 3 options to choose from that is all you have to do:"
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "[Abort] will leave the subroutine and if it run again and the same error occurs, it will then also leave the subrotine..."
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "[Retry] will retry the subroutine and if it run again and the same error occurs."
            txtErr = txtErr + vbCrLf
            txtErr = txtErr + "[Ignore] will not bother with the error except log it in the memory for the email later, this is like the other two options..."
            
            Select Case MsgBox(txtErr, vbAbortRetryIgnore + critial, "[" & errno & "] - Error Sytax: " & Err.Description)
            Case vbIgnore
                chkError = vbResumeNext
                If Me.colError.FindError(errno) <> 0 Then
                    Set objBuffer = Me.colError(Me.colError.FindError(errno))
                    objBuffer.CaseStatement = vbTab & "Case " & errno & " ' " & errdesc & vbCrLf & vbTab & vbTab & "chkError = vbResumeNext"
                Else
                    Set objBuffer = Me.colError.Add("err_" & errno, vbTab & "Case " & errno & " ' " & errdesc & vbCrLf & vbTab & vbTab & "chkError = vbResumeNext", , RoutineName, ContainerName)
                    objBuffer.ErrNumber = errno
                End If
            Case vbAbort
                chkError = vbExit
                If Me.colError.FindError(errno) <> 0 Then
                    Set objBuffer = Me.colError(Me.colError.FindError(errno))
                    objBuffer.CaseStatement = vbTab & "Case " & errno & " ' " & errdesc & vbCrLf & vbTab & vbTab & "chkError = vbExit"
                Else
                    Set objBuffer = Me.colError.Add("err_" & errno, vbTab & "Case " & errno & " ' " & errdesc & vbCrLf & vbTab & vbTab & "chkError = vbExit", , RoutineName, ContainerName)
                    objBuffer.ErrNumber = errno
                End If
            Case vbRetry
                chkError = vbResume
                If Me.colError.FindError(errno) <> 0 Then
                    Set objBuffer = Me.colError(Me.colError.FindError(errno))
                    objBuffer.CaseStatement = vbTab & "Case " & errno & " ' " & errdesc & vbCrLf & vbTab & vbTab & "chkError = vbResume"
                Else
                    Set objBuffer = Me.colError.Add("err_" & errno, vbTab & "Case " & errno & " ' " & errdesc & vbCrLf & vbTab & vbTab & "chkError = vbResume", , RoutineName, ContainerName)
                    objBuffer.ErrNumber = errno
                End If
            End Select
            
            objBuffer.DateTime = sysNOW
            'objBuffer.LBL = frmDebug.txtDebug.Tag
        End If
        'If bDebug = True Then Stop
    End Select
    
    If errno > 0 Then cDebug errdesc
'    frmMDIMain.bRefresh = True
    
End Function
