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



Function chkError(objErr As ErrObject) As Byte

    Static sErr As ErrObject
    Static LoopCount As Byte
    Dim objBuffer As New clsErrBuffer
        
    Select Case objErr.Number
    'Case 13 ' Type mismatch
    '    chkError = vbResumeNext
    Case 3420 ' Object invalid or no longer set.
        chkError = vbExit
        
    Case 35600 ' Index out of bounds
        chkError = vbExit
        
    Case -2147467259 ' [MySQL][ODBC 3.51 Driver][mysqld-3.23.55-log]Can't create/write to file '/var/tmp/#sql12c85_4873_1.MYD' (Errcode: 28)
        chkError = vbExit
        
    Case 3704 ' Operation is not allowed when the object is closed.
        chkError = vbExit
    
    Case 20 ' Resume without error
        chkError = vbExit
        
    Case Else
        If sErr Is Nothing Then
            Set sErr = objErr
            chkError = vbResume
            LoopCount = 0
        ElseIf sErr <> objErr Then
            Set sErr = objErr
            chkError = vbResume
            LoopCount = 0
        Else
            LoopCount = LoopCount + 1
        End If
        
        If LoopCount = 1 Then
            Select Case MsgBox(Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & "Abort will exit the subroutine (advised), Retry will resume the erroring line of code and Ignore will move on to the next line of code.", vbAbortRetryIgnore)
            Case vbIgnore
                chkError = vbResumeNext
                If Me.colError.FindError(objErr.Number) <> 0 Then
                    Set objBuffer = Me.colError(Me.colError.FindError(objErr.Number))
                    objBuffer.CaseStatement = vbTab & "Case " & objErr.Number & " ' " & objErr.Description & vbCrLf & vbTab & vbTab & "chkError = vbResumeNext"
                Else
                    Set objBuffer = Me.colError.Add("err_" & objErr.Number, vbTab & "Case " & objErr.Number & " ' " & objErr.Description & vbCrLf & vbTab & vbTab & "chkError = vbResumeNext")
                    objBuffer.ErrNumber = objErr.Number
                End If
            Case vbAbort
                chkError = vbExit
                If Me.colError.FindError(objErr.Number) <> 0 Then
                    Set objBuffer = Me.colError(Me.colError.FindError(objErr.Number))
                    objBuffer.CaseStatement = vbTab & "Case " & objErr.Number & " ' " & objErr.Description & vbCrLf & vbTab & vbTab & "chkError = vbExit"
                Else
                    Set objBuffer = Me.colError.Add("err_" & objErr.Number, vbTab & "Case " & objErr.Number & " ' " & objErr.Description & vbCrLf & vbTab & vbTab & "chkError = vbExit")
                    objBuffer.ErrNumber = objErr.Number
                End If
            Case vbRetry
                chkError = vbResume
                If Me.colError.FindError(objErr.Number) <> 0 Then
                    Set objBuffer = Me.colError(Me.colError.FindError(objErr.Number))
                    objBuffer.CaseStatement = vbTab & "Case " & objErr.Number & " ' " & objErr.Description & vbCrLf & vbTab & vbTab & "chkError = vbResume"
                Else
                    Set objBuffer = Me.colError.Add("err_" & objErr.Number, vbTab & "Case " & objErr.Number & " ' " & objErr.Description & vbCrLf & vbTab & vbTab & "chkError = vbResume")
                    objBuffer.ErrNumber = objErr.Number
                End If
            End Select
        End If
        'If bDebug = True Then Stop
    End Select
    
    If objErr.Number > 0 Then cDebug objErr.Description
    
End Function
