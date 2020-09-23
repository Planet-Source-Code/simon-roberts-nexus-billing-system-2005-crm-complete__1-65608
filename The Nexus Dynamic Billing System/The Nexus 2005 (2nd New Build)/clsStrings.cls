VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Function ReplaceString(sMessage As String, sSearch As String, sReplace As String) As String

    Dim iPos As Variant
    
    Do
        iPos = 0
        If InStr(sMessage, sSearch) > 0 Then
            iPos = InStr(iPos + 1, sMessage, sSearch)
            If iPos = 1 Then
                sMessage = sReplace & Mid(sMessage, Len(sSearch) + 1)
            Else
                sMessage = Left(sMessage, iPos - 1) & sReplace & Mid(sMessage, iPos + Len(sSearch))
            End If
            While InStr(iPos + Len(sReplace), sMessage, sSearch) > 0 And Err.Number = 0
                iPos = InStr(iPos + Len(sReplace), sMessage, sSearch)
                If iPos = 1 Then
                    sMessage = sReplace & Mid(sMessage, Len(sSearch))
                Else
                    sMessage = Left(sMessage, iPos - 1) & sReplace & Mid(sMessage, iPos + Len(sSearch))
                End If
            Wend
        Else
            Exit Do
        End If
    Loop While sReplace = "" And Err.Number = 0
    ReplaceString = sMessage
    
End Function


