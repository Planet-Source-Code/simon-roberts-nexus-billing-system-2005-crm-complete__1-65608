VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"mapAccHoldings"
Attribute VB_Ext_KEY = "Member1" ,"mapCategory"
Attribute VB_Ext_KEY = "Member2" ,"mapSalesChan"
Attribute VB_Ext_KEY = "Member3" ,"mapSysops"
Attribute VB_Ext_KEY = "Member4" ,"mapTax"
Attribute VB_Ext_KEY = "Member5" ,"mapVirtualISP"


Public Function ColumnSort(ByVal ColumnHeader As MSComctlLib.ColumnHeader, ByVal lv As ListView) As Boolean

    lv.Sorted = True
    lv.SortKey = ColumnHeader.Index - 1
    If lv.SortOrder = lvwAscending Then
        lv.SortOrder = lvwDescending
    Else
        lv.SortOrder = lvwAscending
    End If
    
End Function


Public Function SaveColWidths(lv As ListView, frm As Form)

    Dim ix As Integer
    
    For ix = 1 To lv.ColumnHeaders.Count
        SaveSetting App.ProductName, frm.Name, lv.Name & "_Width_COL" & ix, lv.ColumnHeaders(ix).Width
        SaveSetting App.ProductName, frm.Name, lv.Name & "_Text_COL" & ix, lv.ColumnHeaders(ix).Text
    Next
       
End Function

Public Function LoadColWidths(lv As ListView, frm As Form)

    Dim ix As Integer
    
    For ix = 1 To lv.ColumnHeaders.Count
        lv.ColumnHeaders(ix).Width = GetSetting(App.ProductName, frm.Name, lv.Name & "_Width_COL" & ix, lv.ColumnHeaders(ix).Width)
        'lv.ColumnHeaders(ix).Text = GetSetting("projectalpha", frm.Name, lv.Name & "_Text_COL" & ix, lv.ColumnHeaders(ix).Text)
    Next
            
End Function

Public Function LV2CSV(Filename As String, ByVal lv As ListView, Header As String)

    Dim ioTXT As String
    Dim iY As Long
    Dim ix As Integer
    Dim iFree As Integer
    
    iFree = FreeFile
    On Error Resume Next
    
    Open Filename For Output As #iFree
        
    If Err.Number > 0 Then
        MsgBox "This filename could not be opened" & vbCrLf & vbCrLf & Err.Description, vbCritical, "Error Number " & Err.Number
        Exit Function
    End If
    
    Print #iFree, Header
    
    For ix = 1 To lv.ColumnHeaders.Count
        ioTXT = ioTXT + Chr$(34) + lv.ColumnHeaders(ix).Text + Chr$(34) + ","
    Next ix
        
    ioTXT = Left(ioTXT, Len(ioTXT) - 1)
    
    Print #iFree, ioTXT
    
    For iY = 1 To lv.ListItems.Count
        ioTXT = ""
        ioTXT = ioTXT & IIf(IsNumeric(lv.ListItems(iY).Text) = False, Chr$(34) + lv.ListItems(iY).Text + Chr$(34), lv.ListItems(iY).Text) + ","
        For ix = 2 To lv.ColumnHeaders.Count
            ioTXT = ioTXT & Chr$(34) + lv.ListItems(iY).SubItems(ix - 1) + Chr$(34) + ","
        Next ix
        ioTXT = Left(ioTXT, Len(ioTXT) - 1)
        Print #iFree, ioTXT
    Next
    
    Close #iFree
    
End Function
