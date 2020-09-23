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
Private mvarmapVirtualISP As mapVirtualISP
Private mvarmapAccHoldings As mapAccHoldings
Private mvarmapCategory As mapCategory
Private mvarmapTax As mapTax
Private mvarmapSysops As mapSysops
Private mvarmapSalesChan As mapSalesChan


Public Property Get mapSalesChan() As mapSalesChan
    If mvarmapSalesChan Is Nothing Then
        Set mvarmapSalesChan = New mapSalesChan
    End If


    Set mapSalesChan = mvarmapSalesChan
End Property


Public Property Set mapSalesChan(vData As mapSalesChan)
    Set mvarmapSalesChan = vData
End Property



Public Property Get mapSysops() As mapSysops
    If mvarmapSysops Is Nothing Then
        Set mvarmapSysops = New mapSysops
    End If


    Set mapSysops = mvarmapSysops
End Property


Public Property Set mapSysops(vData As mapSysops)
    Set mvarmapSysops = vData
End Property



Public Property Get mapTax() As mapTax
    If mvarmapTax Is Nothing Then
        Set mvarmapTax = New mapTax
    End If


    Set mapTax = mvarmapTax
End Property


Public Property Set mapTax(vData As mapTax)
    Set mvarmapTax = vData
End Property



Public Property Get mapCategory() As mapCategory
    If mvarmapCategory Is Nothing Then
        Set mvarmapCategory = New mapCategory
    End If


    Set mapCategory = mvarmapCategory
End Property


Public Property Set mapCategory(vData As mapCategory)
    Set mvarmapCategory = vData
End Property



Public Property Get mapAccHoldings() As mapAccHoldings
    If mvarmapAccHoldings Is Nothing Then
        Set mvarmapAccHoldings = New mapAccHoldings
    End If


    Set mapAccHoldings = mvarmapAccHoldings
End Property


Public Property Set mapAccHoldings(vData As mapAccHoldings)
    Set mvarmapAccHoldings = vData
End Property



Public Property Get mapVirtualISP() As mapVirtualISP
    If mvarmapVirtualISP Is Nothing Then
        Set mvarmapVirtualISP = New mapVirtualISP
    End If


    Set mapVirtualISP = mvarmapVirtualISP
End Property


Public Property Set mapVirtualISP(vData As mapVirtualISP)
    Set mvarmapVirtualISP = vData
End Property
Private Sub Class_Terminate()
  Set mvarmapSalesChan = Nothing
  Set mvarmapSysops = Nothing
  Set mvarmapTax = Nothing
  Set mvarmapCategory = Nothing
  Set mvarmapAccHoldings = Nothing
    Set mvarmapVirtualISP = Nothing
End Sub



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
        
    If Err.Number <> 0 Then
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
