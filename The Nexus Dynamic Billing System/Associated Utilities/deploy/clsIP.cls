VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"


Public Function calcNumIPs(Ipfrom As String, ipTo As String) As Double

    On Error Resume Next
    
    Dim IPadd1 As IPs_Type
    Dim IPadd2 As IPs_Type
    
    Dim iPos As Byte
    Dim ioldPos As Byte
    
    ioldPos = 0
    iPos = InStr(ioldPos + 1, ipTo, ".")
    If iPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd2.First = Val(Left(ipTo, iPos - 1))
        ioldPos = iPos
    End If
    
    iPos = InStr(ioldPos + 1, ipTo, ".")
    If iPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd2.Second = Val(Mid(ipTo, ioldPos + 1, iPos - ioldPos - 1))
        ioldPos = iPos
    End If
    
    iPos = InStr(ioldPos + 1, ipTo, ".")
    If iPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd2.Third = Val(Mid(ipTo, ioldPos + 1, iPos - ioldPos - 1))
        ioldPos = iPos
    End If

    If ioldPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd2.Fourth = Val(Right(ipTo, Len(Ipfrom) - ioldPos))
        ioldPos = iPos
    End If
    
    ioldPos = 0
    iPos = InStr(ioldPos + 1, Ipfrom, ".")
    If iPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd1.First = Val(Left(Ipfrom, iPos - 1))
        ioldPos = iPos
    End If
    
    iPos = InStr(ioldPos + 1, Ipfrom, ".")
    If iPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd1.Second = Val(Mid(Ipfrom, ioldPos + 1, iPos - ioldPos - 1))
        ioldPos = iPos
    End If
    
    iPos = InStr(ioldPos + 1, Ipfrom, ".")
    If iPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd1.Third = Val(Mid(Ipfrom, ioldPos + 1, iPos - ioldPos - 1))
        ioldPos = iPos
    End If

    If ioldPos = 0 Then
        calcNumIPs = -1
        Exit Function
    Else
        IPadd1.Fourth = Val(Right(Ipfrom, Len(Ipfrom) - ioldPos))
        ioldPos = iPos
    End If
    
    Dim iCount As Double
    
    Do
        If IPadd1.First = IPadd2.First And IPadd1.Second = IPadd2.Second And IPadd1.Third = IPadd2.Third And IPadd1.Fourth = IPadd2.Fourth Then
            iCount = iCount + 1
            Exit Do
        Else
            IPadd1.Fourth = IPadd1.Fourth + 1
            iCount = iCount + 1
            If IPadd1.Fourth = 256 Then
                IPadd1.Third = IPadd1.Third + 1
                IPadd1.Fourth = 0
                If IPadd1.Third = 256 Then
                    IPadd1.Second = IPadd1.Second + 1
                    IPadd1.Third = 0
                    If IPadd1.Second = 256 Then
                        IPadd1.First = IPadd1.First + 1
                        IPadd1.Second = 0
                        If IPadd1.First = 256 Then
                            Exit Do
                        End If
                    End If
                End If
             End If
        End If
        DoEvents
    Loop
    
    calcNumIPs = iCount
    
End Function

